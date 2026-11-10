using System.Diagnostics;

namespace FileOrganizer;

public enum OrganizePhase { Idle, Listing, Scanning, Done, Cancelled, Failed }

/// <summary>A snapshot of a running scan, reported to the interfaces.</summary>
public sealed record OrganizeProgress
{
    public OrganizePhase Phase { get; init; }
    public string Message { get; init; } = "";
    public string Folder { get; init; } = "";

    public int FilesTotal { get; init; }
    public int FilesDone { get; init; }
    public int Matched { get; init; }
    public int Unmatched { get; init; }
    /// <summary>Videos that could not be read (they go to _CHECK_).</summary>
    public int Corrupt { get; init; }
    /// <summary>Videos whose OCR was already known from an earlier scan.</summary>
    public int CacheHits { get; init; }
    /// <summary>Videos that were read with OCR in this scan.</summary>
    public int OcrDone { get; init; }

    public string CurrentFile { get; init; } = "";
    /// <summary>What is being checked for the current file: "Filename", "Tags" or "Reading video".</summary>
    public string Stage { get; init; } = "";
    public int OcrFrame { get; init; }
    public int OcrFrames { get; init; }
    /// <summary>The last text lines found in video frames.</summary>
    public IReadOnlyList<string> RecentLines { get; init; } = Array.Empty<string>();
    /// <summary>A picture of the last video that was read (a file in the thumbnail folder).</summary>
    public string? LastThumbnail { get; init; }
    /// <summary>The video that picture is of (its full path), so it can be played.</summary>
    public string? LastThumbnailVideo { get; init; }

    public TimeSpan Elapsed { get; init; }
    public TimeSpan? Remaining { get; init; }

    public double Fraction => FilesTotal > 0 ? Math.Min(FilesDone / (double)FilesTotal, 1) : 0;
}

/// <summary>
/// Looks at the files in a folder and decides, for each, which category it belongs to: by its name, its title and comment
/// tags, and (videos) the text found in frames by OCR. Nothing is moved: the result is shown, changed by the user,
/// and then turned into scripts or applied (ScriptWriter, MoveExecutor).
/// </summary>
public sealed class OrganizeEngine : IDisposable
{
    public const string CorruptCategory = "_CHECK_";

    /// <summary>The script, report and cache files the program makes itself: never treated as files to organize.</summary>
    public static readonly HashSet<string> OwnFiles = new(StringComparer.OrdinalIgnoreCase)
    {
        "settings.json", "unmatched_files.txt", OcrCache.FileName,
        "move_files.bat", "rename_files.bat", "move_files.ps1", "rename_files.ps1", "move_files.sh", "rename_files.sh",
    };

    private readonly Func<IVideoOcr> ocrFactory;
    private IVideoOcr? ocr;
    private readonly object sync = new();
    private CancellationTokenSource? cts;

    public OrganizeProgress Current { get; private set; } = new() { Phase = OrganizePhase.Idle };
    public bool IsRunning { get; private set; }

    /// <summary>Where pictures of the videos that were read go (null: none are made).</summary>
    public string? ThumbFolder { get; set; }

    /// <summary>Raised on every progress update (from a background thread).</summary>
    public event Action<OrganizeProgress>? ProgressChanged;
    /// <summary>Raised when a scan has finished or was stopped, with what it found so far.</summary>
    public event Action<OrganizeResult>? Completed;
    /// <summary>A line of what is going on, as the command line prints it.</summary>
    public event Action<string>? Log;

    public OrganizeEngine(Func<IVideoOcr> ocrFactory) => this.ocrFactory = ocrFactory;

    public void Dispose() => (ocr as IDisposable)?.Dispose();

    public void Cancel() => cts?.Cancel();

    /// <summary>Starts a scan in the background. Returns false if one is already running.</summary>
    public bool Start(string folder, Settings settings)
    {
        lock (sync)
        {
            if (IsRunning) return false;
            IsRunning = true;
            cts = new CancellationTokenSource();
        }
        var token = cts.Token;
        var snapshot = settings.Clone();
        Task.Run(() =>
        {
            try
            {
                var result = Run(folder, snapshot, token);
                Completed?.Invoke(result);
            }
            catch (Exception e)
            {
                Report(new OrganizeProgress { Phase = OrganizePhase.Failed, Message = e.Message, Folder = folder });
            }
            finally { lock (sync) IsRunning = false; }
        });
        return true;
    }

    private void Report(OrganizeProgress p)
    {
        Current = p;
        ProgressChanged?.Invoke(p);
    }

    private void Say(string line) => Log?.Invoke(line);

    /// <summary>The scan itself (blocking). On cancel it returns what it found so far, with the phase Cancelled.</summary>
    public OrganizeResult Run(string folder, Settings settings, CancellationToken ct)
    {
        if (!Directory.Exists(folder)) throw new DirectoryNotFoundException($"The folder '{folder}' does not exist.");
        var timer = Stopwatch.StartNew();
        var result = new OrganizeResult { Folder = folder };
        var videoExtensions = new HashSet<string>(settings.Opt.VideoExtensions ?? Options.DefaultVideoExtensions, StringComparer.OrdinalIgnoreCase);

        Report(new OrganizeProgress { Phase = OrganizePhase.Listing, Message = "Listing the files", Folder = folder });
        var files = ListEntries(folder, settings);
        var cache = OcrCache.Load(folder);

        int matched = 0, unmatched = 0, corrupt = 0, cacheHits = 0, ocrDone = 0, done = 0;
        double ocrSeconds = 0;
        int videosSeen = 0;
        string? lastThumb = null, lastThumbVideo = null;
        var recent = new List<string>();
        // (The files inside opened folders are not read by OCR: the folder's name says where they go)
        int remainingUncachedVideos = files.Count(f => f.FolderName == null && videoExtensions.Contains(Path.GetExtension(f.Name)) && !cache.ContainsKey(f.Relative));

        OrganizeProgress Snapshot(string file, string stage, int frame = 0, int frames = 0, OrganizePhase phase = OrganizePhase.Scanning)
        {
            TimeSpan? remaining = null;
            // Time left: the videos still to read, at the pace so far (not every video needs reading: filename and tag matches skip it)
            if (ocrDone >= 3 && videosSeen > 0)
            {
                double needsOcr = ocrDone / (double)videosSeen;
                remaining = TimeSpan.FromSeconds(ocrSeconds / ocrDone * Math.Max(remainingUncachedVideos, 0) * needsOcr);
            }
            return new OrganizeProgress
            {
                Phase = phase, Folder = folder, FilesTotal = files.Count, FilesDone = done,
                Matched = matched, Unmatched = unmatched, Corrupt = corrupt, CacheHits = cacheHits, OcrDone = ocrDone,
                CurrentFile = file, Stage = stage, OcrFrame = frame, OcrFrames = frames,
                RecentLines = recent.ToList(), LastThumbnail = lastThumb, LastThumbnailVideo = lastThumbVideo,
                Elapsed = timer.Elapsed, Remaining = remaining,
            };
        }

        bool cancelled = false;
        try
        {
            foreach (var entry in files)
            {
                ct.ThrowIfCancellationRequested();
                string file = entry.Name;
                string key = entry.Relative;   // what is shown, and the key in the OCR cache (the name for a file directly in the folder)
                string full = Path.Combine(folder, key);
                var d = new FileDecision { FileName = file, SubFolder = entry.SubFolder, FolderName = entry.FolderName, IsVideo = videoExtensions.Contains(Path.GetExtension(file)) };
                if (d.IsVideo) videosSeen++;

                // 1. The file name
                Report(Snapshot(key, "Filename"));
                if (!Decide(d, settings))
                {
                    // 2. The tags (title, comment), and a new name from them
                    Report(Snapshot(key, "Tags"));
                    var tags = MediaMetadata.Read(full);
                    d.Title = tags.Title;
                    d.Comment = tags.Comment;
                    string? newBaseName = !string.IsNullOrWhiteSpace(tags.Title) ? tags.Title : !string.IsNullOrWhiteSpace(tags.Comment) ? tags.Comment : null;
                    if (newBaseName != null)
                    {
                        string safeName = FolderNames.SafeFileName(newBaseName) + Path.GetExtension(file);
                        if (!string.Equals(file, safeName, StringComparison.OrdinalIgnoreCase))
                        {
                            d.RenameTo = safeName;
                            d.RenameSource = !string.IsNullOrWhiteSpace(tags.Title) ? "Title" : "Comment";
                        }
                    }
                    if (!string.IsNullOrWhiteSpace(tags.Title)) Say($"* Title Found '{tags.Title}' for file: {key}");
                    if (!string.IsNullOrWhiteSpace(tags.Comment)) Say($"* Comment Found '{tags.Comment}' for file: {key}");

                    if (!Decide(d, settings) && d.IsVideo)
                    {
                        // 3. The deep scan: text in the video's frames
                        VideoOcrResult read;
                        // A video that was once found unreadable is tried again: it may have been a false alarm that is fixed since
                        if (cache.TryGetValue(key, out var cached) && !cached.IsCorrupt)
                        {
                            Say($"[Cache Hit] Using existing OCR data for: {key}");
                            read = cached;
                            d.FromCache = true;
                            cacheHits++;
                        }
                        else
                        {
                            Say($"[Deep Scan] {key}...");
                            ocr ??= ocrFactory();
                            var started = Stopwatch.StartNew();
                            Func<string, bool> stopAt = line => settings.Opt.StopOcrAtFirstMatch && RuleMatcher.Match(line, settings) != null;
                            read = ocr.Read(full, stopAt, frame =>
                            {
                                foreach (var line in frame.NewLines)
                                {
                                    Say($"  OCR detection: {line}");
                                    recent.Add(line);
                                }
                                while (recent.Count > 12) recent.RemoveAt(0);
                                Report(Snapshot(key, "Reading video", frame.Frame, frame.Frames));
                            }, ct);
                            cache[key] = read;
                            OcrCache.Save(folder, cache);
                            ocrDone++;
                            remainingUncachedVideos--;
                            ocrSeconds += started.Elapsed.TotalSeconds;
                            if (ThumbFolder != null && Thumbnails.Ensure(full, ThumbFolder) is { } made) { lastThumb = made; lastThumbVideo = full; }
                        }
                        d.OcrLines = read.Lines.ToList();
                        d.OcrCorrupt = read.IsCorrupt;
                        Decide(d, settings);
                    }
                }

                result.Decisions.Add(d);
                done++;
                if (d.Source == MatchSource.Corrupt) corrupt++;
                if (d.IsMatched) { matched++; Say($"[Match] {key} -> {d.Category} ({d.Trigger})"); }
                else { unmatched++; Say($"[Unmatched] {key}"); }
                Report(Snapshot(key, ""));
            }
        }
        catch (OperationCanceledException)
        {
            cancelled = true;
        }

        result.Elapsed = timer.Elapsed;
        result.FinishedUtc = DateTime.UtcNow;
        Say($"\nProcessing Complete. Matches: {matched} | Unmatched: {unmatched}");
        Report(Snapshot("", "", phase: cancelled ? OrganizePhase.Cancelled : OrganizePhase.Done) with
        {
            Message = cancelled ? $"Stopped after {done:N0} of {files.Count:N0} files." : $"Looked at {files.Count:N0} files."
        });
        return result;
    }

    /// <summary>A file to look at: where it is inside the scanned folder, and the folder that was opened for it (null: it is directly in the scanned folder).</summary>
    private sealed record ScanEntry(string Name, string SubFolder, string? FolderName)
    {
        public string Relative => SubFolder.Length == 0 ? Name : Path.Combine(SubFolder, Name);
    }

    /// <summary>
    /// The files directly in the folder and, with Process folders on, the files (at any depth) of each folder directly inside it whose name
    /// fits a category. A category's own folder is never opened: what is in it has been organized already.
    /// </summary>
    private static List<ScanEntry> ListEntries(string folder, Settings settings)
    {
        var list = Directory.GetFiles(folder)
            .Select(Path.GetFileName).OfType<string>()
            .Where(f => !OwnFiles.Contains(f))
            .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
            .Select(f => new ScanEntry(f, "", null))
            .ToList();
        if (!settings.Opt.ProcessFolders) return list;

        var categoryFolders = new HashSet<string>(settings.Categories.Select(c => c.Name).Append(CorruptCategory), StringComparer.OrdinalIgnoreCase);
        var options = new EnumerationOptions { RecurseSubdirectories = true, IgnoreInaccessible = true, AttributesToSkip = 0, ReturnSpecialDirectories = false };
        foreach (var dir in Directory.GetDirectories(folder).OrderBy(d => d, StringComparer.OrdinalIgnoreCase))
        {
            string name = Path.GetFileName(dir);
            if (categoryFolders.Contains(name)) continue;               // a category's own folder: organized already
            if (RuleMatcher.Match(name, settings) == null) continue;    // its name fits no category
            try
            {
                foreach (var relative in Directory.EnumerateFiles(dir, "*", options).Select(f => Path.GetRelativePath(folder, f))
                             .Where(r => !OwnFiles.Contains(Path.GetFileName(r))).OrderBy(r => r, StringComparer.OrdinalIgnoreCase))
                    list.Add(new ScanEntry(Path.GetFileName(relative), Path.GetDirectoryName(relative) ?? "", name));
            }
            catch (Exception) { /* an unreadable folder is left out */ }
        }
        return list;
    }

    /// <summary>
    /// Decides a file from what is known about it, in this order: the file name, the title tag, the comment tag, a video that
    /// can't be read (_CHECK_), the text found in the video. Facts that have not been read yet are simply empty, so the same
    /// call is used after each step. Returns true when the file got a category.
    /// </summary>
    public static bool Decide(FileDecision d, Settings settings)
    {
        d.Category = null; d.Source = MatchSource.None; d.Trigger = "";

        var m = RuleMatcher.Match(d.FileName, settings);
        if (m != null) return Set(d, m.Category, MatchSource.Filename, m.Trigger);

        if (RuleMatcher.Match(d.Title, settings) is { } t) return Set(d, t.Category, MatchSource.Title, $"Title Tag: {t.Trigger}");
        if (RuleMatcher.Match(d.Comment, settings) is { } c) return Set(d, c.Category, MatchSource.Comment, $"Comment Tag: {c.Trigger}");

        // In a folder that was opened (Process folders): a file with no match of its own goes with the folder, by the folder's name. So the text
        // of a video in it is not read: its folder already says where it belongs.
        if (d.FolderName != null && RuleMatcher.Match(d.FolderName, settings) is { } f)
            return Set(d, f.Category, MatchSource.Folder, $"Folder: {d.FolderName} ({f.Trigger})");

        if (d.OcrCorrupt) return Set(d, CorruptCategory, MatchSource.Corrupt, "Hard Error: Video stream unreadable.");

        // The text read from the video, and the rules with several terms ("All of"), which look at everything known about the file
        var texts = new List<string> { d.FileName, d.Title, d.Comment };
        texts.AddRange(d.OcrLines);
        if (RuleMatcher.MatchVideo(d.OcrLines, texts, settings) is { } o)
            return Set(d, o.Category, o.IsAll && d.OcrLines.Count == 0 ? MatchSource.Filename : MatchSource.Ocr,
                o.IsAll && d.OcrLines.Count == 0 ? o.Trigger : $"OCR Match: '{o.Trigger}'");

        return false;
    }

    private static bool Set(FileDecision d, string category, MatchSource source, string trigger)
    {
        d.Category = category; d.Source = source; d.Trigger = trigger;
        return true;
    }

    /// <summary>
    /// Decides again from the facts already read (after the rules changed), without touching the files: instant, also for
    /// videos, because their OCR text is kept. The user's own choices (another category, leave it) stay. Files that were
    /// matched by name were never opened, so a scan is needed for tags or OCR they would now need.
    /// </summary>
    public static void Reevaluate(OrganizeResult result, Settings settings)
    {
        foreach (var d in result.Decisions)
            if (!d.Done) Decide(d, settings);
    }
}
