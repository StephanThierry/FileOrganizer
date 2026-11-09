using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OpenCvSharp;

namespace FileOrganizer;

/// <summary>Title and comment tags of a media file.</summary>
public static class MediaMetadata
{
    /// <param name="TagCorrupt">The file's header is damaged (a hint only: the video OCR decides whether the video is unreadable).</param>
    public readonly record struct Tags(string Title, string Comment, bool TagCorrupt);

    public static Tags Read(string fullPath)
    {
        try
        {
            // TagLib parses the headers of the different file types
            using var file = TagLib.File.Create(fullPath);
            return new Tags(file.Tag.Title ?? "", file.Tag.Comment ?? "", false);
        }
        catch (TagLib.CorruptFileException)
        {
            return new Tags("", "", true);
        }
        catch (Exception)
        {
            // Not a media file (a .txt, for example), or unreadable: no tags
            return new Tags("", "", false);
        }
    }
}

/// <summary>The OCR results already made, in the scanned folder (the same file the command line always made).</summary>
public static class OcrCache
{
    public const string FileName = "_FileProcessor_VideoOcrResult.json";

    public static string PathIn(string folder) => Path.Combine(folder, FileName);

    public static Dictionary<string, VideoOcrResult> Load(string folder)
    {
        try
        {
            string path = PathIn(folder);
            if (File.Exists(path))
            {
                var loaded = JsonSerializer.Deserialize<Dictionary<string, VideoOcrResult>>(File.ReadAllText(path));
                if (loaded != null) return new Dictionary<string, VideoOcrResult>(loaded, StringComparer.OrdinalIgnoreCase);
            }
        }
        catch (Exception)
        {
            // A broken cache is just a cold cache
        }
        return new Dictionary<string, VideoOcrResult>(StringComparer.OrdinalIgnoreCase);
    }

    public static void Save(string folder, Dictionary<string, VideoOcrResult> cache)
    {
        try
        {
            string path = PathIn(folder), tmp = path + ".tmp";
            File.WriteAllText(tmp, JsonSerializer.Serialize(cache, new JsonSerializerOptions { WriteIndented = true }));
            File.Move(tmp, path, overwrite: true);
        }
        catch (Exception)
        {
            // Read-only folder: the next scan simply reads the video again
        }
    }
}

/// <summary>One frame of a video has been read.</summary>
public sealed record OcrFrameInfo(int Frame, int Frames, double Seconds, IReadOnlyList<string> NewLines);

/// <summary>Reads the text in a video. A seam so the engine can be tested without OpenCV and the OCR models.</summary>
public interface IVideoOcr
{
    /// <param name="isMatch">Checked for every line; reading stops after the first frame with a match (when it returns true).</param>
    VideoOcrResult Read(string path, Func<string, bool> isMatch, Action<OcrFrameInfo>? onFrame, CancellationToken ct);
}

/// <summary>Frames out of a video with OpenCV, text out of the frames with PP-OCRv5 (VideoTextReader).</summary>
public sealed class OpenCvVideoOcr : IVideoOcr, IDisposable
{
    private readonly string modelsFolder;
    private VideoTextReader? reader;   // created on the first video, reused for all frames and videos

    public OpenCvVideoOcr(string modelsFolder)
    {
        this.modelsFolder = modelsFolder;
        // Quiet: OpenCV is chatty about codecs it can't handle (must be set before it loads)
        Environment.SetEnvironmentVariable("OPENCV_FFMPEG_LOGLEVEL", "-8");
        Environment.SetEnvironmentVariable("OPENCV_LOG_LEVEL", "OFF");
        Environment.SetEnvironmentVariable("OPENCV_VIDEOIO_DEBUG", "0");
    }

    public VideoOcrResult Read(string path, Func<string, bool> isMatch, Action<OcrFrameInfo>? onFrame, CancellationToken ct)
    {
        var result = new VideoOcrResult();

        // Kept outside the try below, so a model problem stops the run instead of marking every video as corrupt
        reader ??= new VideoTextReader(modelsFolder);

        string? tempCopy = null;
        try
        {
            using var capture = OpenVideo(path, ct, out tempCopy);
            if (!capture.IsOpened()) { result.IsCorrupt = true; return result; }

            double fps = capture.Fps;
            double frameCount = capture.Get(VideoCaptureProperties.FrameCount);
            double duration = frameCount / fps;
            if (fps <= 0 || frameCount <= 0 || double.IsNaN(duration)) { result.IsCorrupt = true; return result; }

            var ordered = SampleTimes(duration);

            int validFramesRead = 0, n = 0;
            foreach (double t in ordered)
            {
                ct.ThrowIfCancellationRequested();
                n++;
                capture.Set(VideoCaptureProperties.PosMsec, t * 1000);
                using var frame = new Mat();
                if (!capture.Read(frame) || frame.Empty()) { onFrame?.Invoke(new OcrFrameInfo(n, ordered.Count, t, Array.Empty<string>())); continue; }
                validFramesRead++;

                // No pre-processing: PP-OCR works best on the original colour frame
                var lines = new List<string>();
                bool frameMatched = false;
                foreach (var line in reader.ReadLines(frame).Where(l => l.Length > 2))
                {
                    if (IsMostlyJunk(line)) continue;
                    result.Lines.Add(line);
                    lines.Add(line);
                    if (isMatch(line)) frameMatched = true;
                }
                onFrame?.Invoke(new OcrFrameInfo(n, ordered.Count, t, lines));
                if (frameMatched) break;
            }

            if (validFramesRead == 0 && frameCount > 0) result.IsCorrupt = true;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception)
        {
            result.IsCorrupt = true;
        }
        finally
        {
            if (tempCopy != null) { try { File.Delete(tempCopy); } catch (Exception) { } }
        }
        return result;
    }

    /// <summary>
    /// Opens a video. OpenCV's ffmpeg on Windows cannot open a path with letters outside plain ASCII (a Cyrillic or accented
    /// letter, an emoji): it says the file can't be opened although the file is fine. For such a file a copy with a plain name is
    /// made in the temp folder and opened instead (<paramref name="tempCopy"/>: delete it afterwards).
    /// </summary>
    private static VideoCapture OpenVideo(string path, CancellationToken ct, out string? tempCopy)
    {
        tempCopy = null;
        var capture = new VideoCapture(path, VideoCaptureAPIs.FFMPEG);
        if (capture.IsOpened() || !OperatingSystem.IsWindows() || path.All(c => c < 128)) return capture;

        tempCopy = CopyToPlainName(path, ct);
        if (tempCopy == null) return capture;   // no room for a copy: the video stays unreadable
        capture.Dispose();
        return new VideoCapture(tempCopy, VideoCaptureAPIs.FFMPEG);
    }

    /// <summary>A copy of the file in the temp folder under a plain ASCII name (null: there is no room for it, or it can't be copied).</summary>
    private static string? CopyToPlainName(string path, CancellationToken ct)
    {
        string? target = null;
        try
        {
            var info = new FileInfo(path);
            string temp = Path.GetTempPath();
            var drive = new DriveInfo(Path.GetPathRoot(Path.GetFullPath(temp))!);
            if (drive.AvailableFreeSpace < info.Length + 512L * 1024 * 1024) return null;   // keep half a GB free

            string ext = info.Extension;
            if (ext.Length == 0 || ext.Any(c => c >= 128)) ext = ".video";
            target = Path.Combine(temp, "FileOrganizer_" + Guid.NewGuid().ToString("N") + ext);
            using (var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 20))
            using (var destination = new FileStream(target, FileMode.CreateNew, FileAccess.Write, FileShare.None, 1 << 20))
            {
                var buffer = new byte[1 << 20];
                int read;
                while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ct.ThrowIfCancellationRequested();
                    destination.Write(buffer, 0, read);
                }
            }
            return target;
        }
        catch (OperationCanceledException)
        {
            if (target != null) { try { File.Delete(target); } catch (Exception) { } }
            throw;
        }
        catch (Exception)
        {
            if (target != null) { try { File.Delete(target); } catch (Exception) { } }
            return null;
        }
    }

    /// <summary>
    /// The moments of a video to read: the very first frame (titles and logos often show only at the start), then one frame a second
    /// for the first 15 seconds, then 10 more spread over the rest.
    /// </summary>
    public static List<double> SampleTimes(double duration)
    {
        var times = new List<double> { 0 };
        for (int i = 1; i <= 15; i++) { if (i <= duration) times.Add(i); }
        if (duration > 15)
        {
            double chunk = (duration - 15) / 11;
            for (int i = 1; i <= 10; i++) times.Add(15 + chunk * i);
        }
        return times.Distinct().OrderBy(x => x).ToList();
    }

    // Lines that are mostly symbols are noise
    private static bool IsMostlyJunk(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return true;
        int alphaNumeric = s.Count(char.IsLetterOrDigit);
        return (double)alphaNumeric / s.Length < 0.3;
    }

    public void Dispose() => reader?.Dispose();
}

/// <summary>One way a piece of text was read from a video (capitals do not count as a difference), and how many times.</summary>
public sealed record OcrVariant(string Text, int Count);

/// <summary>
/// One piece of text on screen as the OCR read it in different frames: the readings differ a little (P-P-P.TV, Ppp.tV, p-p-p.tv).
/// <paramref name="Best"/> is the most complete reading; <paramref name="Variants"/> are all the different readings, the most seen first.
/// </summary>
public sealed record OcrFinding(string Best, IReadOnlyList<OcrVariant> Variants);

/// <summary>Tidying the text found in a video before it is shown (matching still uses every line that was read).</summary>
public static class OcrText
{
    private sealed class Group
    {
        public string Key = "";
        public readonly List<string> Lines = new();
    }

    /// <summary>
    /// The pieces of text found in a video, each with every different reading of it. Readings that are the same apart from case,
    /// spacing and punctuation belong together, and so does a reading that is only a piece of a longer one (cut off). The pieces
    /// are in the order they were first seen. Text without a single letter or digit is left out.
    /// </summary>
    public static List<OcrFinding> Findings(IEnumerable<string> lines)
    {
        var groups = new List<Group>();
        foreach (var raw in lines)
        {
            string line = (raw ?? "").Trim();
            string key = Key(line);
            if (key.Length == 0) continue;

            var into = groups.FirstOrDefault(g => g.Key.Contains(key, StringComparison.Ordinal));   // the same text, or a piece of a longer one
            if (into != null) { into.Lines.Add(line); continue; }

            // A longer reading takes over the shorter pieces of itself, at the place of the first of them
            var pieces = groups.Where(g => key.Contains(g.Key, StringComparison.Ordinal)).ToList();
            var group = new Group { Key = key };
            int at = groups.Count;
            if (pieces.Count > 0)
            {
                at = groups.IndexOf(pieces[0]);
                foreach (var piece in pieces) group.Lines.AddRange(piece.Lines);
                groups.RemoveAll(g => pieces.Contains(g));
            }
            group.Lines.Add(line);
            groups.Insert(Math.Min(at, groups.Count), group);
        }

        return groups.Select(g =>
        {
            // Different readings: the same text in other capitals is one reading, shown as it was seen most often
            var variants = g.Lines.Select((text, i) => (text, i)).GroupBy(x => x.text.ToLowerInvariant())
                .Select(x => (Text: x.GroupBy(y => y.text).OrderByDescending(y => y.Count()).ThenBy(y => y.First().i).First().Key,
                              Count: x.Count(), First: x.Min(y => y.i), KeyLength: Key(x.Key).Length))
                .ToList();
            // The most complete reading: the most letters and digits, then the most often seen, then the one with most punctuation
            var best = variants.OrderByDescending(v => v.KeyLength).ThenByDescending(v => v.Count).ThenByDescending(v => v.Text.Length).ThenBy(v => v.First).First();
            var ordered = variants.OrderByDescending(v => v.Count).ThenByDescending(v => v.KeyLength).ThenByDescending(v => v.Text.Length).ThenBy(v => v.First)
                .Select(v => new OcrVariant(v.Text, v.Count)).ToList();
            return new OcrFinding(best.Text, ordered);
        }).ToList();
    }

    private static string Key(string line) => new string(line.Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());

    /// <summary>
    /// The lines without repeats: the same text apart from case, spacing and punctuation counts as one, and so does a line that is
    /// only a piece of a longer one (the same text read in another frame, or cut off). The longest wording is kept, in first-seen order.
    /// </summary>
    public static List<string> Distinct(IEnumerable<string> lines)
    {
        var kept = new List<(string Line, string Key)>();
        foreach (var line in lines)
        {
            string key = Key(line);
            if (key.Length == 0) continue;
            if (kept.Any(k => k.Key.Contains(key, StringComparison.Ordinal))) continue;      // already there, or part of a longer line
            kept.RemoveAll(k => key.Contains(k.Key, StringComparison.Ordinal));               // this one replaces shorter pieces of itself
            kept.Add((line, key));
        }
        return kept.Select(k => k.Line).ToList();
    }
}

/// <summary>Small pictures of videos for the interfaces, kept in the data folder.</summary>
public static class Thumbnails
{
    public const int Width = 240;

    /// <summary>How many pictures Review shows for a video.</summary>
    public const int Parts = 3;

    /// <summary>
    /// One of <paramref name="parts"/> pictures spread over a video (part 0 of 3 is a sixth in, part 1 the middle, part 2 five sixths in:
    /// away from the often black start and end). Made on first use; null when it can't be made.
    /// </summary>
    public static string? EnsurePart(string videoPath, string folder, int part, int parts = Parts)
    {
        try
        {
            if (parts < 1 || part < 0 || part >= parts) return null;
            var info = new FileInfo(videoPath);
            if (!info.Exists) return null;
            string target = Path.Combine(folder, KeyOf(info) + $"_{part + 1}of{parts}.png");
            if (File.Exists(target)) return target;

            using var capture = new VideoCapture(videoPath, VideoCaptureAPIs.FFMPEG);
            if (!capture.IsOpened()) return null;
            double fps = capture.Fps, frames = capture.Get(VideoCaptureProperties.FrameCount);
            double duration = fps > 0 ? frames / fps : 0;

            using var frame = new Mat();
            double at = duration > 0 ? duration * (part + 0.5) / parts : 0;
            capture.Set(VideoCaptureProperties.PosMsec, at * 1000);
            if (!capture.Read(frame) || frame.Empty())
            {
                capture.Set(VideoCaptureProperties.PosMsec, 0);
                if (!capture.Read(frame) || frame.Empty()) return null;
            }
            return Save(frame, folder, target);
        }
        catch (Exception)
        {
            return null;
        }
    }

    private static string KeyOf(FileInfo info) =>
        Convert.ToHexString(SHA1.HashData(Encoding.UTF8.GetBytes($"{info.FullName}|{info.Length}|{info.LastWriteTimeUtc.Ticks}"))).ToLowerInvariant();

    private static string Save(Mat frame, string folder, string target)
    {
        using var small = new Mat();
        int height = Math.Max(1, (int)Math.Round(frame.Height * (double)Width / frame.Width));
        Cv2.Resize(frame, small, new Size(Width, height));
        Directory.CreateDirectory(folder);
        // Encoded in memory and written with .NET: OpenCV's own writer can't take non-ASCII folder names on Windows
        Cv2.ImEncode(".png", small, out byte[] png);
        File.WriteAllBytes(target, png);
        return target;
    }

    /// <summary>The picture file for a video (made on first use), or null when it can't be made.</summary>
    public static string? Ensure(string videoPath, string folder)
    {
        try
        {
            var info = new FileInfo(videoPath);
            if (!info.Exists) return null;
            string target = Path.Combine(folder, KeyOf(info) + ".png");
            if (File.Exists(target)) return target;

            using var capture = new VideoCapture(videoPath, VideoCaptureAPIs.FFMPEG);
            if (!capture.IsOpened()) return null;
            double fps = capture.Fps, frames = capture.Get(VideoCaptureProperties.FrameCount);
            double duration = fps > 0 ? frames / fps : 0;

            using var frame = new Mat();
            // A frame a little way in (the very start is often black), else the first one
            double at = duration > 4 ? Math.Min(duration * 0.15, 30) : 0;
            capture.Set(VideoCaptureProperties.PosMsec, at * 1000);
            if (!capture.Read(frame) || frame.Empty())
            {
                capture.Set(VideoCaptureProperties.PosMsec, 0);
                if (!capture.Read(frame) || frame.Empty()) return null;
            }
            return Save(frame, folder, target);
        }
        catch (Exception)
        {
            return null;
        }
    }
}
