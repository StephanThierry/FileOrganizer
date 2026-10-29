using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OpenCvSharp;

public class FileProcessor : IDisposable
{
    private readonly string _targetFolder;
    private readonly string _settingsPath;
    private readonly string _batPath;
    private readonly string _unmatchedPath;
    private readonly string _ocrCachePath;
    private readonly string _renameScriptPath;
    private VideoTextReader? _textReader; // created on first video, reused for all frames
    private const string PossiblyCorruptCategory = "_CHECK_";

    // Windows gets .bat scripts, Linux/macOS get .sh scripts
    private static readonly bool IsWindows = OperatingSystem.IsWindows();
    private static readonly string ScriptExtension = IsWindows ? ".bat" : ".sh";

    private readonly HashSet<string> _videoExtensions = new(StringComparer.OrdinalIgnoreCase) 
        { ".mp4", ".mpeg", ".mpv", ".flv", ".mkv", ".avi", ".mov", ".wmv" };

    private readonly HashSet<string> _ignoredFiles;

    public FileProcessor(string targetFolder, string settingsPath)
    {
        _targetFolder = targetFolder;
        _settingsPath = settingsPath;
        _batPath = Path.Combine(targetFolder, "move_files" + ScriptExtension);
        _renameScriptPath = Path.Combine(targetFolder, "rename_files" + ScriptExtension);
        _ocrCachePath = Path.Combine(targetFolder, "_FileProcessor_VideoOcrResult.json");
        _unmatchedPath = Path.Combine(targetFolder, "unmatched_files.txt");

        //Console.SetError(TextWriter.Null);
        Cv2.SetLogLevel(0);

        _ignoredFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "settings.json", "unmatched_files.txt",
            "move_files.bat", "rename_files.bat", "move_files.sh", "rename_files.sh",
            Path.GetFileName(_ocrCachePath),
            AppDomain.CurrentDomain.FriendlyName,          // Linux/macOS executable has no extension
            AppDomain.CurrentDomain.FriendlyName + ".exe",
            AppDomain.CurrentDomain.FriendlyName + ".dll"
        };
    }

    public void Dispose() => _textReader?.Dispose();

    public void Run()
    {
        var settings = LoadSettings();
        if (settings?.Categories == null) return;

        var ocrCache = LoadOcrCache();

        var allFiles = Directory.GetFiles(_targetFolder)
                                .Select(Path.GetFileName)
                                .Where(f => f != null && !_ignoredFiles.Contains(f))
                                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                                .ToList();

        var moveActions = new List<MoveAction>();
        var renameActions = new List<RenameAction>();
        var unmatchedData = new List<UnmatchedFileInfo>();

        foreach (var file in allFiles)
        {
         
            // 1. Filename Match
            if (TryMatch(file!, settings, out var action))
            {
                moveActions.Add(action!);
                Console.WriteLine($"[Match] {file} -> {action.CategoryName} ({action.RuleTriggered})");
                continue;
            }

            // 2. Metadata Match (Cross-Platform)
            var (title, comments) = GetMediaMetadata(Path.Combine(_targetFolder, file!));

            // --- RENAME LOGIC ---
            string extension = Path.GetExtension(file);
            string? newBaseName = !string.IsNullOrWhiteSpace(title) ? title : 
                                    (!string.IsNullOrWhiteSpace(comments) ? comments : null);

            if (newBaseName != null)
            {
                string safeName = GetSafeFileName(newBaseName) + extension;
                if (!string.Equals(file, safeName, StringComparison.OrdinalIgnoreCase))
                {
                    renameActions.Add(new RenameAction(file!, safeName, !string.IsNullOrWhiteSpace(title) ? "Title" : "Comment"));
                }
            }               

            if (!string.IsNullOrWhiteSpace(title)) {
                Console.WriteLine($"* Title Found '{title}' for file: {file}") ;
                if(TryMatch(title, settings, out var titleAction))
                {
                    moveActions.Add(new MoveAction(file!, titleAction!.CategoryName, $"Title Tag: {titleAction.RuleTriggered}"));
                    Console.WriteLine($"[Match: Tag] {file} -> {titleAction.CategoryName} (Title: \"{title}\")");
                    continue;
                }
            }

            if (!string.IsNullOrWhiteSpace(comments)){
                Console.WriteLine($"* Comment Found '{comments}' for file: {file}") ;
                if (TryMatch(comments, settings, out var commentAction))
                {
                    moveActions.Add(new MoveAction(file!, commentAction!.CategoryName, $"Comment Tag: {commentAction.RuleTriggered}"));
                    Console.WriteLine($"[Match: Tag] {file} -> {commentAction.CategoryName} (Comment: \"{comments}\")");
                    continue;
                }
            }

            // 3. Video Deep Scan
            if (_videoExtensions.Contains(Path.GetExtension(file)))
            {
                VideoOcrResult ocrResult;

                // CHECK CACHE: Skip OCR if we already have data for this filename
                if (ocrCache.TryGetValue(file!, out var cachedResult))
                {
                    Console.WriteLine($"[Cache Hit] Using existing OCR data for: {file}");
                    ocrResult = cachedResult;
                }
                else
                {
                    Console.WriteLine($"[Deep Scan] {file}...");
                    ocrResult = GetVideoOcrText(Path.Combine(_targetFolder, file!));
                    // UPDATE CACHE: Store new result
                    ocrCache[file!] = ocrResult;
                    SaveOcrCache(ocrCache);
                }

                if (ocrResult.IsCorrupt)
                {
                    moveActions.Add(new MoveAction(file!, PossiblyCorruptCategory, "Hard Error: Video stream unreadable."));
                    continue;
                }

                bool found = false;
                foreach (var line in ocrResult.Lines)
                {
                    if (TryMatch(line, settings, out var ocrAction))
                    {
                        string ocrTrigger = $"OCR Match: '{ocrAction.RuleTriggered}'";
                        moveActions.Add(new MoveAction(file!, ocrAction!.CategoryName, ocrTrigger));
                        Console.WriteLine($"{ocrTrigger} -> {ocrAction!.CategoryName} \n");
                        found = true;
                        break;
                    }
                }
                if (found) continue;

                // For unmatched files, we store the OCR lines for the unmatched report
                unmatchedData.Add(new UnmatchedFileInfo(file!, ocrResult.Lines.ToList(), title ?? "", comments ?? ""));
            }
            else
            {
                // Non-video unmatched
                unmatchedData.Add(new UnmatchedFileInfo(file!, new List<string>(), title ?? "", comments ?? ""));
            }

            // No Match (already added to unmatchedData above, with any detected OCR text)
           Console.WriteLine($"[Unmatched] {file}");
           
        }

        

        GenerateBatFile(moveActions);
        GenerateRenameBatFile(renameActions);
        GenerateUnmatchedFile(unmatchedData);

        Console.WriteLine($"\nProcessing Complete. Matches: {moveActions.Count} | Unmatched: {unmatchedData.Count}");
    }
private Dictionary<string, VideoOcrResult> LoadOcrCache()
    {
        try
        {
            if (File.Exists(_ocrCachePath))
            {
                string json = File.ReadAllText(_ocrCachePath);
                return JsonSerializer.Deserialize<Dictionary<string, VideoOcrResult>>(json) 
                       ?? new Dictionary<string, VideoOcrResult>(StringComparer.OrdinalIgnoreCase);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Could not load OCR cache: {ex.Message}");
        }
        return new Dictionary<string, VideoOcrResult>(StringComparer.OrdinalIgnoreCase);
    }

    private void SaveOcrCache(Dictionary<string, VideoOcrResult> cache)
    {
        try
        {
            var options = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(cache, options);
            File.WriteAllText(_ocrCachePath, json);
            Console.WriteLine($"  [Cache] Saved OCR data for {cache.Count} files to {_ocrCachePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving OCR cache: {ex.Message}");
        }
    }
    private bool TryMatch(string input, Settings settings, out MoveAction? action)
    {
        action = null;
        foreach (var category in settings.Categories)
        {
            string searchName = category.Name.Replace(".com", "", StringComparison.OrdinalIgnoreCase);
            if (!category.IgnoreCategoryInContent && input.Contains(searchName, StringComparison.OrdinalIgnoreCase))
            {
                action = new MoveAction(input, category.Name, "Category Name Match");
                return true;
            }
            foreach (var rule in category.Rules)
            {
                if (rule.Type.Equals("Regex", StringComparison.OrdinalIgnoreCase) && Regex.IsMatch(input, rule.Value, RegexOptions.IgnoreCase))
                {
                    action = new MoveAction(input, category.Name, $"Regex: {rule.Value}");
                    return true;
                }
                if (rule.Type.Equals("Contains", StringComparison.OrdinalIgnoreCase) && input.Contains(rule.Value, StringComparison.OrdinalIgnoreCase))
                {
                    action = new MoveAction(input, category.Name, $"Contains: {rule.Value}");
                    return true;
                }
            }
        }
        return false;
    }

    private VideoOcrResult GetVideoOcrText(string path)
{
    var result = new VideoOcrResult();

    // Load the OCR models once (first video only), then reuse them for every frame.
    // Kept outside the try below, so a model problem stops the run instead of
    // marking every video as corrupt.
    _textReader ??= new VideoTextReader(Path.Combine(AppContext.BaseDirectory, "models"));

    try
    {
        using var capture = new VideoCapture(path, VideoCaptureAPIs.FFMPEG);
        if (!capture.IsOpened()) { result.IsCorrupt = true; return result; }

        double fps = capture.Fps;
        double frameCount = capture.Get(VideoCaptureProperties.FrameCount);
        double duration = frameCount / fps;

        if (fps <= 0 || frameCount <= 0 || double.IsNaN(duration)) { result.IsCorrupt = true; return result; }

        var ts = new List<double>();
        for (int i = 1; i <= 15; i++) { if (i <= duration) ts.Add(i); }
        if (duration > 15)
        {
            double chunk = (duration - 15) / 11;
            for (int i = 1; i <= 10; i++) ts.Add(15 + (chunk * i));
        }

        int validFramesRead = 0;
        foreach (double t in ts.Distinct().OrderBy(x => x))
        {
            capture.Set(VideoCaptureProperties.PosMsec, t * 1000);
            using Mat frame = new Mat();
            if (!capture.Read(frame) || frame.Empty()) continue;

            validFramesRead++;

            // No pre-processing: PP-OCR works best on the original colour frame.
            // (The old grayscale/upscale/bilateral/CLAHE steps were tuned for Windows OCR,
            // and the bilateral filter on a 2x upscaled frame was a large part of the time per frame.)
            // Low-resolution frames are upscaled automatically by the text detector.
            foreach (var line in _textReader.ReadLines(frame).Where(l => l.Length > 2))
            {
                // Basic junk filter (ignores lines that are just symbols)
                if (IsMostlyJunk(line)) continue;

                Console.WriteLine($"  OCR detection: {line}");
                result.Lines.Add(line);
            }
        }

        if (validFramesRead == 0 && frameCount > 0) result.IsCorrupt = true;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error: {ex.Message}");
        result.IsCorrupt = true; 
    }
    return result;
}

// Helper to filter out the noise you're seeing
private bool IsMostlyJunk(string s)
{
    if (string.IsNullOrWhiteSpace(s)) return true;
    // If the line has no letters or numbers, it's probably noise symbols
    int alphaNumericCount = s.Count(c => char.IsLetterOrDigit(c));
    return (double)alphaNumericCount / s.Length < 0.3; 
}

    private void GenerateBatFile(List<MoveAction> actions)
    {
        // Use UTF8 encoding with a Byte Order Mark (BOM) so Windows recognizes it; no BOM for sh
        var encoding = new System.Text.UTF8Encoding(IsWindows);

        using var writer = new StreamWriter(_batPath, false, encoding);
        if (!IsWindows) writer.NewLine = "\n";

        WriteScriptHeader(writer);

        // Create directories
        foreach (var cat in actions.Select(a => a.CategoryName).Distinct())
        {
            writer.WriteLine(IsWindows
                ? $"IF NOT EXIST {BatQuote(cat)} mkdir {BatQuote(cat)}"
                : $"mkdir -p -- {ShQuote(cat)}");
        }

        writer.WriteLine();
        writer.WriteLine(IsWindows ? "REM --- Moving Files ---" : "# --- Moving Files ---");
        foreach (var act in actions)
        {
            if (IsWindows)
            {
                writer.WriteLine($"move {BatQuote(act.FileName)} \"{act.CategoryName.Replace("%", "%%")}\\\"");
                writer.WriteLine($"REM Triggered: {act.RuleTriggered}");
            }
            else
            {
                writer.WriteLine($"mv -n -- {ShQuote(act.FileName)} {ShQuote(act.CategoryName + "/")}");
                writer.WriteLine($"# Triggered: {act.RuleTriggered.Replace('\n', ' ')}");
            }
        }

        MakeExecutable(_batPath);
    }

    // --- Script helpers (Windows .bat vs Linux/macOS .sh) ---

    private static void WriteScriptHeader(StreamWriter writer)
    {
        if (IsWindows)
        {
            // 1. Set Code Page to UTF-8 (65001) to support special chars like ’ and –
            // 2. Redirect to nul so the user doesn't see the "Active code page" message
            writer.WriteLine("@echo off");
            writer.WriteLine("chcp 65001 > nul");
            writer.WriteLine("cd /d \"%~dp0\"");
        }
        else
        {
            writer.WriteLine("#!/bin/sh");
            writer.WriteLine("cd \"$(dirname \"$0\")\" || exit 1");
        }
        writer.WriteLine();
    }

    // ESCAPE PERCENT SIGNS: % must be written as %% in batch files
    private static string BatQuote(string s) => "\"" + s.Replace("%", "%%") + "\"";

    // Single quotes keep everything literal in sh; a ' inside is written as '\''
    private static string ShQuote(string s) => "'" + s.Replace("'", "'\\''") + "'";

    private static void MakeExecutable(string path)
    {
        if (OperatingSystem.IsWindows()) return;
        try
        {
            File.SetUnixFileMode(path, File.GetUnixFileMode(path)
                | UnixFileMode.UserExecute | UnixFileMode.GroupExecute | UnixFileMode.OtherExecute);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: could not mark {path} as executable: {ex.Message}");
        }
    }

    private void GenerateUnmatchedFile(List<UnmatchedFileInfo> unmatched)
    {
        if (!unmatched.Any()) return;

        using var writer = new StreamWriter(_unmatchedPath);
        foreach (var item in unmatched)
        {
            writer.WriteLine(item.FileName + "...");

            // Write Metadata if it exists
            if (!string.IsNullOrWhiteSpace(item.Title))    writer.WriteLine($"    Metadata Title: \"{item.Title}\"");
            if (!string.IsNullOrWhiteSpace(item.Comment))  writer.WriteLine($"    Metadata Comment: \"{item.Comment}\"");

            foreach (var ocrLine in item.OcrLines)
            {
                writer.WriteLine($"    OCR Detected Text: \"{ocrLine}\"");
            }
            writer.WriteLine(); // Add empty line between files for readability
        }
    }
    private string GetSafeFileName(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return "unnamed";
        
        // Characters illegal in Windows filenames: \ / : * ? " < > |
        char[] illegalChars = Path.GetInvalidFileNameChars();
        string safe = name;
        foreach (char c in illegalChars)
        {
            safe = safe.Replace(c, '_');
        }
        
        // Trim spaces and periods from end (Windows requirement)
        return safe.Trim().TrimEnd('.');
    }
private void GenerateRenameBatFile(List<RenameAction> actions)
{
    if (!actions.Any()) return;

    var encoding = new System.Text.UTF8Encoding(IsWindows);

    using (var writer = new StreamWriter(_renameScriptPath, false, encoding))
    {
        if (!IsWindows) writer.NewLine = "\n";

        WriteScriptHeader(writer);
        writer.WriteLine(IsWindows
            ? "REM === Auto-generated Rename Script (Title/Comment) ==="
            : "# === Auto-generated Rename Script (Title/Comment) ===");
        writer.WriteLine();

        foreach (var action in actions)
        {
            if (IsWindows)
            {
                writer.WriteLine($"ren {BatQuote(action.OldFileName)} {BatQuote(action.NewFileName)}");
                writer.WriteLine($"REM Source: {action.SourceField}\n");
            }
            else
            {
                writer.WriteLine($"mv -n -- {ShQuote(action.OldFileName)} {ShQuote(action.NewFileName)}");
                writer.WriteLine($"# Source: {action.SourceField}\n");
            }
        }
    }

    MakeExecutable(_renameScriptPath);
    Console.WriteLine($"Generated '{Path.GetFileName(_renameScriptPath)}' with {actions.Count} commands.");
}    
    private Settings? LoadSettings()
    {
        var settings = JsonSerializer.Deserialize<Settings>(File.ReadAllText(_settingsPath), new JsonSerializerOptions { PropertyNameCaseInsensitive = true }); 
        return settings;
    }

    private (string Title, string Comments) GetMediaMetadata(string FullFilePath)
    {
        try
        {
            // TagLib handles the heavy lifting of parsing different file headers
            using var tfile = TagLib.File.Create(FullFilePath);
            
            string title = tfile.Tag.Title ?? "";
            string comments = tfile.Tag.Comment ?? "";

            return (title, comments);
        }
        catch(TagLib.UnsupportedFormatException unsupported)
        {
            // If the file format isn't supported (like a .txt file) 
            // or the header is corrupt, just return empty strings.
            Console.WriteLine($"Warning: Unsupported format for '{FullFilePath}'. Exception: {unsupported.Message}");
            return ("", "");
        } catch(TagLib.CorruptFileException corrupt)
        {
            // If the file format isn't supported (like a .txt file) 
            // or the header is corrupt, just return empty strings.
            Console.WriteLine($"Warning: Corrupt file '{FullFilePath}'. Exception: {corrupt.Message}");
            return (PossiblyCorruptCategory, "");
        }        
        catch(Exception ex)
        {
            // catch all other exceptions to prevent the program from crashing on unexpected file types
            Console.WriteLine($"Warning: Could not read metadata for '{FullFilePath}'. Exception: {ex.Message}");
            return ("", "");
        }
    }    
}

// --- DATA MODELS ---
public record Settings(List<Category> Categories);
public record Category(string Name, bool IgnoreCategoryInContent, List<Rule> Rules);
public record Rule(string Type, string Value);
public record MoveAction(string FileName, string CategoryName, string RuleTriggered);
public class VideoOcrResult { public HashSet<string> Lines { get; set; } = new(); public bool IsCorrupt { get; set; } = false; }

// NEW MODEL for tracking unmatched details
public record UnmatchedFileInfo(string FileName, List<string> OcrLines, string Title, string Comment);
public record RenameAction(string OldFileName, string NewFileName, string SourceField);