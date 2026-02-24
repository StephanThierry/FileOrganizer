using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OpenCvSharp;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;

public class FileProcessor
{
    private readonly string _targetFolder;
    private readonly string _settingsPath;
    private readonly string _batPath;
    private readonly string _unmatchedPath;
    private readonly string _ocrCachePath;
    private const string PossiblyCorruptCategory = "_CHECK_";

    private readonly HashSet<string> _videoExtensions = new(StringComparer.OrdinalIgnoreCase) 
        { ".mp4", ".mpeg", ".mpv", ".flv", ".mkv", ".avi", ".mov", ".wmv" };

    private readonly HashSet<string> _ignoredFiles;

    public FileProcessor(string targetFolder, string settingsPath)
    {
        _targetFolder = targetFolder;
        _settingsPath = settingsPath;
        _batPath = Path.Combine(targetFolder, "move_files.bat");
        _ocrCachePath = Path.Combine(targetFolder, "_FileProcessor_VideoOcrResult.json");
        _unmatchedPath = Path.Combine(targetFolder, "unmatched_files.txt");

        //Console.SetError(TextWriter.Null);
        Cv2.SetLogLevel(0);

        _ignoredFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "settings.json", "move_files.bat", "unmatched_files.txt",
            AppDomain.CurrentDomain.FriendlyName + ".exe",
            AppDomain.CurrentDomain.FriendlyName + ".dll"
        };
    }

    public async Task RunAsync()
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
                    ocrResult = await GetVideoOcrTextAsync(Path.Combine(_targetFolder, file!));
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

            // No Match - Save file + all detected OCR text
           unmatchedData.Add(new UnmatchedFileInfo(file!, new List<string>(), title ?? "", comments ?? ""));
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

    private async Task<VideoOcrResult> GetVideoOcrTextAsync(string path)
{
    var result = new VideoOcrResult();
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

            // --- PRE-PROCESSING STEP 2.0 (Noise Reduction Focus) ---
            
            using Mat gray = new Mat();
            Cv2.CvtColor(frame, gray, ColorConversionCodes.BGR2GRAY);

            // 1. Upscale for better character definition
            using Mat resized = new Mat();
            Cv2.Resize(gray, resized, new OpenCvSharp.Size(gray.Width * 2, gray.Height * 2), interpolation: InterpolationFlags.Cubic);

            // 2. Bilateral Filter: This is the "Magic" for noise.
            // It smooths flat areas (noise) but keeps the edges of text sharp.
            using Mat denoised = new Mat();
            Cv2.BilateralFilter(resized, denoised, d: 9, sigmaColor: 75, sigmaSpace: 75);

            // 3. CLAHE for local contrast (brings out the "Angelo Godshack" text)
            using Mat finalProcessed = new Mat();
            using (var clahe = Cv2.CreateCLAHE(clipLimit: 3.0, tileGridSize: new OpenCvSharp.Size(8, 8)))
            {
                clahe.Apply(denoised, finalProcessed);
            }

            // NOTE: We are NOT using AdaptiveThreshold here because it creates the symbol noise.
            // Windows OCR often performs better on a clean, high-contrast grayscale image.

            string text = await RunWindowsMediaOcr(finalProcessed);
            
            foreach (var line in text.Split('\n', '\r').Where(l => l.Trim().Length > 2))
            {
                string cleanLine = line.Trim();

                // OPTIONAL: Basic junk filter (ignores lines that are just symbols)
                if (IsMostlyJunk(cleanLine)) continue;

                Console.WriteLine($"  OCR detection: {cleanLine}");
                result.Lines.Add(cleanLine);
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

    private async Task<string> RunWindowsMediaOcr(Mat mat)
    {
        byte[] bytes = mat.ImEncode(".png");
        using var stream = new InMemoryRandomAccessStream();
        using (var writer = new DataWriter(stream.GetOutputStreamAt(0)))
        {
            writer.WriteBytes(bytes);
            await writer.StoreAsync();
        }
        var decoder = await BitmapDecoder.CreateAsync(stream);
        var bitmap = await decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);
        var engine = OcrEngine.TryCreateFromUserProfileLanguages() ?? OcrEngine.TryCreateFromLanguage(new Windows.Globalization.Language("en-US"));
        var ocrResult = await engine.RecognizeAsync(bitmap);
        return ocrResult.Text;
    }

    private void GenerateBatFile(List<MoveAction> actions)
    {
        // Use UTF8 encoding with a Byte Order Mark (BOM) so Windows recognizes it
        var utf8WithBom = new System.Text.UTF8Encoding(true);

        using var writer = new StreamWriter(_batPath, false, utf8WithBom);
        
        // 1. Set Code Page to UTF-8 (65001) to support special chars like ’ and –
        // 2. Clear the screen or redirect to nul so the user doesn't see the "Active code page" message
        writer.WriteLine("@echo off");
        writer.WriteLine("chcp 65001 > nul"); 
        writer.WriteLine("cd /d \"%~dp0\"");
        writer.WriteLine();

        // Create directories
        foreach (var cat in actions.Select(a => a.CategoryName).Distinct())
        {
            // Escape % in directory names too
            string safeCat = cat.Replace("%", "%%");
            writer.WriteLine($"IF NOT EXIST \"{safeCat}\" mkdir \"{safeCat}\"");
        }

        writer.WriteLine("\nREM --- Moving Files ---");
        foreach (var act in actions)
        {
            // ESCAPE PERCENT SIGNS: % must be written as %% in batch files
            string safeFileName = act.FileName.Replace("%", "%%");
            string safeCategory = act.CategoryName.Replace("%", "%%");

            writer.WriteLine($"move \"{safeFileName}\" \"{safeCategory}\\\"");
            writer.WriteLine($"REM Triggered: {act.RuleTriggered}");
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

    string renameBatPath = Path.Combine(_targetFolder, "rename_files.bat");
    var utf8WithBom = new System.Text.UTF8Encoding(true);

    using var writer = new StreamWriter(renameBatPath, false, utf8WithBom);
    
    writer.WriteLine("@echo off");
    writer.WriteLine("chcp 65001 > nul");
    writer.WriteLine("cd /d \"%~dp0\"");
    writer.WriteLine("REM === Auto-generated Rename Script (Title/Comment) ===");
    writer.WriteLine();

    foreach (var action in actions)
    {
        // Escape % for batch
        string oldFile = action.OldFileName.Replace("%", "%%");
        string newFile = action.NewFileName.Replace("%", "%%");

        writer.WriteLine($"ren \"{oldFile}\" \"{newFile}\"");
        writer.WriteLine($"REM Source: {action.SourceField}\n");
    }
    
    Console.WriteLine($"Generated 'rename_files.bat' with {actions.Count} commands.");
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
            return ("", "");
        } catch(TagLib.CorruptFileException corrupt)
        {
            // If the file format isn't supported (like a .txt file) 
            // or the header is corrupt, just return empty strings.
            return (PossiblyCorruptCategory, "");
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