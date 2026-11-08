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

        try
        {
            using var capture = new VideoCapture(path, VideoCaptureAPIs.FFMPEG);
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
        return result;
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

/// <summary>Tidying the text found in a video before it is shown (matching still uses every line that was read).</summary>
public static class OcrText
{
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

    /// <summary>The picture file for a video (made on first use), or null when it can't be made.</summary>
    public static string? Ensure(string videoPath, string folder)
    {
        try
        {
            var info = new FileInfo(videoPath);
            if (!info.Exists) return null;
            string key = Convert.ToHexString(SHA1.HashData(Encoding.UTF8.GetBytes($"{info.FullName}|{info.Length}|{info.LastWriteTimeUtc.Ticks}"))).ToLowerInvariant();
            string target = Path.Combine(folder, key + ".png");
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

            using var small = new Mat();
            int height = Math.Max(1, (int)Math.Round(frame.Height * (double)Width / frame.Width));
            Cv2.Resize(frame, small, new Size(Width, height));
            Directory.CreateDirectory(folder);
            // Encoded in memory and written with .NET: OpenCV's own writer can't take non-ASCII folder names on Windows
            Cv2.ImEncode(".png", small, out byte[] png);
            File.WriteAllBytes(target, png);
            return target;
        }
        catch (Exception)
        {
            return null;
        }
    }
}
