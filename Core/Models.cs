using System.Text.Json;
using System.Text.Json.Serialization;

namespace FileOrganizer;

/// <summary>The rules (settings.json). The order of the categories is their priority: the first match wins.</summary>
public sealed class Settings
{
    public List<Category> Categories { get; set; } = new();

    /// <summary>Optional: absent in settings files made before the interfaces existed.</summary>
    public Options? Options { get; set; }

    [JsonIgnore] public Options Opt => Options ??= new Options();

    public Settings Clone() => JsonSerializer.Deserialize<Settings>(JsonSerializer.Serialize(this, SettingsStore.Json), SettingsStore.Json)!;

    /// <summary>Fills in what is missing, so code never meets a null list or name.</summary>
    public void Normalize()
    {
        Categories ??= new();
        foreach (var c in Categories)
        {
            c.Name = (c.Name ?? "").Trim();
            c.Rules ??= new();
            foreach (var r in c.Rules)
            {
                r.Type = r.Type?.Equals("Regex", StringComparison.OrdinalIgnoreCase) == true ? "Regex" : "Contains";
                r.Value ??= "";
            }
        }
    }
}

public sealed class Category
{
    public string Name { get; set; } = "";

    /// <summary>
    /// False: the name itself is a rule (it matches when the name is found in the file name, a tag or the OCR text).
    /// True: for a general word (like "Barry") that must not match by itself.
    /// </summary>
    public bool IgnoreCategoryInContent { get; set; }

    public List<Rule> Rules { get; set; } = new();
}

public sealed class Rule
{
    /// <summary>"Contains" (plain text) or "Regex".</summary>
    public string Type { get; set; } = "Contains";
    public string Value { get; set; } = "";

    public Rule() { }
    public Rule(string type, string value) { Type = type; Value = value; }
}

public sealed class Options
{
    /// <summary>File types that get the deep scan (OCR of video frames). Null: the defaults.</summary>
    public string[]? VideoExtensions { get; set; }

    /// <summary>Stop reading frames of a video as soon as one frame has text that matches a rule (faster).</summary>
    public bool StopOcrAtFirstMatch { get; set; } = true;

    public static readonly string[] DefaultVideoExtensions = { ".mp4", ".mpeg", ".mpv", ".flv", ".mkv", ".avi", ".mov", ".wmv" };
}

/// <summary>Where a file's category came from.</summary>
public enum MatchSource { None, Filename, Title, Comment, Ocr, Corrupt }

/// <summary>What the engine found out about one file, and what it decided.</summary>
public sealed class FileDecision
{
    public string FileName { get; set; } = "";
    public bool IsVideo { get; set; }

    // --- Facts (read from the file) ---
    public string Title { get; set; } = "";
    public string Comment { get; set; } = "";
    public List<string> OcrLines { get; set; } = new();
    public bool OcrCorrupt { get; set; }
    public bool FromCache { get; set; }

    // --- The decision ---
    public string? Category { get; set; }
    public MatchSource Source { get; set; }
    /// <summary>The rule that matched, as the scripts show it ("Regex: ...", "OCR Match: ...").</summary>
    public string Trigger { get; set; } = "";

    /// <summary>A new name from the title or comment tag, when it differs from the file name.</summary>
    public string? RenameTo { get; set; }
    public string? RenameSource { get; set; }

    // --- The user's choices in the interfaces ---
    /// <summary>Another category than the rules chose.</summary>
    public string? OverrideCategory { get; set; }
    /// <summary>Leave the file where it is.</summary>
    public bool Skip { get; set; }
    /// <summary>The file has been moved (or renamed) by the program.</summary>
    public bool Done { get; set; }

    [JsonIgnore] public bool IsMatched => Category != null;
    /// <summary>The category the file will go to: the user's choice, else the rules' (null: stays).</summary>
    [JsonIgnore] public string? EffectiveCategory => Skip ? null : (OverrideCategory ?? Category);
}

public sealed class OrganizeResult
{
    public string Folder { get; set; } = "";
    public List<FileDecision> Decisions { get; set; } = new();
    public TimeSpan Elapsed { get; set; }
    public DateTime FinishedUtc { get; set; }

    public IEnumerable<FileDecision> Matched => Decisions.Where(d => d.Category != null);
    public IEnumerable<FileDecision> Unmatched => Decisions.Where(d => d.Category == null);
    public IEnumerable<FileDecision> Renames => Decisions.Where(d => d.RenameTo != null);
}

/// <summary>Cached OCR of one video (the same file format the command line always made).</summary>
public sealed class VideoOcrResult
{
    public HashSet<string> Lines { get; set; } = new();
    public bool IsCorrupt { get; set; }
}
