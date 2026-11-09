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

    /// <summary>
    /// Joins categories that have the same name (ignoring case) into the first of them: every rule that is not there yet is added to it.
    /// The name stays a rule of the category unless every one of them said to ignore it. Returns how many categories were joined away.
    /// </summary>
    public int MergeDuplicateCategories()
    {
        Normalize();
        var first = new Dictionary<string, Category>(StringComparer.OrdinalIgnoreCase);
        var merged = new List<Category>();
        int joined = 0;
        foreach (var c in Categories)
        {
            if (c.Name.Length == 0 || !first.TryGetValue(c.Name, out var keep))
            {
                if (c.Name.Length > 0) first[c.Name] = c;
                merged.Add(c);
                continue;
            }
            joined++;
            keep.IgnoreCategoryInContent &= c.IgnoreCategoryInContent;
            foreach (var r in c.Rules)
                if (!keep.Rules.Any(k => k.SameAs(r)))
                    keep.Rules.Add(r);
        }
        Categories = merged;
        return joined;
    }

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
                r.Type = r.Type?.Equals("Regex", StringComparison.OrdinalIgnoreCase) == true ? "Regex"
                    : r.Type?.Equals("All", StringComparison.OrdinalIgnoreCase) == true ? "All" : "Contains";
                r.Value ??= "";
                if (r.Type == "All")
                {
                    r.Terms ??= new();
                    foreach (var t in r.Terms) { t.Type = t.IsRegex ? "Regex" : "Contains"; t.Value ??= ""; }
                }
                else r.Terms = null;
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

/// <summary>One of the terms of an "All" rule: plain text ("Contains") or a "Regex".</summary>
public sealed class RuleTerm
{
    public string Type { get; set; } = "Contains";
    public string Value { get; set; } = "";

    public RuleTerm() { }
    public RuleTerm(string type, string value) { Type = type; Value = value; }
    [JsonIgnore] public bool IsRegex => Type.Equals("Regex", StringComparison.OrdinalIgnoreCase);

    /// <summary>For messages: plain text as it is, a regex between slashes.</summary>
    public override string ToString() => IsRegex ? "/" + Value + "/" : Value;
}

public sealed class Rule
{
    /// <summary>
    /// "Contains" (plain text), "Regex", or "All": a rule with several terms (Terms) that all have to be found somewhere in the
    /// file: its name, its tags and the text read from the video (each term in any of them, not necessarily in the same place).
    /// </summary>
    public string Type { get; set; } = "Contains";
    public string Value { get; set; } = "";

    /// <summary>The terms of an "All" rule (each Contains or Regex); not used, and not saved, by the other types.</summary>
    public List<RuleTerm>? Terms { get; set; }

    public Rule() { }
    public Rule(string type, string value) { Type = type; Value = value; }
    public static Rule AllOf(IEnumerable<RuleTerm> terms) => new() { Type = "All", Value = "", Terms = terms.ToList() };

    [JsonIgnore] public bool IsAll => Type.Equals("All", StringComparison.OrdinalIgnoreCase);

    /// <summary>The rule in words: the text, the pattern, or "All of: rec + /R.?C/ + hub".</summary>
    public string Describe() => IsAll ? "All of: " + string.Join(" + ", (Terms ?? new()).Select(t => t.ToString())) : Value;

    /// <summary>The same rule (same type and text, or the same terms): adding it twice is pointless.</summary>
    public bool SameAs(Rule other) =>
        IsAll || other.IsAll
            ? IsAll && other.IsAll && (Terms ?? new()).Select(t => t.Type + ":" + t.Value.Trim().ToLowerInvariant()).OrderBy(x => x)
                .SequenceEqual((other.Terms ?? new()).Select(t => t.Type + ":" + t.Value.Trim().ToLowerInvariant()).OrderBy(x => x))
            : Type == other.Type && Value.Trim().Equals(other.Value.Trim(), StringComparison.OrdinalIgnoreCase);
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

    /// <summary>No rule matched it and it has not been moved: it is listed under Unmatched, also while the user has picked a folder for it
    /// and not moved it yet.</summary>
    [JsonIgnore] public bool IsUnmatched => Category == null && !Done;
    /// <summary>Where it is now, once Done.</summary>
    public string? MovedTo { get; set; }

    public const string LeaveHere = "(leave it here)";

    /// <summary>What the file's folder choice shows: the category it goes to, or "(leave it here)".</summary>
    [JsonIgnore] public string Choice => EffectiveCategory ?? LeaveHere;

    /// <summary>The user picks a folder for the file, or "(leave it here)". Picking the rules' own category clears the override.</summary>
    public void Choose(string choice)
    {
        if (choice == LeaveHere) { Skip = true; return; }
        Skip = false;
        OverrideCategory = string.Equals(choice, Category, StringComparison.OrdinalIgnoreCase) ? null : choice;
    }

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
