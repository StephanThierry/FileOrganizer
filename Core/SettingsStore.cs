using System.Text.Json;
using System.Text.Json.Serialization;

namespace FileOrganizer;

/// <summary>Where the program keeps its files, and reading and writing settings.json.</summary>
public static class DataLocation
{
    /// <summary>Value after a command-line option, e.g. Arg(args, "--data").</summary>
    public static string? Arg(string[] args, string name)
    {
        int i = Array.IndexOf(args, name);
        return i >= 0 && i + 1 < args.Length ? args[i + 1] : null;
    }

    /// <summary>--data, then FILEORG_DATA, then the fallback (may be null).</summary>
    public static string? ResolveDataFolder(string[] args, string? fallback)
    {
        string? chosen = Arg(args, "--data");
        if (string.IsNullOrWhiteSpace(chosen)) chosen = Environment.GetEnvironmentVariable("FILEORG_DATA");
        return string.IsNullOrWhiteSpace(chosen) ? fallback : Path.GetFullPath(chosen);
    }

    /// <summary>
    /// The data folder when none was chosen: the current folder or the program's own folder if they hold a settings.json
    /// (what the command-line version always did), otherwise the user's application data folder.
    /// </summary>
    public static string DefaultDataFolder()
    {
        string current = Directory.GetCurrentDirectory();
        if (File.Exists(Path.Combine(current, SettingsStore.FileName))) return current;
        if (File.Exists(Path.Combine(AppContext.BaseDirectory, SettingsStore.FileName))) return AppContext.BaseDirectory;
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "FileOrganizer");
    }

    /// <summary>Pictures of videos, for the interfaces.</summary>
    public static string ThumbFolder(string dataFolder) => Path.Combine(dataFolder, "_thumbs");

    /// <summary>
    /// The OCR models (about 13 MB, downloaded once). A "models" folder next to the program, as the command line always used,
    /// is used if it has them; otherwise they live in the data folder, so every interface shares one download.
    /// </summary>
    public static string ModelsFolder(string dataFolder)
    {
        string beside = Path.Combine(AppContext.BaseDirectory, "models");
        if (Directory.Exists(beside) && Directory.GetFiles(beside, "*.onnx").Length >= 2) return beside;
        return Path.Combine(dataFolder, "models");
    }
}

public static class SettingsStore
{
    public const string FileName = "settings.json";

    public static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    public static string PathIn(string dataFolder) => Path.Combine(dataFolder, FileName);

    public static bool Exists(string dataFolder) => File.Exists(PathIn(dataFolder));

    /// <summary>Reads settings.json. A missing file gives empty settings; a broken one an InvalidDataException with the reason.</summary>
    public static Settings Load(string dataFolder)
    {
        string path = PathIn(dataFolder);
        if (!File.Exists(path)) return new Settings();
        try
        {
            var settings = JsonSerializer.Deserialize<Settings>(File.ReadAllText(path), Json) ?? new Settings();
            settings.Normalize();
            return settings;
        }
        catch (JsonException e)
        {
            throw new InvalidDataException($"{path} is not valid: {e.Message}", e);
        }
    }

    /// <summary>Writes settings.json (the previous one is kept as settings.json.bck first).</summary>
    public static void Save(string dataFolder, Settings settings)
    {
        Directory.CreateDirectory(dataFolder);
        settings.Normalize();
        string path = PathIn(dataFolder), tmp = path + ".tmp";
        File.WriteAllText(tmp, JsonSerializer.Serialize(settings, Json));
        if (File.Exists(path)) File.Copy(path, path + ".bck", overwrite: true);
        File.Move(tmp, path, overwrite: true);
    }

    /// <summary>What is wrong with the settings (empty: nothing). Shown before saving and before a scan.</summary>
    public static List<string> Validate(Settings settings)
    {
        var problems = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var c in settings.Categories)
        {
            if (string.IsNullOrWhiteSpace(c.Name)) { problems.Add("A category has no name."); continue; }
            if (!FolderNames.IsSafe(c.Name)) problems.Add($"\"{c.Name}\" can't be a folder name (no \\ / : * ? \" < > | and not . or ..).");
            if (!seen.Add(c.Name)) problems.Add($"There are two categories called \"{c.Name}\".");
            foreach (var r in c.Rules)
            {
                var error = RuleMatcher.Validate(r);
                if (error != null) problems.Add($"{c.Name}: {error}");
            }
        }
        return problems;
    }
}

/// <summary>A category becomes a folder inside the scanned folder, so its name must be a plain folder name.</summary>
public static class FolderNames
{
    public static bool IsSafe(string name)
    {
        if (string.IsNullOrWhiteSpace(name) || name.Length > 200) return false;
        if (name is "." or "..") return false;
        if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) return false;
        return name == name.TrimEnd('.', ' ');
    }

    /// <summary>A file name made from a tag: illegal characters become _, and no trailing dots or spaces.</summary>
    public static string SafeFileName(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return "unnamed";
        string safe = name;
        foreach (char c in Path.GetInvalidFileNameChars()) safe = safe.Replace(c, '_');
        return safe.Trim().TrimEnd('.');
    }
}
