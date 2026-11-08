using System.Text.Json;

namespace FileOrganizer;

/// <summary>What was written when the scripts were made.</summary>
public sealed record ScriptReport(string? MoveScript, int Moves, string? RenameScript, int Renames, string? UnmatchedReport)
{
    public bool Any => MoveScript != null || RenameScript != null;
}

/// <summary>
/// Everything the desktop and web interfaces share: the data folder, the rules, the folder being organized, scanning, the
/// result and what is done with it (moving files, writing scripts, making rules from what a scan found).
/// </summary>
public sealed class AppController : IDisposable
{
    private sealed class State
    {
        public string LastFolder { get; set; } = "";
        public List<string> Recent { get; set; } = new();
    }

    private State state = new();

    public string DataFolder { get; }
    public Settings Settings { get; private set; }
    /// <summary>Set when settings.json exists but could not be read (the rules are then empty, and saving would replace it).</summary>
    public string? SettingsError { get; }
    /// <summary>True when there was no settings.json yet.</summary>
    public bool IsFirstRun { get; }
    public OrganizeEngine Engine { get; }
    public OrganizeResult? Result { get; private set; }
    public string TargetFolder => state.LastFolder;
    public IReadOnlyList<string> RecentFolders => state.Recent;
    public string ThumbFolder => DataLocation.ThumbFolder(DataFolder);
    public string ModelsFolder => DataLocation.ModelsFolder(DataFolder);

    /// <summary>Raised when there is a new result (a scan finished, files were moved, the rules were applied again).</summary>
    public event Action? ResultChanged;
    /// <summary>Raised when the rules were saved.</summary>
    public event Action? SettingsChanged;
    /// <summary>Raised when another folder was chosen.</summary>
    public event Action? TargetChanged;
    /// <summary>Raised when a page changed what will happen to files (a folder choice): the counts in the top bar follow it.</summary>
    public event Action? DecisionsEdited;

    public void NotifyDecisionsEdited() => DecisionsEdited?.Invoke();

    private AppController(string dataFolder)
    {
        DataFolder = dataFolder;
        IsFirstRun = !SettingsStore.Exists(dataFolder);
        try { Settings = SettingsStore.Load(dataFolder); }
        catch (InvalidDataException e) { Settings = new Settings(); SettingsError = e.Message; }

        try
        {
            string stateFile = StatePath;
            if (File.Exists(stateFile)) state = JsonSerializer.Deserialize<State>(File.ReadAllText(stateFile)) ?? new State();
        }
        catch (Exception) { /* a broken state file is just forgotten */ }

        Engine = new OrganizeEngine(() => new OpenCvVideoOcr(ModelsFolder)) { ThumbFolder = ThumbFolder };
        Engine.Completed += result => { Result = result; ResultChanged?.Invoke(); };
    }

    public static AppController Open(string dataFolder)
    {
        dataFolder = Path.GetFullPath(dataFolder);
        Directory.CreateDirectory(dataFolder);
        return new AppController(dataFolder);
    }

    public void Dispose() => Engine.Dispose();

    private string StatePath => Path.Combine(DataFolder, "state.json");

    private void SaveState()
    {
        try { File.WriteAllText(StatePath, JsonSerializer.Serialize(state, new JsonSerializerOptions { WriteIndented = true })); }
        catch (Exception) { /* read-only data folder: the folder is just not remembered */ }
    }

    // ---------- The folder to organize ----------

    /// <summary>Chooses the folder to organize. Returns an error message, or null when it was accepted.</summary>
    public string? SetTargetFolder(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return "Choose a folder first.";
        path = path.Trim().Trim('"');
        if (!Directory.Exists(path)) return $"The folder '{path}' does not exist.";
        path = Path.GetFullPath(path);
        bool changed = !string.Equals(state.LastFolder, path, StringComparison.OrdinalIgnoreCase);
        state.LastFolder = path;
        state.Recent.RemoveAll(f => string.Equals(f, path, StringComparison.OrdinalIgnoreCase));
        state.Recent.Insert(0, path);
        if (state.Recent.Count > 8) state.Recent.RemoveRange(8, state.Recent.Count - 8);
        SaveState();
        if (changed) TargetChanged?.Invoke();
        return null;
    }

    /// <summary>Takes a folder out of the list of recently used folders (the folder itself is not touched).</summary>
    public void ForgetFolder(string path)
    {
        if (state.Recent.RemoveAll(f => string.Equals(f, path, StringComparison.OrdinalIgnoreCase)) > 0) SaveState();
    }

    // ---------- Scanning ----------

    /// <summary>Starts a scan of the chosen folder. Returns an error message, or null when it started.</summary>
    public string? StartScan()
    {
        if (SettingsError != null) return SettingsError;
        if (string.IsNullOrWhiteSpace(TargetFolder) || !Directory.Exists(TargetFolder)) return "Choose the folder to organize first.";
        if (Settings.Categories.Count == 0) return "There are no categories yet. Add some under Rules first.";
        if (!Engine.Start(TargetFolder, Settings)) return "A scan is already running.";
        return null;
    }

    public void StopScan() => Engine.Cancel();

    // ---------- The rules ----------

    /// <summary>Saves the rules. Returns what is wrong (nothing is saved then), or an empty list when it worked.</summary>
    public List<string> SaveSettings(Settings settings)
    {
        settings.Normalize();
        var problems = SettingsStore.Validate(settings);
        if (problems.Count > 0) return problems;
        SettingsStore.Save(DataFolder, settings);
        Settings = settings.Clone();
        if (Result != null) OrganizeEngine.Reevaluate(Result, Settings);
        SettingsChanged?.Invoke();
        if (Result != null) ResultChanged?.Invoke();
        return problems;
    }

    /// <summary>
    /// Adds a rule to a category (made if it is new; a new category does not match by its own name, only by its rules) and applies
    /// the rules again to the result, so the files it catches move out of Unmatched at once. Returns a message for the user.
    /// </summary>
    public string AddRule(string category, string type, string value, out bool ok)
    {
        ok = false;
        category = (category ?? "").Trim();
        value = (value ?? "").Trim();
        if (category.Length == 0) return "Choose a category, or type the name of a new one.";
        if (!FolderNames.IsSafe(category)) return $"\"{category}\" can't be a folder name.";
        var rule = new Rule(type.Equals("Regex", StringComparison.OrdinalIgnoreCase) ? "Regex" : "Contains", value);
        var error = RuleMatcher.Validate(rule);
        if (error != null) return error[..1].ToUpper() + error[1..];

        var settings = Settings.Clone();
        var existing = settings.Categories.FirstOrDefault(c => c.Name.Equals(category, StringComparison.OrdinalIgnoreCase));
        bool isNew = existing == null;
        if (existing == null)
        {
            existing = new Category { Name = category, IgnoreCategoryInContent = true };
            settings.Categories.Add(existing);
        }
        if (existing.Rules.Any(r => r.Type == rule.Type && r.Value.Equals(rule.Value, StringComparison.OrdinalIgnoreCase)))
            return $"{existing.Name} has that rule already.";
        existing.Rules.Add(rule);

        var problems = SaveSettings(settings);
        if (problems.Count > 0) return problems[0];
        ok = true;
        return isNew ? $"Made the category {category} with the rule \"{value}\"." : $"Added the rule \"{value}\" to {existing.Name}.";
    }

    // ---------- What to do with the result ----------

    /// <summary>The files that will go somewhere: the rules' or the user's choice, not left alone, not moved yet.</summary>
    public IEnumerable<FileDecision> Pending => Result?.Decisions.Where(d => !d.Done && d.EffectiveCategory != null) ?? Enumerable.Empty<FileDecision>();

    /// <summary>Moves the files now (never overwriting). With alsoRename, files with a title or comment tag get that as their name.</summary>
    public MoveReport MoveNow(bool alsoRename)
    {
        if (Result == null) return new MoveReport(0, 0, 0, 0, new List<string>());
        var report = MoveExecutor.Apply(Result.Folder, Result.Decisions, alsoRename);
        ResultChanged?.Invoke();
        return report;
    }

    /// <summary>Writes move_files, rename_files and unmatched_files.txt in the scanned folder, from the result as it is now.</summary>
    public ScriptReport WriteScripts()
    {
        if (Result == null) return new ScriptReport(null, 0, null, 0, null);
        var todo = Result.Decisions.Where(d => !d.Done).ToList();
        string? move = ScriptWriter.WriteMoveScript(Result.Folder, todo);
        string? rename = ScriptWriter.WriteRenameScript(Result.Folder, todo);
        string? unmatched = ScriptWriter.WriteUnmatchedReport(Result.Folder, todo.Where(d => d.EffectiveCategory == null));
        return new ScriptReport(move, todo.Count(d => d.EffectiveCategory != null), rename, todo.Count(d => d.RenameTo != null), unmatched);
    }

    /// <summary>Where the file is now (moved files are in their category folder).</summary>
    public string? PathOf(FileDecision d)
    {
        if (Result == null) return null;
        return d.MovedTo ?? Path.Combine(Result.Folder, d.FileName);
    }

    /// <summary>A small picture of a video (made on first use), or null.</summary>
    public string? ThumbOf(FileDecision d)
    {
        if (!d.IsVideo) return null;
        string? path = PathOf(d);
        return path == null ? null : Thumbnails.Ensure(path, ThumbFolder);
    }

    /// <summary>The names of the categories, in priority order (and the folder for unreadable videos).</summary>
    public List<string> CategoryNames()
    {
        var names = Settings.Categories.Select(c => c.Name).Where(n => n.Length > 0).ToList();
        if (!names.Contains(OrganizeEngine.CorruptCategory)) names.Add(OrganizeEngine.CorruptCategory);
        return names;
    }
}
