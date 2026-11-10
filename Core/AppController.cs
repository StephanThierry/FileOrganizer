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
        try
        {
            Settings = SettingsStore.Load(dataFolder);
            // Categories with the same name are joined once, here (the old file is kept as settings.json.bck)
            if (Settings.MergeDuplicateCategories() > 0)
            {
                try { SettingsStore.Save(dataFolder, Settings); }
                catch (Exception) { /* read-only: it is joined in memory and saved with the next save */ }
            }
        }
        catch (InvalidDataException e) { Settings = new Settings(); SettingsError = e.Message; }

        try
        {
            string stateFile = StatePath;
            if (File.Exists(stateFile)) state = JsonSerializer.Deserialize<State>(File.ReadAllText(stateFile)) ?? new State();
        }
        catch (Exception) { /* a broken state file is just forgotten */ }

        Engine = new OrganizeEngine(() => new OpenCvVideoOcr(ModelsFolder)) { ThumbFolder = ThumbFolder };
        Engine.Completed += result => { Result = result; LastLeftovers = null; dismissedFolders.Clear(); ResultChanged?.Invoke(); };
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

    /// <summary>
    /// Switches "Process folders (experimental)" on or off for the next scans and keeps the choice in settings.json. (Not while settings.json
    /// is unreadable: saving would replace it.)
    /// </summary>
    public void SetProcessFolders(bool on)
    {
        if (SettingsError != null || Settings.Opt.ProcessFolders == on) return;
        var s = Settings.Clone();
        s.Opt.ProcessFolders = on;
        try { SettingsStore.Save(DataFolder, s); } catch (Exception) { return; }
        Settings = s;
        SettingsChanged?.Invoke();
    }

    /// <summary>
    /// Switches "Delete confirmation" on or off and keeps the choice in settings.json (default on). Off: the delete icons on Review and Unmatched
    /// delete the file at once. (Not while settings.json is unreadable: saving would replace it.)
    /// </summary>
    public void SetConfirmDelete(bool on)
    {
        if (SettingsError != null || Settings.Opt.ConfirmDelete == on) return;
        var s = Settings.Clone();
        s.Opt.ConfirmDelete = on;
        try { SettingsStore.Save(DataFolder, s); } catch (Exception) { return; }
        Settings = s;
        SettingsChanged?.Invoke();
    }

    /// <summary>What is left, after moves and deletes, in the folders the scan opened ("Process folders"); null when it opened none or nothing was done yet.</summary>
    public LeftoverReport? LastLeftovers { get; private set; }

    // Folders taken out of the overview by "Delete empty folders" (and not shown as gone afterwards)
    private readonly HashSet<string> dismissedFolders = new(StringComparer.OrdinalIgnoreCase);

    private void RefreshLeftovers()
    {
        var opened = Result?.Decisions.Where(d => d.FolderName != null).Select(d => d.FolderName!).Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(n => !dismissedFolders.Contains(n)).ToList();
        LastLeftovers = Result == null || opened == null || opened.Count == 0 ? null : FolderLeftovers.Build(Result.Folder, opened);
    }

    /// <summary>
    /// "Delete empty folders" in the Status window: deletes the opened folders with no file left in them from the drive (each checked
    /// again first) and takes them, and the ones that are gone already, out of the overview. Returns what was done.
    /// </summary>
    public EmptyFoldersResult DeleteEmptyFolders()
    {
        if (Result == null || LastLeftovers == null) return new EmptyFoldersResult(Array.Empty<string>(), Array.Empty<string>());
        var result = FolderLeftovers.DeleteEmpty(Result.Folder, LastLeftovers);
        foreach (var name in result.Deleted) dismissedFolders.Add(name);
        foreach (var gone in LastLeftovers.Folders.Where(f => !f.Exists)) dismissedFolders.Add(gone.Name);
        RefreshLeftovers();
        ResultChanged?.Invoke();
        return result;
    }

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
        return AddRule(category, new Rule(type.Equals("Regex", StringComparison.OrdinalIgnoreCase) ? "Regex" : "Contains", value), out ok);
    }

    /// <summary>The same, for a rule with several terms that all have to be found ("All of").</summary>
    public string AddAllRule(string category, IEnumerable<RuleTerm> terms, out bool ok) =>
        AddRule(category, Rule.AllOf(terms.Select(t => new RuleTerm(t.IsRegex ? "Regex" : "Contains", (t.Value ?? "").Trim()))), out ok);

    private string AddRule(string category, Rule rule, out bool ok)
    {
        ok = false;
        category = (category ?? "").Trim();
        if (category.Length == 0) return "Choose a category, or type the name of a new one.";
        if (!FolderNames.IsSafe(category)) return $"\"{category}\" can't be a folder name.";
        string value = rule.Describe();
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
        if (existing.Rules.Any(r => r.SameAs(rule)))
            return $"{existing.Name} has that rule already.";
        existing.Rules.Add(rule);

        var problems = SaveSettings(settings);
        if (problems.Count > 0)
            return "The rule was not added, because the rules saved so far have a problem: " + string.Join(" ", problems) + " Fix it under Rules, then try again.";
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
        RefreshLeftovers();
        ResultChanged?.Invoke();
        return report;
    }

    /// <summary>Moves one file now to the folder it goes to (never overwriting). Nothing happens for a file that stays or was moved.</summary>
    public MoveReport MoveOne(FileDecision d, bool alsoRename)
    {
        if (Result == null || d.Done || d.EffectiveCategory == null) return new MoveReport(0, 0, 0, 0, new List<string>());
        var report = MoveExecutor.Apply(Result.Folder, new[] { d }, alsoRename);
        RefreshLeftovers();
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

    /// <summary>True when deleting this file puts it in the Recycle Bin (trash): for the question asked before deleting.</summary>
    public bool CanRecover(FileDecision d) => PathOf(d) is { } path && RecycleBin.IsRecoverable(path);

    /// <summary>
    /// Deletes the file: to the Recycle Bin (trash), or for good where there is none (a network share), which the interface says
    /// before asking. The file leaves the result, so it disappears from Review and Unmatched. Returns whether it worked, and a message.
    /// </summary>
    public (bool Ok, string Message) DeleteFile(FileDecision d)
    {
        string? path = PathOf(d);
        if (Result == null || path == null) return (false, "There is no file to delete.");
        try
        {
            string name = Path.GetFileName(path);
            bool recoverable = RecycleBin.Delete(path);
            Result.Decisions.Remove(d);
            if (d.FolderName != null) RefreshLeftovers();   // it was in a folder that was opened: what is left there changed
            ResultChanged?.Invoke();
            return (true, recoverable ? $"Moved {name} to the {RecycleBin.Name}." : $"Deleted {name} permanently: this location has no {RecycleBin.Name}.");
        }
        catch (Exception e)
        {
            return (false, $"Could not delete {d.FileName}: {e.Message}");
        }
    }

    /// <summary>Where the file is now (moved files are in their category folder).</summary>
    public string? PathOf(FileDecision d)
    {
        if (Result == null) return null;
        return d.MovedTo ?? Path.Combine(Result.Folder, d.RelativePath);
    }

    /// <summary>A small picture of a video (made on first use), or null.</summary>
    public string? ThumbOf(FileDecision d)
    {
        if (!d.IsVideo) return null;
        string? path = PathOf(d);
        return path == null ? null : Thumbnails.Ensure(path, ThumbFolder);
    }

    /// <summary>One of the pictures of a video (0 to <see cref="Thumbnails.Parts"/> - 1), made on first use, or null.</summary>
    public string? ThumbOf(FileDecision d, int part)
    {
        if (!d.IsVideo) return null;
        string? path = PathOf(d);
        return path == null ? null : Thumbnails.EnsurePart(path, ThumbFolder, part);
    }

    /// <summary>The names of the categories for a list to choose from: alphabetical (and the folder for unreadable videos).</summary>
    public List<string> CategoryNames()
    {
        var names = Settings.Categories.Select(c => c.Name).Where(n => n.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        names.RemoveAll(n => n.Equals(OrganizeEngine.CorruptCategory, StringComparison.OrdinalIgnoreCase));
        names.Sort(StringComparer.CurrentCultureIgnoreCase);
        names.Add(OrganizeEngine.CorruptCategory);   // last: the folder for videos that can't be read
        return names;
    }
}
