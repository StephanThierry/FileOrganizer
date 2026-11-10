using System.Collections.ObjectModel;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace FileOrganizer.Desktop.ViewModels;

/// <summary>One folder that was opened, as a line in the Status window.</summary>
public sealed record StatusRow(string Folder, string Left, string Size, string Status, string Details, bool IsEmpty, bool HasFiles);

/// <summary>The files that will move: each with its category, why, and a choice to send it elsewhere or leave it.</summary>
public partial class ReviewViewModel : ViewModelBase
{
    public const string LeaveHere = FileDecision.LeaveHere;
    public const string AllCategories = "All categories";

    private readonly AppController app;
    private readonly MainViewModel main;
    private readonly Dictionary<FileDecision, FileRow> rows = new();

    public ObservableCollection<FileRow> Rows { get; } = new();
    /// <summary>What a file can be sent to: leave it, or any category.</summary>
    public ObservableCollection<string> Choices { get; } = new();
    public ObservableCollection<string> Filters { get; } = new();

    [ObservableProperty] private string selectedFilter = AllCategories;
    [ObservableProperty] private string search = "";
    [ObservableProperty] private bool alsoRename;
    [ObservableProperty] private bool hasRows;
    [ObservableProperty] private bool hasResult;
    [ObservableProperty] private string summary = "";
    [ObservableProperty] private string detail = "";
    [ObservableProperty] private bool hasPending;
    [ObservableProperty] private string? scriptsText;
    [ObservableProperty] private string? scriptsFolder;

    // The Status window: what is left in the folders that were opened (Process folders), after files were moved or deleted
    public ObservableCollection<StatusRow> StatusRows { get; } = new();
    /// <summary>There is something to show (the Status button is there).</summary>
    [ObservableProperty] private bool hasLeftovers;
    [ObservableProperty] private bool statusVisible;
    [ObservableProperty] private string statusHeadline = "";
    [ObservableProperty] private bool canDeleteEmpty;
    [ObservableProperty] private string statusNote = "";
    public bool HasStatusNote => StatusNote.Length > 0;
    partial void OnStatusNoteChanged(string value) => OnPropertyChanged(nameof(HasStatusNote));

    private void ShowLeftovers()
    {
        var report = app.LastLeftovers;
        HasLeftovers = report != null;
        StatusRows.Clear();
        if (report == null) { StatusHeadline = ""; CanDeleteEmpty = false; StatusVisible = false; return; }
        StatusHeadline = report.Headline;
        CanDeleteEmpty = report.CanDeleteEmpty;
        foreach (var f in report.Folders.OrderBy(f => f.Exists && f.Files.Count > 0 ? 0 : f.Exists ? 1 : 2).ThenBy(f => f.Name, StringComparer.OrdinalIgnoreCase))
            StatusRows.Add(new StatusRow(f.Name, f.Exists ? f.Files.Count.ToString("N0") : "-", f.Exists ? Fmt.Size(f.Bytes) : "-", f.Status, f.Details(), f.IsEmpty, f.Exists && f.Files.Count > 0));
    }

    [RelayCommand]
    private void OpenStatus()
    {
        StatusNote = "";
        ShowLeftovers();
        StatusVisible = HasLeftovers;
    }

    [RelayCommand] private void CloseStatus() => StatusVisible = false;

    // "Delete empty folders": deletes the folders with no file left from the drive and takes them out of the overview
    [RelayCommand]
    private void DeleteEmptyFolders()
    {
        var result = app.DeleteEmptyFolders();
        ShowLeftovers();
        var parts = new List<string>();
        if (result.Deleted.Count > 0) parts.Add($"Deleted {Fmt.Plural(result.Deleted.Count, "empty folder", "empty folders")}.");
        if (result.Messages.Count > 0) parts.Add(string.Join(" ", result.Messages.Take(3)));
        if (parts.Count == 0) parts.Add("There was no empty folder.");
        StatusNote = HasLeftovers ? string.Join(" ", parts) : "";
        if (!HasLeftovers) ShowToast(string.Join(" ", parts) + " Nothing more to show.", 6000);
    }

    public bool HasScripts => ScriptsText != null;
    public bool HasRenames => app.Result?.Decisions.Any(d => d.RenameTo != null) ?? false;

    public ReviewViewModel(AppController app, MainViewModel main)
    {
        this.app = app;
        this.main = main;
        BindConfirmDelete(app);
        app.ResultChanged += () => Dispatcher.UIThread.Post(Reload);
        app.SettingsChanged += () => Dispatcher.UIThread.Post(RefreshChoices);
        RefreshChoices();
        Reload();
    }

    partial void OnSelectedFilterChanged(string value) => Filter();
    partial void OnSearchChanged(string value) => Filter();
    partial void OnScriptsTextChanged(string? value) => OnPropertyChanged(nameof(HasScripts));

    // Saving any setting (e.g. "Delete confirmation") comes here too. The list is only rebuilt when the categories really changed: clearing it
    // makes every "Goes to" drop-down lose its choice (it reads as "Leave here", and the files would no longer move).
    private void RefreshChoices()
    {
        var names = new List<string> { LeaveHere };
        names.AddRange(app.CategoryNames());
        if (Choices.SequenceEqual(names)) return;
        foreach (var row in rows.Values) row.Updating = true;
        Choices.Clear();
        foreach (var c in names) Choices.Add(c);
        foreach (var row in rows.Values) { row.Updating = false; row.Refresh(); }   // every row shows the choice of its decision again
    }

    /// <summary>Takes the files from the result: those that go somewhere by the rules or by the user's choice.</summary>
    public void Reload()
    {
        var result = app.Result;
        HasResult = result != null;
        var wanted = result?.Decisions.Where(d => d.Category != null || d.OverrideCategory != null).ToList() ?? new();
        var keep = new HashSet<FileDecision>(wanted);
        foreach (var gone in rows.Keys.Where(k => !keep.Contains(k)).ToList()) rows.Remove(gone);
        foreach (var d in wanted)
            if (!rows.ContainsKey(d)) rows[d] = new FileRow(d, app, Choices, Play, UpdateSummary) { DeleteRequested = row => AskDeleteFile(app, row.Decision), MoveRequested = MoveOne };

        foreach (var row in rows.Values) row.Refresh();   // a file that was moved shows as moved
        var current = SelectedFilter;
        Filters.Clear();
        Filters.Add(AllCategories);
        foreach (var name in wanted.Select(d => d.Category ?? d.OverrideCategory!).Distinct().OrderBy(n => n, StringComparer.OrdinalIgnoreCase)) Filters.Add(name);
        SelectedFilter = Filters.Contains(current) ? current : AllCategories;
        OnPropertyChanged(nameof(HasRenames));
        ShowLeftovers();
        Filter();
    }

    private void Filter()
    {
        var query = rows.Values.AsEnumerable();
        if (SelectedFilter != AllCategories) query = query.Where(r => r.EffectiveName == SelectedFilter || r.RuleCategory == SelectedFilter);
        if (!string.IsNullOrWhiteSpace(Search))
            query = query.Where(r => r.FileName.Contains(Search.Trim(), StringComparison.OrdinalIgnoreCase) || r.Trigger.Contains(Search.Trim(), StringComparison.OrdinalIgnoreCase));
        var list = query.OrderBy(r => r.RuleCategory ?? "", StringComparer.OrdinalIgnoreCase).ThenBy(r => r.FileName, StringComparer.OrdinalIgnoreCase).ToList();
        Rows.Clear();
        foreach (var r in list) Rows.Add(r);
        HasRows = Rows.Count > 0;
        UpdateSummary();
    }

    private void UpdateSummary()
    {
        var all = rows.Keys.ToList();
        var pending = all.Where(d => !d.Done && d.EffectiveCategory != null).ToList();
        int folders = pending.Select(d => d.EffectiveCategory!).Distinct(StringComparer.OrdinalIgnoreCase).Count();
        int left = all.Count(d => !d.Done && d.EffectiveCategory == null);
        int moved = all.Count(d => d.Done);
        HasPending = pending.Count > 0;
        Summary = $"{Fmt.Plural(pending.Count, "file", "files")} will move to {Fmt.Plural(folders, "folder", "folders")}";
        var parts = new List<string>();
        if (left > 0) parts.Add($"{left:N0} left where they are");
        if (moved > 0) parts.Add($"{moved:N0} already moved");
        Detail = string.Join(", ", parts);
        main.UpdateBadges();
    }

    [RelayCommand]
    private void AskMove()
    {
        var pending = app.Pending.ToList();
        if (pending.Count == 0) return;
        int folders = pending.Select(d => d.EffectiveCategory!).Distinct(StringComparer.OrdinalIgnoreCase).Count();
        Ask($"Move {Fmt.Plural(pending.Count, "file", "files")}?",
            $"The files go into {Fmt.Plural(folders, "folder", "folders")} inside {app.Result!.Folder}. Nothing is overwritten: a file whose target exists already stays where it is."
            + (AlsoRename && pending.Any(d => d.RenameTo != null) ? " Files with a title or comment tag also get that as their name." : ""),
            "Move files", DoMove);
    }

    // The move icon of one row: moves just that file, at once (nothing is overwritten)
    private void MoveOne(FileRow row)
    {
        var report = app.MoveOne(row.Decision, AlsoRename);
        string text = report.Summary;
        if (report.Messages.Count > 0) text += " " + string.Join(" ", report.Messages.Take(2));
        ShowToast(text, report.Messages.Count > 0 ? 10000 : 5000);
    }

    private void DoMove()
    {
        var report = app.MoveNow(AlsoRename);
        string text = report.Summary;
        if (report.Messages.Count > 0) text += " " + string.Join(" ", report.Messages.Take(2)) + (report.Messages.Count > 2 ? $" (and {report.Messages.Count - 2} more)" : "");
        ShowToast(text, report.Messages.Count > 0 ? 12000 : 6000);
        // Folders were opened (Process folders): the Status window shows what is left in them
        if (app.LastLeftovers != null) OpenStatus();
    }

    [RelayCommand]
    private void CreateScripts()
    {
        var report = app.WriteScripts();
        if (!report.Any && report.UnmatchedReport == null) { ShowToast("There is nothing to write a script for."); return; }
        var parts = new List<string>();
        if (report.MoveScript != null) parts.Add($"{Path.GetFileName(report.MoveScript)} ({Fmt.Plural(report.Moves, "move", "moves")})");
        if (report.RenameScript != null) parts.Add($"{Path.GetFileName(report.RenameScript)} ({Fmt.Plural(report.Renames, "rename", "renames")})");
        if (report.UnmatchedReport != null) parts.Add(Path.GetFileName(report.UnmatchedReport));
        ScriptsFolder = app.Result!.Folder;
        string how = report.MoveScript == null ? "" : OperatingSystem.IsWindows()
            ? $" Run it with: powershell -ExecutionPolicy Bypass -File \"{report.MoveScript}\""
            : $" Run it with: sh \"{report.MoveScript}\"";
        ScriptsText = "Wrote " + string.Join(", ", parts) + " in " + ScriptsFolder + "." + how;
    }

    [RelayCommand] private async Task OpenScriptsFolder() { if (ScriptsFolder != null && OpenFolder != null) await OpenFolder(ScriptsFolder); }
    [RelayCommand] private void GoToScan() => main.ShowScan();
    [RelayCommand] private void GoToUnmatched() => main.ShowUnmatched();
}

/// <summary>One file in Review or Unmatched.</summary>
public partial class FileRow : ObservableObject
{
    private readonly AppController app;
    private readonly Action<string?> play;
    private readonly Action changed;
    private readonly ThumbLoader thumb;
    private readonly ThumbLoader[] parts = new ThumbLoader[Thumbnails.Parts];
    private bool settingChoice;
    /// <summary>True while the list of categories is rebuilt: the drop-down is emptied then, which is not a choice of the user.</summary>
    public bool Updating { get; set; }

    public FileDecision Decision { get; }
    public ObservableCollection<string> Choices { get; }

    public string FileName => Decision.FileName;
    /// <summary>For a file found inside a folder that was opened: where it is (shown under the name).</summary>
    public string FromFolderText => Decision.SubFolder.Length > 0 ? "in " + Decision.SubFolder : "";
    public bool HasFromFolder => Decision.SubFolder.Length > 0;
    public string Trigger => Decision.Trigger;
    public string? RuleCategory => Decision.Category;
    public bool IsVideo => Decision.IsVideo;
    public string Title => Decision.Title;
    public string Comment => Decision.Comment;
    public bool HasTitle => Decision.Title.Length > 0;
    public bool HasComment => Decision.Comment.Length > 0;
    public string RenameText => Decision.RenameTo == null ? "" : $"{Decision.RenameSource} tag: {Decision.RenameTo}";
    public bool HasRename => Decision.RenameTo != null;
    public List<string> TextLines => Decision.OcrLines;
    public bool HasText => Decision.OcrLines.Count > 0;
    public string SourceText => Decision.Source switch
    {
        MatchSource.Filename => "File name",
        MatchSource.Title => "Title tag",
        MatchSource.Comment => "Comment tag",
        MatchSource.Ocr => "Text in the video",
        MatchSource.Corrupt => "Can't be read",
        MatchSource.Folder => "Folder name",
        _ => Decision.OverrideCategory != null ? "Your choice" : "",
    };

    public Bitmap? Thumb => thumb.Bitmap;

    // Review shows three pictures spread over the video (each is made when its row first shows)
    public Bitmap? Thumb1 => parts[0].Bitmap;
    public Bitmap? Thumb2 => parts[1].Bitmap;
    public Bitmap? Thumb3 => parts[2].Bitmap;

    [ObservableProperty] private string choice = "";
    [ObservableProperty] private bool isDone;
    [ObservableProperty] private bool isLeft;
    /// <summary>The file goes somewhere and has not been moved yet: the move icon works.</summary>
    [ObservableProperty] private bool canMove;

    /// <summary>The category the file goes to now (empty: it stays).</summary>
    public string EffectiveName => Decision.EffectiveCategory ?? "";

    public FileRow(FileDecision d, AppController app, ObservableCollection<string> choices, Action<string?> play, Action changed)
    {
        Decision = d; this.app = app; Choices = choices; this.play = play; this.changed = changed;
        thumb = new ThumbLoader(() => app.ThumbOf(d), () => OnPropertyChanged(nameof(Thumb)));
        parts[0] = new ThumbLoader(() => app.ThumbOf(d, 0), () => OnPropertyChanged(nameof(Thumb1)));
        parts[1] = new ThumbLoader(() => app.ThumbOf(d, 1), () => OnPropertyChanged(nameof(Thumb2)));
        parts[2] = new ThumbLoader(() => app.ThumbOf(d, 2), () => OnPropertyChanged(nameof(Thumb3)));
        Refresh();
    }

    /// <summary>Shows the decision as it is now.</summary>
    public void Refresh()
    {
        settingChoice = true;
        Choice = Decision.Choice;
        settingChoice = false;
        IsDone = Decision.Done;
        IsLeft = Decision.EffectiveCategory == null;
        CanMove = !Decision.Done && Decision.EffectiveCategory != null;
        OnPropertyChanged(nameof(EffectiveName));
        OnPropertyChanged(nameof(SourceText));
    }

    // The user sends the file elsewhere, or leaves it
    partial void OnChoiceChanged(string value)
    {
        if (settingChoice || Updating || value == null) return;
        Decision.Choose(value);
        IsLeft = Decision.EffectiveCategory == null;
        CanMove = !Decision.Done && Decision.EffectiveCategory != null;
        OnPropertyChanged(nameof(EffectiveName));
        OnPropertyChanged(nameof(SourceText));
        changed();
    }

    [RelayCommand] private void PlayFile() => play(app.PathOf(Decision));

    /// <summary>Set by the screen: asks before deleting this file.</summary>
    public Action<FileRow>? DeleteRequested { get; set; }
    /// <summary>Set by the screen: moves this one file.</summary>
    public Action<FileRow>? MoveRequested { get; set; }
    [RelayCommand] private void Move() => MoveRequested?.Invoke(this);
    [RelayCommand] private void Delete() => DeleteRequested?.Invoke(this);
}
