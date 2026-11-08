using System.Collections.ObjectModel;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace FileOrganizer.Desktop.ViewModels;

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

    public bool HasScripts => ScriptsText != null;
    public bool HasRenames => app.Result?.Decisions.Any(d => d.RenameTo != null) ?? false;

    public ReviewViewModel(AppController app, MainViewModel main)
    {
        this.app = app;
        this.main = main;
        app.ResultChanged += () => Dispatcher.UIThread.Post(Reload);
        app.SettingsChanged += () => Dispatcher.UIThread.Post(RefreshChoices);
        RefreshChoices();
        Reload();
    }

    partial void OnSelectedFilterChanged(string value) => Filter();
    partial void OnSearchChanged(string value) => Filter();
    partial void OnScriptsTextChanged(string? value) => OnPropertyChanged(nameof(HasScripts));

    private void RefreshChoices()
    {
        Choices.Clear();
        Choices.Add(LeaveHere);
        foreach (var c in app.CategoryNames()) Choices.Add(c);
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
            if (!rows.ContainsKey(d)) rows[d] = new FileRow(d, app, Choices, Play, UpdateSummary);

        var current = SelectedFilter;
        Filters.Clear();
        Filters.Add(AllCategories);
        foreach (var name in wanted.Select(d => d.Category ?? d.OverrideCategory!).Distinct().OrderBy(n => n, StringComparer.OrdinalIgnoreCase)) Filters.Add(name);
        SelectedFilter = Filters.Contains(current) ? current : AllCategories;
        OnPropertyChanged(nameof(HasRenames));
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

    private void DoMove()
    {
        var report = app.MoveNow(AlsoRename);
        string text = report.Summary;
        if (report.Messages.Count > 0) text += " " + string.Join(" ", report.Messages.Take(2)) + (report.Messages.Count > 2 ? $" (and {report.Messages.Count - 2} more)" : "");
        ShowToast(text, report.Messages.Count > 0 ? 12000 : 6000);
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

    public FileDecision Decision { get; }
    public ObservableCollection<string> Choices { get; }

    public string FileName => Decision.FileName;
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
        OnPropertyChanged(nameof(EffectiveName));
        OnPropertyChanged(nameof(SourceText));
    }

    // The user sends the file elsewhere, or leaves it
    partial void OnChoiceChanged(string value)
    {
        if (settingChoice || value == null) return;
        Decision.Choose(value);
        IsLeft = Decision.EffectiveCategory == null;
        OnPropertyChanged(nameof(EffectiveName));
        OnPropertyChanged(nameof(SourceText));
        changed();
    }

    [RelayCommand] private void PlayFile() => play(app.PathOf(Decision));
}
