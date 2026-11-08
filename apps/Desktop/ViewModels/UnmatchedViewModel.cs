using System.Collections.ObjectModel;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace FileOrganizer.Desktop.ViewModels;

/// <summary>A piece of text found for a file (a tag or text in the video) that can be made into a rule.</summary>
public sealed record TextLine(string Label, string Text, IRelayCommand Command);

/// <summary>The files nothing matched, with what was found in them: click a text to make a rule from it.</summary>
public partial class UnmatchedViewModel : ViewModelBase
{
    private readonly AppController app;
    private readonly MainViewModel main;
    private readonly Dictionary<FileDecision, UnmatchedRow> rows = new();

    public ObservableCollection<UnmatchedRow> Rows { get; } = new();
    public ObservableCollection<string> Choices { get; } = new();

    [ObservableProperty] private bool hasResult;
    [ObservableProperty] private bool hasRows;
    [ObservableProperty] private string summary = "";

    // The "make a rule" dialog
    public ObservableCollection<string> RuleCategories { get; } = new();
    public string[] RuleTypes { get; } = { "Contains", "Regex" };
    [ObservableProperty] private bool ruleVisible;
    [ObservableProperty] private string ruleText = "";
    [ObservableProperty] private string ruleType = "Contains";
    [ObservableProperty] private string? ruleCategory;
    [ObservableProperty] private string newCategory = "";
    [ObservableProperty] private string rulePreview = "";
    [ObservableProperty] private bool rulePreviewBad;
    [ObservableProperty] private bool canAddRule;
    /// <summary>A new category was typed: it wins over the one chosen in the list.</summary>
    public bool HasNewCategory => (NewCategory ?? "").Trim().Length > 0;

    public UnmatchedViewModel(AppController app, MainViewModel main)
    {
        this.app = app;
        this.main = main;
        app.ResultChanged += () => Dispatcher.UIThread.Post(Reload);
        app.SettingsChanged += () => Dispatcher.UIThread.Post(RefreshChoices);
        RefreshChoices();
        Reload();
    }

    partial void OnRuleTextChanged(string value) => UpdatePreview();
    partial void OnRuleTypeChanged(string value) => UpdatePreview();
    partial void OnRuleCategoryChanged(string? value) => UpdatePreview();
    partial void OnNewCategoryChanged(string value) { OnPropertyChanged(nameof(HasNewCategory)); UpdatePreview(); }

    private void RefreshChoices()
    {
        Choices.Clear();
        Choices.Add(ReviewViewModel.LeaveHere);
        RuleCategories.Clear();
        foreach (var c in app.CategoryNames()) { Choices.Add(c); RuleCategories.Add(c); }
    }

    /// <summary>Takes the files nothing has put anywhere (by the rules or by the user's choice).</summary>
    public void Reload()
    {
        var result = app.Result;
        HasResult = result != null;
        var wanted = result?.Decisions.Where(d => d.Category == null && d.OverrideCategory == null).ToList() ?? new();
        var keep = new HashSet<FileDecision>(wanted);
        foreach (var gone in rows.Keys.Where(k => !keep.Contains(k)).ToList()) rows.Remove(gone);
        foreach (var d in wanted)
            if (!rows.ContainsKey(d)) rows[d] = new UnmatchedRow(d, app, Choices, p => Play(p), OnAssigned, OpenRule) { DeleteRequested = row => AskDeleteFile(app, row.Decision) };

        Rows.Clear();
        foreach (var r in wanted.Select(d => rows[d])) Rows.Add(r);
        HasRows = Rows.Count > 0;
        Summary = !HasResult ? "" : HasRows
            ? $"{Fmt.Plural(Rows.Count, "file", "files")} matched nothing. Click a text to make a rule from it, or choose a folder for the file."
            : "Every file has a category.";
        main.UpdateBadges();
    }

    // A file was given a folder by hand: it leaves this list (and shows up in Review)
    private void OnAssigned() => Dispatcher.UIThread.Post(() => { Reload(); });

    // ---------- Make a rule ----------

    private void OpenRule(string text)
    {
        RuleText = text;
        RuleType = "Contains";
        NewCategory = "";
        RuleCategory = RuleCategories.FirstOrDefault();
        RuleVisible = true;
        UpdatePreview();
    }

    // How many of the files would this rule catch? (From what was read: the name, the tags and the video text.)
    private void UpdatePreview()
    {
        string text = (RuleText ?? "").Trim();
        string typed = (NewCategory ?? "").Trim();
        string category = typed.Length > 0 ? typed : (RuleCategory ?? "");
        var rule = new Rule(RuleType, text);
        var error = RuleMatcher.Validate(rule);
        RulePreviewBad = error != null;
        CanAddRule = error == null && category.Length > 0 && FolderNames.IsSafe(category);
        if (error != null) { RulePreview = text.Length == 0 ? "Type the text to look for." : error[..1].ToUpper() + error[1..]; return; }

        var trial = new Settings { Categories = { new Category { Name = "x", IgnoreCategoryInContent = true, Rules = { rule } } } };
        int hits = rows.Keys.Count(d => RuleMatcher.Match(d.FileName, trial) != null || RuleMatcher.Match(d.Title, trial) != null
            || RuleMatcher.Match(d.Comment, trial) != null || d.OcrLines.Any(l => RuleMatcher.Match(l, trial) != null));
        RulePreview = category.Length == 0 ? "Choose a category, or name a new one."
            : !FolderNames.IsSafe(category) ? $"\"{category}\" can't be a folder name."
            : $"Would match {hits:N0} of the {rows.Count:N0} files here, and send them to {category}{(HasNewCategory ? " (a new category)" : "")}.";
    }

    [RelayCommand]
    private void AddRule()
    {
        string typed = (NewCategory ?? "").Trim();
        string category = typed.Length > 0 ? typed : (RuleCategory ?? "");
        string message = app.AddRule(category, RuleType, RuleText, out bool ok);
        ShowToast(message, ok ? 5000 : 9000);
        if (ok) RuleVisible = false;
    }

    [RelayCommand] private void CancelRule() => RuleVisible = false;
    [RelayCommand] private void GoToScan() => main.ShowScan();
    [RelayCommand] private void GoToReview() => main.ShowReview();
}

/// <summary>One unmatched file.</summary>
public sealed class UnmatchedRow : FileRow
{
    public List<TextLine> Lines { get; }

    public UnmatchedRow(FileDecision d, AppController app, ObservableCollection<string> choices, Action<string?> play, Action assigned, Action<string> makeRule)
        : base(d, app, choices, play, assigned)
    {
        Lines = new List<TextLine>();
        TextLine Make(string label, string text) => new(label, text, new RelayCommand(() => makeRule(text)));
        if (d.Title.Length > 0) Lines.Add(Make("Title tag", d.Title));
        if (d.Comment.Length > 0) Lines.Add(Make("Comment tag", d.Comment));
        // Lines without a single letter (time codes, numbers) are noise as a starting point for a rule
        foreach (var line in OcrText.Distinct(d.OcrLines).Where(l => l.Any(char.IsLetter))) Lines.Add(Make("", line));
    }

    public bool HasLines => Lines.Count > 0;
}
