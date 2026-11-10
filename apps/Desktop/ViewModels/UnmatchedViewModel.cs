using System.Collections.ObjectModel;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace FileOrganizer.Desktop.ViewModels;

/// <summary>
/// A piece of text found for a file (a tag or text in the video) that can be made into a rule: click it for a rule from this text
/// alone, or tick several to make one rule that needs all of them.
/// </summary>
public sealed partial class TextLine : ObservableObject
{
    private readonly Action selectionChanged;

    public string Label { get; }
    public string Text { get; }
    public IRelayCommand Command { get; }
    [ObservableProperty] private bool isSelected;

    /// <summary>Every different way the OCR read this text (it reads the same text a little differently in each frame).</summary>
    public IReadOnlyList<OcrVariant> Variants { get; }
    public bool HasVariants => Variants.Count > 1;
    public string VariantsBadge => HasVariants ? $"{Variants.Count} variants" : "";
    public string VariantsTip => HasVariants
        ? "Read as:\n" + string.Join("\n", Variants.Select(v => v.Count > 1 ? $"{v.Text}   ({v.Count}x)" : v.Text)) + "\n\nClick to see which rule fits them all."
        : "Make a rule from this text";

    public TextLine(string label, string text, IRelayCommand command, Action selectionChanged, IReadOnlyList<OcrVariant>? variants = null)
    {
        Label = label; Text = text; Command = command; this.selectionChanged = selectionChanged;
        Variants = variants ?? new[] { new OcrVariant(text, 1) };
    }

    partial void OnIsSelectedChanged(bool value) => selectionChanged();
}

/// <summary>One way a text was read, in the "Make a rule" dialog: does the rule being made catch it?</summary>
public sealed partial class VariantItem : ObservableObject
{
    public string Text { get; }
    public string CountText { get; }
    [ObservableProperty] private bool caught;
    public string Mark => Caught ? "\u2713" : "\u2717";
    public IRelayCommand UseCommand { get; }

    public VariantItem(OcrVariant variant, Action<string> use)
    {
        Text = variant.Text;
        CountText = variant.Count > 1 ? $"{variant.Count}x" : "";
        UseCommand = new RelayCommand(() => use(variant.Text));
    }

    partial void OnCaughtChanged(bool value) => OnPropertyChanged(nameof(Mark));
}

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
    public string[] RuleTypes { get; } = { "Contains", "Regex", RuleItem.AllLabel };
    /// <summary>The terms of a rule with several terms ("All of"): all of them must be found.</summary>
    public ObservableCollection<TermItem> RuleTerms { get; } = new();
    public bool IsAllRule => RuleType == RuleItem.AllLabel;
    public bool IsSingleRule => !IsAllRule;

    /// <summary>The different readings of the text(s) the rule is made from: which of them would the rule being made catch?</summary>
    public ObservableCollection<VariantItem> RuleVariants { get; } = new();
    [ObservableProperty] private string variantSummary = "";
    /// <summary>Shown when there is more than one reading to compare (for "All of", whenever there are texts).</summary>
    public bool ShowVariants => IsAllRule ? RuleVariants.Count > 0 : RuleVariants.Count > 1;

    private void SetVariants(IEnumerable<OcrVariant> variants)
    {
        RuleVariants.Clear();
        foreach (var v in variants.GroupBy(v => v.Text, StringComparer.OrdinalIgnoreCase).Select(g => g.First()))
            RuleVariants.Add(new VariantItem(v, text => { if (IsSingleRule) RuleText = text; }));
        OnPropertyChanged(nameof(ShowVariants));
    }

    // Which readings does the rule catch? (A rule with several terms: which does one of its terms catch.)
    private void UpdateVariants(Rule rule, bool valid)
    {
        int caught = 0;
        var trial = valid && !IsAllRule ? new Settings { Categories = { new Category { Name = "x", IgnoreCategoryInContent = true, Rules = { rule } } } } : null;
        foreach (var v in RuleVariants)
        {
            bool hit = valid && (IsAllRule
                ? RuleTerms.Any(t => RuleMatcher.AllOfMatches(Rule.AllOf(new[] { t.ToTerm() }), new[] { v.Text }))
                : RuleMatcher.Match(v.Text, trial!) != null);
            v.Caught = hit;
            if (hit) caught++;
        }
        VariantSummary = !ShowVariants || !valid ? ""
            : IsAllRule ? $"The terms catch {caught} of the {RuleVariants.Count} readings."
            : $"This rule catches {caught} of the {RuleVariants.Count} readings.";
    }
    [ObservableProperty] private bool ruleVisible;
    [ObservableProperty] private string ruleText = "";
    [ObservableProperty] private string ruleType = "Contains";
    [ObservableProperty] private string? ruleCategory;
    [ObservableProperty] private string newCategory = "";
    [ObservableProperty] private string rulePreview = "";
    [ObservableProperty] private bool rulePreviewBad;
    [ObservableProperty] private bool canAddRule;
    /// <summary>The first entry of the category list: the rule goes to a new category, named in the box below the list.</summary>
    public const string NewCategoryChoice = "+ A new category (type its name below)";
    /// <summary>The list shows that a new category is being made (typing a name selects it, choosing a real category clears the name).</summary>
    public bool HasNewCategory => RuleCategory == NewCategoryChoice;

    // The regex guide, to the right of the dialog while a pattern is being made (Regex chosen)
    public RegexGuideViewModel Guide { get; }
    private readonly DialogPattern dialogPattern;

    /// <summary>The text of the dialog as a pattern the guide can try.</summary>
    private sealed class DialogPattern(UnmatchedViewModel owner) : IGuideTarget
    {
        public string GuideValue => owner.RuleText ?? "";
    }

    /// <summary>Shows the guide when the rule is a Regex (called when the text box is clicked into).</summary>
    public void ShowRegexGuide() { if (RuleType == "Regex") Guide.Show(dialogPattern); }

    partial void OnRuleVisibleChanged(bool value) { if (!value) Guide.Close(); }

    public UnmatchedViewModel(AppController app, MainViewModel main)
    {
        this.app = app;
        this.main = main;
        dialogPattern = new DialogPattern(this);
        Guide = new RegexGuideViewModel(() => app.Settings)
        {
            // Open while the dialog is: for its own text when that is a Regex, for a term of a rule with several terms
            StillValid = t => RuleVisible && (ReferenceEquals(t, dialogPattern) ? RuleType == "Regex" : RuleTerms.Any(x => ReferenceEquals(x, t))),
        };
        BindConfirmDelete(app);
        app.ResultChanged += () => Dispatcher.UIThread.Post(Reload);
        app.SettingsChanged += () => Dispatcher.UIThread.Post(RefreshChoices);
        RefreshChoices();
        Reload();
    }

    partial void OnRuleTextChanged(string value) => UpdatePreview();
    partial void OnRuleTypeChanged(string value)
    {
        OnPropertyChanged(nameof(IsAllRule)); OnPropertyChanged(nameof(IsSingleRule));
        // Becoming a rule with several terms: the text so far is the first term, and there is room for a second
        OnPropertyChanged(nameof(ShowVariants));
        if (IsAllRule && RuleTerms.Count == 0)
        {
            RuleTerms.Add(NewTerm(RuleText ?? ""));
            RuleTerms.Add(NewTerm(""));
        }
        // Choosing Regex shows the guide beside the dialog; another type puts it away (a Regex term of "All of" shows it again)
        if (value == "Regex") Guide.Show(dialogPattern); else Guide.Close();
        UpdatePreview();
    }

    private TermItem NewTerm(string text) => new(new RuleTerm("Contains", text), t => { RuleTerms.Remove(t); UpdatePreview(); }, UpdatePreview, t => Guide.Show(t));

    [RelayCommand]
    private void AddTerm()
    {
        RuleTerms.Add(NewTerm(""));
        UpdatePreview();
    }
    partial void OnRuleCategoryChanged(string? value)
    {
        OnPropertyChanged(nameof(HasNewCategory));
        if (value != NewCategoryChoice && NewCategory.Length > 0) NewCategory = "";   // an existing category was chosen: forget the typed name
        UpdatePreview();
    }

    partial void OnNewCategoryChanged(string value)
    {
        if (value.Trim().Length > 0) RuleCategory = NewCategoryChoice;   // typing a name means "a new category"
        UpdatePreview();
    }

    private void RefreshChoices()
    {
        // Saving any setting (e.g. "Delete confirmation") comes here too: only rebuild the lists when the categories changed
        var names = new List<string> { ReviewViewModel.LeaveHere };
        names.AddRange(app.CategoryNames());
        if (Choices.SequenceEqual(names)) return;
        Choices.Clear();
        Choices.Add(ReviewViewModel.LeaveHere);
        RuleCategories.Clear();
        RuleCategories.Add(NewCategoryChoice);
        foreach (var c in app.CategoryNames()) { Choices.Add(c); RuleCategories.Add(c); }
    }

    /// <summary>Takes the files nothing has put anywhere (by the rules or by the user's choice).</summary>
    public void Reload()
    {
        var result = app.Result;
        HasResult = result != null;
        var wanted = result?.Decisions.Where(d => d.IsUnmatched).ToList() ?? new();
        var keep = new HashSet<FileDecision>(wanted);
        foreach (var gone in rows.Keys.Where(k => !keep.Contains(k)).ToList()) rows.Remove(gone);
        foreach (var d in wanted)
            if (!rows.ContainsKey(d)) rows[d] = new UnmatchedRow(d, app, Choices, p => Play(p), OnAssigned, OpenRule, OpenAllRule) { DeleteRequested = row => AskDeleteFile(app, row.Decision), MoveRequested = MoveOne };

        Rows.Clear();
        foreach (var r in wanted.Select(d => rows[d])) Rows.Add(r);
        HasRows = Rows.Count > 0;
        Summary = !HasResult ? "" : HasRows
            ? $"{Fmt.Plural(Rows.Count, "file", "files")} matched nothing. Click a text to make a rule from it, or choose a folder for the file."
            : "Every file has a category.";
        main.UpdateBadges();
    }

    // A folder was chosen by hand: the file stays in this list until its move icon is clicked (Review lists it as waiting to move)
    private void OnAssigned() => Dispatcher.UIThread.Post(main.UpdateBadges);

    // The move icon of one row: moves just that file now (nothing is overwritten); it then leaves this list
    private void MoveOne(FileRow row)
    {
        var report = app.MoveOne(row.Decision, alsoRename: false);
        string text = report.Summary;
        if (report.Messages.Count > 0) text += " " + string.Join(" ", report.Messages.Take(2));
        ShowToast(text, report.Messages.Count > 0 ? 10000 : 5000);
    }

    // ---------- Make a rule ----------

    private void OpenRule(string text, IReadOnlyList<OcrVariant> variants, string fileName)
    {
        Guide.Close();
        Guide.Text = fileName;   // the regex guide's test box starts with the file name: the pattern can be tried on it at once
        RuleTerms.Clear();
        SetVariants(variants);
        RuleText = text;
        RuleType = "Contains";
        NewCategory = "";
        RuleCategory = RuleCategories.Skip(1).FirstOrDefault() ?? NewCategoryChoice;   // the first real category, or "a new category" when there is none
        RuleVisible = true;
        UpdatePreview();
    }

    // Several texts of one file were ticked: one rule that needs all of them (each can be edited, or made a pattern, here)
    private void OpenAllRule(IEnumerable<TextLine> lines, string fileName)
    {
        Guide.Close();
        Guide.Text = fileName;
        RuleTerms.Clear();
        var chosen = lines.ToList();
        SetVariants(chosen.SelectMany(l => l.Variants));
        foreach (var l in chosen) RuleTerms.Add(NewTerm(l.Text));
        RuleText = "";
        RuleType = RuleItem.AllLabel;
        NewCategory = "";
        RuleCategory = RuleCategories.Skip(1).FirstOrDefault() ?? NewCategoryChoice;
        RuleVisible = true;
        UpdatePreview();
    }

    // How many of the files would this rule catch? (From what was read: the name, the tags and the video text.)
    private void UpdatePreview()
    {
        Guide.Update();
        string text = (RuleText ?? "").Trim();
        string category = ChosenCategory();
        var rule = IsAllRule ? Rule.AllOf(RuleTerms.Select(t => t.ToTerm())) : new Rule(RuleType, text);
        var error = RuleMatcher.Validate(rule);
        UpdateVariants(rule, error == null);
        RulePreviewBad = error != null;
        CanAddRule = error == null && category.Length > 0 && FolderNames.IsSafe(category);
        if (error != null) { RulePreview = !IsAllRule && text.Length == 0 ? "Type the text to look for." : error[..1].ToUpper() + error[1..]; return; }

        var trial = new Settings { Categories = { new Category { Name = "x", IgnoreCategoryInContent = true, Rules = { rule } } } };
        // A rule with several terms looks at everything known about the file together; the others at one text at a time
        int hits = rows.Keys.Count(d => IsAllRule
            ? RuleMatcher.AllOfMatches(rule, new[] { d.FileName, d.Title, d.Comment }.Concat(d.OcrLines).ToList())
            : RuleMatcher.Match(d.FileName, trial) != null || RuleMatcher.Match(d.Title, trial) != null
              || RuleMatcher.Match(d.Comment, trial) != null || d.OcrLines.Any(l => RuleMatcher.Match(l, trial) != null));
        RulePreview = category.Length == 0 ? (HasNewCategory ? "Type the name of the new category (a new folder) in the box below the list." : "Choose a category, or name a new one.")
            : !FolderNames.IsSafe(category) ? $"\"{category}\" can't be a folder name."
            : $"Would match {hits:N0} of the {rows.Count:N0} files here, and send them to {category}{(HasNewCategory ? " (a new category)" : "")}.";
    }

    /// <summary>The category the rule goes to: the typed name when "a new category" is chosen, else the one chosen in the list.</summary>
    private string ChosenCategory() => HasNewCategory ? (NewCategory ?? "").Trim() : (RuleCategory ?? "");

    [RelayCommand]
    private void AddRule()
    {
        string category = ChosenCategory();
        string message = IsAllRule
            ? app.AddAllRule(category, RuleTerms.Select(t => t.ToTerm()), out bool ok)
            : app.AddRule(category, RuleType, RuleText, out ok);
        ShowToast(message, ok ? 5000 : 9000);
        if (ok) RuleVisible = false;
    }

    [RelayCommand] private void CancelRule() => RuleVisible = false;
    [RelayCommand] private void GoToScan() => main.ShowScan();
    [RelayCommand] private void GoToReview() => main.ShowReview();
}

/// <summary>One unmatched file.</summary>
public sealed partial class UnmatchedRow : FileRow
{
    public List<TextLine> Lines { get; }

    /// <summary>The texts ticked for one rule that needs all of them (two or more make that possible).</summary>
    public int SelectedCount => Lines.Count(l => l.IsSelected);
    public bool CanMakeAllRule => SelectedCount >= 2;
    public string MakeAllRuleText => $"Make one rule from the {SelectedCount} ticked texts (all must be found)";

    private readonly Action<IEnumerable<TextLine>, string> makeAllRule;
    private readonly string fileNameForGuide;

    /// <summary>The file name (without its extension) as a text to make a rule from, like the tags and the text found in the video.</summary>
    public IRelayCommand NameCommand { get; }

    [RelayCommand]
    private void MakeAllRule() => makeAllRule(Lines.Where(l => l.IsSelected).ToList(), fileNameForGuide);

    private void SelectionChanged()
    {
        OnPropertyChanged(nameof(SelectedCount)); OnPropertyChanged(nameof(CanMakeAllRule)); OnPropertyChanged(nameof(MakeAllRuleText));
    }

    public UnmatchedRow(FileDecision d, AppController app, ObservableCollection<string> choices, Action<string?> play, Action assigned, Action<string, IReadOnlyList<OcrVariant>, string> makeRule, Action<IEnumerable<TextLine>, string> makeAllRule)
        : base(d, app, choices, play, assigned)
    {
        this.makeAllRule = makeAllRule;
        fileNameForGuide = d.FileName;
        string nameText = System.IO.Path.GetFileNameWithoutExtension(d.FileName);
        NameCommand = new RelayCommand(() => makeRule(nameText, new[] { new OcrVariant(nameText, 1) }, d.FileName));
        Lines = new List<TextLine>();
        TextLine Make(string label, string text, IReadOnlyList<OcrVariant>? variants = null)
        {
            variants ??= new[] { new OcrVariant(text, 1) };
            return new(label, text, new RelayCommand(() => makeRule(text, variants, d.FileName)), SelectionChanged, variants);
        }
        if (d.Title.Length > 0) Lines.Add(Make("Title tag", d.Title));
        if (d.Comment.Length > 0) Lines.Add(Make("Comment tag", d.Comment));
        // Lines without a single letter (time codes, numbers) are noise as a starting point for a rule
        // One tag per text on screen, however many different ways the frames read it (all of them are kept, to choose a rule from)
        foreach (var finding in OcrText.Findings(d.OcrLines).Where(f => f.Best.Any(char.IsLetter))) Lines.Add(Make("", finding.Best, finding.Variants));
    }

    public bool HasLines => Lines.Count > 0;
}
