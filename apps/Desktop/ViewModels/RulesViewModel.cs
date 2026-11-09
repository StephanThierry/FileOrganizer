using System.Collections.ObjectModel;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace FileOrganizer.Desktop.ViewModels;

/// <summary>A tip in the regex guide as the list shows it.</summary>
public sealed record RegexTipItem(string Pattern, string Meaning, string Example, string UsedText);

/// <summary>Something with a pattern the regex guide can test: a Regex rule, or a Regex term of an "All of" rule.</summary>
public interface IGuideTarget { string GuideValue { get; } }

/// <summary>One term of an "All of" rule: text to find, or a pattern.</summary>
public partial class TermItem : ObservableObject, IGuideTarget
{
    private readonly Action<TermItem> remove;
    private readonly Action edited;
    private readonly Action<IGuideTarget> activated;

    [ObservableProperty] private string type;
    [ObservableProperty] private string value;

    public string[] Types { get; } = { "Contains", "Regex" };
    public bool IsRegex => Type == "Regex";
    public string GuideValue => Value ?? "";

    public TermItem(RuleTerm term, Action<TermItem> remove, Action edited, Action<IGuideTarget> activated)
    {
        type = term.IsRegex ? "Regex" : "Contains"; value = term.Value;
        this.remove = remove; this.edited = edited; this.activated = activated;
    }

    /// <summary>The user is working on this term (chose Regex, or clicked in its text): the regex guide shows for it.</summary>
    public void Activate() { if (IsRegex) activated(this); }

    partial void OnTypeChanged(string value) { OnPropertyChanged(nameof(IsRegex)); edited(); Activate(); }
    partial void OnValueChanged(string value) { OnPropertyChanged(nameof(GuideValue)); edited(); }

    public RuleTerm ToTerm() => new(Type, Value ?? "");
    [RelayCommand] private void Remove() => remove(this);
}

public partial class RuleItem : ObservableObject, IGuideTarget
{
    /// <summary>The name of the rule type with several terms, as the drop-down shows it (saved as "All").</summary>
    public const string AllLabel = "All of";

    private readonly Action<RuleItem> remove;
    private readonly Action edited;
    private readonly Action<IGuideTarget> activated;

    [ObservableProperty] private string type;
    [ObservableProperty] private string value;
    [ObservableProperty] private string error = "";

    public bool HasError => Error.Length > 0;
    public string[] Types { get; } = { "Contains", "Regex", AllLabel };
    public ObservableCollection<TermItem> Terms { get; } = new();

    public bool IsRegex => Type == "Regex";
    public bool IsAll => Type == AllLabel;
    public bool IsSingle => !IsAll;
    public string GuideValue => Value ?? "";

    public RuleItem(Rule rule, Action<RuleItem> remove, Action edited, Action<IGuideTarget> activated)
    {
        type = rule.IsAll ? AllLabel : rule.Type; value = rule.Value;
        this.remove = remove; this.edited = edited; this.activated = activated;
        if (rule.IsAll) foreach (var t in rule.Terms ?? new()) Terms.Add(NewTerm(t));
        Validate();
    }

    private TermItem NewTerm(RuleTerm t) => new(t, RemoveTerm, TermEdited, activated);
    private void TermEdited() { Validate(); edited(); }

    /// <summary>The user is working on this rule (chose Regex, or clicked in its text): the regex guide shows for it.</summary>
    public void Activate() { if (IsRegex) activated(this); }

    partial void OnTypeChanged(string value)
    {
        OnPropertyChanged(nameof(IsRegex)); OnPropertyChanged(nameof(IsAll)); OnPropertyChanged(nameof(IsSingle));
        // Becoming a rule with several terms: the text so far is the first term, and there is room for a second
        if (IsAll && Terms.Count == 0)
        {
            Terms.Add(NewTerm(new RuleTerm("Contains", Value ?? "")));
            Terms.Add(NewTerm(new RuleTerm("Contains", "")));
        }
        Validate(); edited(); Activate();
    }

    partial void OnValueChanged(string value) { OnPropertyChanged(nameof(GuideValue)); Validate(); edited(); }
    partial void OnErrorChanged(string value) => OnPropertyChanged(nameof(HasError));

    private void Validate()
    {
        var problem = RuleMatcher.Validate(ToRule());
        Error = problem == null ? "" : problem[..1].ToUpper() + problem[1..];
    }

    public bool Owns(IGuideTarget target) => ReferenceEquals(target, this) || Terms.Any(t => ReferenceEquals(t, target));

    public Rule ToRule() => IsAll ? Rule.AllOf(Terms.Select(t => t.ToTerm())) : new(Type, Value ?? "");

    [RelayCommand] private void Remove() => remove(this);

    [RelayCommand]
    private void AddTerm()
    {
        Terms.Add(NewTerm(new RuleTerm("Contains", "")));
        TermEdited();
    }

    private void RemoveTerm(TermItem term)
    {
        Terms.Remove(term);
        TermEdited();
    }
}

public partial class CategoryItem : ObservableObject
{
    private readonly Action<RuleItem> remove;
    private readonly Action edited;
    private readonly Action<IGuideTarget> activated;

    [ObservableProperty] private string name;
    [ObservableProperty] private bool ignoreName;
    [ObservableProperty] private string filesText = "";
    /// <summary>Where the category stands in the order the rules are checked (1 is checked first).</summary>
    [ObservableProperty] private int priority;

    /// <summary>The line under the name in the list.</summary>
    public string InfoText => $"Priority {Priority}" + (FilesText.Length > 0 ? " · " + FilesText : "");
    partial void OnPriorityChanged(int value) => OnPropertyChanged(nameof(InfoText));
    partial void OnFilesTextChanged(string value) => OnPropertyChanged(nameof(InfoText));

    public ObservableCollection<RuleItem> Rules { get; } = new();

    public CategoryItem(Category category, Action<RuleItem> remove, Action edited, Action<IGuideTarget> activated)
    {
        this.remove = remove; this.edited = edited; this.activated = activated;
        name = category.Name; ignoreName = category.IgnoreCategoryInContent;
        foreach (var r in category.Rules) Rules.Add(new RuleItem(r, remove, edited, activated));
    }

    public bool NameIsRule => !IgnoreName;
    partial void OnNameChanged(string value) => edited();
    partial void OnIgnoreNameChanged(bool value) { OnPropertyChanged(nameof(NameIsRule)); edited(); }

    [RelayCommand]
    private void AddRule()
    {
        Rules.Add(new RuleItem(new Rule("Contains", ""), remove, edited, activated));
        edited();
    }

    public bool Owns(IGuideTarget target) => Rules.Any(r => r.Owns(target));

    public Category ToCategory() => new() { Name = (Name ?? "").Trim(), IgnoreCategoryInContent = IgnoreName, Rules = Rules.Select(r => r.ToRule()).ToList() };
}

/// <summary>The categories and their rules. The order of the list is the priority: the first match wins.</summary>
public partial class RulesViewModel : ViewModelBase
{
    private readonly AppController app;
    private readonly MainViewModel main;
    private bool loading;

    /// <summary>The categories in priority order: the first one that matches wins. This is the order that is saved.</summary>
    public ObservableCollection<CategoryItem> Categories { get; } = new();
    /// <summary>The same categories for the list on screen: alphabetical, each showing its priority.</summary>
    public ObservableCollection<CategoryItem> SortedCategories { get; } = new();
    public string[] RuleTypes { get; } = { "Contains", "Regex" };

    [ObservableProperty] private CategoryItem? selected;
    [ObservableProperty] private bool isEdited;
    [ObservableProperty] private bool hasCategories;
    [ObservableProperty] private string problems = "";
    [ObservableProperty] private string tryText = "";
    [ObservableProperty] private string tryResult = "";
    [ObservableProperty] private bool tryMatches;
    [ObservableProperty] private string? settingsError;

    // The regex guide, beside the rule being edited
    public RegexGuideViewModel Guide { get; }

    public bool HasProblems => Problems.Length > 0;
    public bool HasSelected => Selected != null;
    public bool HasSettingsError => SettingsError != null;

    public RulesViewModel(AppController app, MainViewModel main)
    {
        this.app = app;
        this.main = main;
        Guide = new RegexGuideViewModel(ToSettings) { StillValid = t => Selected?.Owns(t) == true };   // a removed rule or term closes it
        SettingsError = app.SettingsError;
        app.SettingsChanged += () => Dispatcher.UIThread.Post(() => { if (!IsEdited) Load(); });
        app.ResultChanged += () => Dispatcher.UIThread.Post(UpdateCounts);
        Load();
    }

    partial void OnProblemsChanged(string value) => OnPropertyChanged(nameof(HasProblems));
    partial void OnSelectedChanged(CategoryItem? value) { OnPropertyChanged(nameof(HasSelected)); if (Guide.Target != null && value?.Owns(Guide.Target) != true) Guide.Close(); }
    partial void OnSettingsErrorChanged(string? value) => OnPropertyChanged(nameof(HasSettingsError));
    partial void OnTryTextChanged(string value) => Try();
    /// <summary>Shows the regex guide for a Regex rule or term (it counts the patterns of the rules as they are being edited).</summary>
    private void ShowGuide(IGuideTarget rule)
    {
        if (loading) return;
        Guide.Show(rule);
    }

    private void Load()
    {
        loading = true;
        var keep = Selected?.Name;
        Categories.Clear();
        SortedCategories.Clear();
        foreach (var c in app.Settings.Categories) Categories.Add(new CategoryItem(c, RemoveRule, Edited, ShowGuide));
        foreach (var c in Categories.OrderBy(c => c.Name, NameOrder)) SortedCategories.Add(c);
        UpdatePriorities();
        Selected = Categories.FirstOrDefault(c => c.Name == keep) ?? SortedCategories.FirstOrDefault();
        loading = false;
        IsEdited = false;
        Problems = "";
        HasCategories = Categories.Count > 0;
        UpdateCounts();
        Try();
    }

    private static readonly StringComparer NameOrder = StringComparer.CurrentCultureIgnoreCase;

    // The list on screen follows the names (moving items, so the selection stays)
    private void Resort()
    {
        var wanted = Categories.OrderBy(c => c.Name, NameOrder).ToList();
        for (int i = 0; i < wanted.Count; i++)
        {
            int at = SortedCategories.IndexOf(wanted[i]);
            if (at < 0) SortedCategories.Insert(i, wanted[i]);
            else if (at != i) SortedCategories.Move(at, i);
        }
    }

    private void UpdatePriorities()
    {
        for (int i = 0; i < Categories.Count; i++) Categories[i].Priority = i + 1;
    }

    private void Edited()
    {
        if (loading) return;
        Resort();
        IsEdited = true;
        Try();
        Guide.Update();
    }

    // How many files of the last scan each category has now
    private void UpdateCounts()
    {
        var counts = app.Result?.Decisions.Where(d => d.Category != null).GroupBy(d => d.Category!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.Count(), StringComparer.OrdinalIgnoreCase);
        foreach (var c in Categories)
            c.FilesText = counts == null ? "" : counts.TryGetValue(c.Name, out int n) ? Fmt.Plural(n, "file", "files") : "no files";
    }

    /// <summary>The settings as they are being edited (what Save would write).</summary>
    private Settings ToSettings()
    {
        var s = app.Settings.Clone();   // the options stay as they are
        s.Categories = Categories.Select(c => c.ToCategory()).ToList();
        return s;
    }

    // Which rule wins for a text (with the unsaved changes)?
    private void Try()
    {
        if (string.IsNullOrWhiteSpace(TryText)) { TryResult = ""; TryMatches = false; return; }
        var hits = RuleMatcher.Explain(TryText, ToSettings());
        TryMatches = hits.Count > 0;
        if (hits.Count == 0) { TryResult = "No category matches this text."; return; }
        TryResult = $"Goes to {hits[0].Category} ({hits[0].Trigger})."
            + (hits.Count > 1 ? " Also matched, but later in the list: " + string.Join(", ", hits.Skip(1).Select(h => h.Category).Distinct()) + "." : "");
    }

    private void RemoveRule(RuleItem rule)
    {
        foreach (var c in Categories)
            if (c.Rules.Remove(rule)) { if (Guide.Target != null && rule.Owns(Guide.Target)) Guide.Close(); Edited(); return; }
    }

    [RelayCommand]
    private void AddCategory()
    {
        string name = "New category";
        for (int i = 2; Categories.Any(c => c.Name.Equals(name, StringComparison.OrdinalIgnoreCase)); i++) name = $"New category {i}";
        var item = new CategoryItem(new Category { Name = name, IgnoreCategoryInContent = true, Rules = { new Rule("Contains", "") } }, RemoveRule, Edited, ShowGuide);
        Categories.Add(item);
        Resort();
        UpdatePriorities();
        Selected = item;
        HasCategories = true;
        Edited();
    }

    [RelayCommand]
    private void DeleteCategory()
    {
        if (Selected == null) return;
        int at = SortedCategories.IndexOf(Selected);
        Categories.Remove(Selected);
        SortedCategories.Remove(Selected);
        UpdatePriorities();
        Selected = SortedCategories.Count == 0 ? null : SortedCategories[Math.Min(at, SortedCategories.Count - 1)];
        HasCategories = Categories.Count > 0;
        Edited();
    }

    [RelayCommand]
    private void MoveUp()
    {
        if (Selected == null) return;
        int at = Categories.IndexOf(Selected);
        if (at > 0) { var s = Selected; Categories.Move(at, at - 1); UpdatePriorities(); Selected = s; Edited(); }
    }

    [RelayCommand]
    private void MoveDown()
    {
        if (Selected == null) return;
        int at = Categories.IndexOf(Selected);
        if (at < Categories.Count - 1) { var s = Selected; Categories.Move(at, at + 1); UpdatePriorities(); Selected = s; Edited(); }
    }

    [RelayCommand]
    private void Save()
    {
        var problemsFound = app.SaveSettings(ToSettings());
        if (problemsFound.Count > 0) { Problems = string.Join("\n", problemsFound); return; }
        Problems = "";
        SettingsError = null;
        IsEdited = false;
        ShowToast(app.Result != null ? "Saved. The files of the last scan were sorted again with the new rules." : "Saved.");
    }

    [RelayCommand] private void Revert() => Load();
    [RelayCommand] private void StartScan() => main.ShowScan();
}
