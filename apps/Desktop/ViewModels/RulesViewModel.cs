using System.Collections.ObjectModel;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace FileOrganizer.Desktop.ViewModels;

public partial class RuleItem : ObservableObject
{
    private readonly Action<RuleItem> remove;
    private readonly Action edited;

    [ObservableProperty] private string type;
    [ObservableProperty] private string value;
    [ObservableProperty] private string error = "";

    public bool HasError => Error.Length > 0;
    public string[] Types { get; } = { "Contains", "Regex" };

    public RuleItem(Rule rule, Action<RuleItem> remove, Action edited)
    {
        type = rule.Type; value = rule.Value;
        this.remove = remove; this.edited = edited;
        Validate();
    }

    partial void OnTypeChanged(string value) { Validate(); edited(); }
    partial void OnValueChanged(string value) { Validate(); edited(); }
    partial void OnErrorChanged(string value) => OnPropertyChanged(nameof(HasError));

    private void Validate()
    {
        var problem = RuleMatcher.Validate(ToRule());
        Error = problem == null ? "" : problem[..1].ToUpper() + problem[1..];
    }

    public Rule ToRule() => new(Type, Value ?? "");
    [RelayCommand] private void Remove() => remove(this);
}

public partial class CategoryItem : ObservableObject
{
    private readonly Action<RuleItem> remove;
    private readonly Action edited;

    [ObservableProperty] private string name;
    [ObservableProperty] private bool ignoreName;
    [ObservableProperty] private string filesText = "";

    public ObservableCollection<RuleItem> Rules { get; } = new();

    public CategoryItem(Category category, Action<RuleItem> remove, Action edited)
    {
        this.remove = remove; this.edited = edited;
        name = category.Name; ignoreName = category.IgnoreCategoryInContent;
        foreach (var r in category.Rules) Rules.Add(new RuleItem(r, remove, edited));
    }

    public bool NameIsRule => !IgnoreName;
    partial void OnNameChanged(string value) => edited();
    partial void OnIgnoreNameChanged(bool value) { OnPropertyChanged(nameof(NameIsRule)); edited(); }

    [RelayCommand]
    private void AddRule()
    {
        Rules.Add(new RuleItem(new Rule("Contains", ""), remove, edited));
        edited();
    }

    public Category ToCategory() => new() { Name = (Name ?? "").Trim(), IgnoreCategoryInContent = IgnoreName, Rules = Rules.Select(r => r.ToRule()).ToList() };
}

/// <summary>The categories and their rules. The order of the list is the priority: the first match wins.</summary>
public partial class RulesViewModel : ViewModelBase
{
    private readonly AppController app;
    private readonly MainViewModel main;
    private bool loading;

    public ObservableCollection<CategoryItem> Categories { get; } = new();
    public string[] RuleTypes { get; } = { "Contains", "Regex" };

    [ObservableProperty] private CategoryItem? selected;
    [ObservableProperty] private bool isEdited;
    [ObservableProperty] private bool hasCategories;
    [ObservableProperty] private string problems = "";
    [ObservableProperty] private string tryText = "";
    [ObservableProperty] private string tryResult = "";
    [ObservableProperty] private bool tryMatches;
    [ObservableProperty] private string? settingsError;

    public bool HasProblems => Problems.Length > 0;
    public bool HasSelected => Selected != null;
    public bool HasSettingsError => SettingsError != null;

    public RulesViewModel(AppController app, MainViewModel main)
    {
        this.app = app;
        this.main = main;
        SettingsError = app.SettingsError;
        app.SettingsChanged += () => Dispatcher.UIThread.Post(() => { if (!IsEdited) Load(); });
        app.ResultChanged += () => Dispatcher.UIThread.Post(UpdateCounts);
        Load();
    }

    partial void OnProblemsChanged(string value) => OnPropertyChanged(nameof(HasProblems));
    partial void OnSelectedChanged(CategoryItem? value) => OnPropertyChanged(nameof(HasSelected));
    partial void OnSettingsErrorChanged(string? value) => OnPropertyChanged(nameof(HasSettingsError));
    partial void OnTryTextChanged(string value) => Try();

    private void Load()
    {
        loading = true;
        var keep = Selected?.Name;
        Categories.Clear();
        foreach (var c in app.Settings.Categories) Categories.Add(new CategoryItem(c, RemoveRule, Edited));
        Selected = Categories.FirstOrDefault(c => c.Name == keep) ?? Categories.FirstOrDefault();
        loading = false;
        IsEdited = false;
        Problems = "";
        HasCategories = Categories.Count > 0;
        UpdateCounts();
        Try();
    }

    private void Edited()
    {
        if (loading) return;
        IsEdited = true;
        Try();
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
            if (c.Rules.Remove(rule)) { Edited(); return; }
    }

    [RelayCommand]
    private void AddCategory()
    {
        string name = "New category";
        for (int i = 2; Categories.Any(c => c.Name.Equals(name, StringComparison.OrdinalIgnoreCase)); i++) name = $"New category {i}";
        var item = new CategoryItem(new Category { Name = name, IgnoreCategoryInContent = true, Rules = { new Rule("Contains", "") } }, RemoveRule, Edited);
        Categories.Add(item);
        Selected = item;
        HasCategories = true;
        Edited();
    }

    [RelayCommand]
    private void DeleteCategory()
    {
        if (Selected == null) return;
        int at = Categories.IndexOf(Selected);
        Categories.Remove(Selected);
        Selected = Categories.Count == 0 ? null : Categories[Math.Min(at, Categories.Count - 1)];
        HasCategories = Categories.Count > 0;
        Edited();
    }

    [RelayCommand]
    private void MoveUp()
    {
        if (Selected == null) return;
        int at = Categories.IndexOf(Selected);
        if (at > 0) { var s = Selected; Categories.Move(at, at - 1); Selected = s; Edited(); }
    }

    [RelayCommand]
    private void MoveDown()
    {
        if (Selected == null) return;
        int at = Categories.IndexOf(Selected);
        if (at < Categories.Count - 1) { var s = Selected; Categories.Move(at, at + 1); Selected = s; Edited(); }
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
