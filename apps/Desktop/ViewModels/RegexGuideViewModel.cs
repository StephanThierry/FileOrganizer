using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace FileOrganizer.Desktop.ViewModels;

/// <summary>
/// The regex guide shown beside a Regex rule (on the Rules screen) or beside the "Make a rule" dialog: the patterns, and a box
/// that shows what the pattern being edited selects in any text. One of these per place it is shown.
/// </summary>
public partial class RegexGuideViewModel : ObservableObject
{
    private readonly Func<Settings> settings;

    /// <summary>Says whether the target still exists (a removed rule or term closes the guide). Null: it always does.</summary>
    public Func<IGuideTarget, bool>? StillValid { get; set; }

    public ObservableCollection<RegexTipItem> Tips { get; } = new();
    [ObservableProperty] private IGuideTarget? target;
    [ObservableProperty] private string text = "";
    [ObservableProperty] private string pattern = "";
    [ObservableProperty] private string message = "";
    [ObservableProperty] private bool bad;
    [ObservableProperty] private IReadOnlyList<TextPiece> pieces = Array.Empty<TextPiece>();

    public bool IsVisible => Target != null;

    /// <param name="settings">The rules as they are being edited: the patterns they use are counted for the tips.</param>
    public RegexGuideViewModel(Func<Settings> settings) => this.settings = settings;

    partial void OnTargetChanged(IGuideTarget? value) => OnPropertyChanged(nameof(IsVisible));
    partial void OnTextChanged(string value) => Update();

    /// <summary>Shows the guide for a pattern (a Regex rule, a Regex term, or the text of the dialog).</summary>
    public void Show(IGuideTarget pattern)
    {
        if (Target == null)
        {
            Tips.Clear();
            foreach (var t in RegexGuide.Tips(settings()))
                Tips.Add(new RegexTipItem(t.Pattern, t.Meaning, t.Example, t.UsedBy == 0 ? "" : $"used by {Fmt.Plural(t.UsedBy, "rule", "rules")}"));
        }
        Target = pattern;
        Update();
    }

    /// <summary>The test text again, with what the pattern selects now (call after the pattern was edited).</summary>
    public void Update()
    {
        if (Target == null) return;
        if (StillValid != null && !StillValid(Target)) { Target = null; return; }
        Pattern = Target.GuideValue;
        var (found, note) = RegexGuide.Highlight(Pattern, Text);
        Pieces = found ?? new List<TextPiece> { new(Text ?? "", false) };
        Message = note;
        Bad = found == null && Pattern.Length > 0;
    }

    [RelayCommand] public void Close() => Target = null;
}
