using Avalonia.Controls;
using FileOrganizer.Desktop.ViewModels;

namespace FileOrganizer.Desktop.Views;

/// <summary>What a view gives its view model: opening a folder in the file manager.</summary>
public static class Hooks
{
    public static void WireOpenFolder(this UserControl view)
    {
        view.DataContextChanged += (_, _) =>
        {
            if (view.DataContext is ViewModelBase vm)
                vm.OpenFolder = async folder =>
                {
                    var top = TopLevel.GetTopLevel(view);
                    if (top != null) await top.Launcher.LaunchUriAsync(new Uri(Path.GetFullPath(folder)));
                };
        };
    }
}

public partial class ReviewView : UserControl
{
    public ReviewView() { InitializeComponent(); this.WireOpenFolder(); }
}

public partial class UnmatchedView : UserControl
{
    public UnmatchedView() { InitializeComponent(); this.WireOpenFolder(); }

    // Clicking into the text of a Regex rule, or of a Regex term, shows the regex guide beside the dialog
    private void RuleTextGotFocus(object? sender, Avalonia.Input.FocusChangedEventArgs e) => (DataContext as UnmatchedViewModel)?.ShowRegexGuide();

    private void TermGotFocus(object? sender, Avalonia.Input.FocusChangedEventArgs e)
    {
        if ((sender as Control)?.DataContext is TermItem term) term.Activate();
    }
}

public partial class RulesView : UserControl
{
    public RulesView() { InitializeComponent(); }

    // Clicking in the text of a Regex rule shows the regex guide beside it
    private void RuleGotFocus(object? sender, Avalonia.Input.FocusChangedEventArgs e)
    {
        if ((sender as Control)?.DataContext is RuleItem rule) rule.Activate();
    }

    // The same for a Regex term of a rule with several terms
    private void TermGotFocus(object? sender, Avalonia.Input.FocusChangedEventArgs e)
    {
        if ((sender as Control)?.DataContext is TermItem term) term.Activate();
    }
}

public partial class RegexGuidePanel : UserControl
{
    public RegexGuidePanel() { InitializeComponent(); }
}

public partial class SettingsView : UserControl
{
    public SettingsView() { InitializeComponent(); this.WireOpenFolder(); }
}
