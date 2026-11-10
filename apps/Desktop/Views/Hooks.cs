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
    public RulesView()
    {
        InitializeComponent();
        // Type-ahead in the category list: a letter or a number typed with the list selected goes to the first category that starts with it
        this.FindControl<ListBox>("CategoryList")?.AddHandler(Avalonia.Input.InputElement.TextInputEvent, (_, e) =>
        {
            if (DataContext is RulesViewModel vm && !string.IsNullOrEmpty(e.Text) && !char.IsControl(e.Text[0]))
            {
                vm.TypeAhead(e.Text);
                e.Handled = true;
            }
        }, Avalonia.Interactivity.RoutingStrategies.Tunnel);
        DataContextChanged += (_, _) =>
        {
            // After Add the cursor goes to the name box (otherwise what is typed is taken by the list as a search for a category)
            if (DataContext is RulesViewModel vm)
                vm.NameFocusRequested += () => Avalonia.Threading.Dispatcher.UIThread.Post(() =>
                {
                    var box = this.FindControl<TextBox>("NameBox");
                    box?.Focus();
                    box?.SelectAll();
                }, Avalonia.Threading.DispatcherPriority.Background);
        };
    }

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
