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
}

public partial class RulesView : UserControl
{
    public RulesView() { InitializeComponent(); }
}

public partial class SettingsView : UserControl
{
    public SettingsView() { InitializeComponent(); this.WireOpenFolder(); }
}
