using Avalonia.Controls;
using FileOrganizer.Desktop.ViewModels;

namespace FileOrganizer.Desktop.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        if (!Design.IsDesignMode && Program.App != null) DataContext = new MainViewModel(Program.App);
        Opened += (_, _) => FitToScreen();
    }

    // The window starts wide, so no one has to resize it: on a small screen it is made to fit instead
    private void FitToScreen()
    {
        var screen = Screens.ScreenFromWindow(this) ?? Screens.Primary;
        if (screen == null) return;
        double scale = screen.Scaling > 0 ? screen.Scaling : 1;
        double maxWidth = screen.WorkingArea.Width / scale, maxHeight = screen.WorkingArea.Height / scale;
        if (Width > maxWidth) Width = Math.Max(MinWidth, maxWidth);
        if (Height > maxHeight) Height = Math.Max(MinHeight, maxHeight);
    }
}
