using Avalonia.Controls;
using FileOrganizer.Desktop.ViewModels;

namespace FileOrganizer.Desktop.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        if (!Design.IsDesignMode && Program.App != null) DataContext = new MainViewModel(Program.App);
    }
}
