using Avalonia;

namespace FileOrganizer.Desktop;

public static class Program
{
    public static AppController App { get; set; } = null!;

    // Usage: FileOrganizer.Desktop [--data <folder>]
    //   --data  folder with settings.json (or FILEORG_DATA). Default: the current folder if it holds settings.json
    //           (like the command-line version), otherwise the user's application data folder.
    [STAThread]
    public static void Main(string[] args)
    {
        App = AppController.Open(DataLocation.ResolveDataFolder(args, null) ?? DataLocation.DefaultDataFolder());
        BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<Desktop.App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
}
