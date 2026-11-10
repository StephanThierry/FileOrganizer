using System.Diagnostics;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace FileOrganizer.Desktop.ViewModels;

/// <summary>A toast and a confirmation dialog, like the other screens of the program family have.</summary>
public abstract partial class ViewModelBase : ObservableObject
{
    private CancellationTokenSource? toastTimer;
    private Action? confirmed;

    [ObservableProperty] private string? toast;
    [ObservableProperty] private bool confirmVisible;
    [ObservableProperty] private string confirmTitle = "";
    [ObservableProperty] private string confirmText = "";
    [ObservableProperty] private string confirmButton = "OK";

    // "Delete confirmation": ask before deleting a file (on by default); kept in settings.json, the same on Review and Unmatched
    private AppController? confirmApp;
    [ObservableProperty] private bool confirmDelete = true;
    partial void OnConfirmDeleteChanged(bool value) => confirmApp?.SetConfirmDelete(value);

    protected void BindConfirmDelete(AppController app)
    {
        confirmApp = app;
        ConfirmDelete = app.Settings.Opt.ConfirmDelete;
        app.SettingsChanged += () => Dispatcher.UIThread.Post(() => ConfirmDelete = app.Settings.Opt.ConfirmDelete);
    }

    protected void ShowToast(string text, int ms = 5000)
    {
        Toast = text;
        toastTimer?.Cancel();
        toastTimer = new CancellationTokenSource();
        var token = toastTimer.Token;
        Task.Delay(ms, token).ContinueWith(t => { if (!t.IsCanceled) Dispatcher.UIThread.Post(() => Toast = null); });
    }

    protected void Ask(string title, string text, string button, Action yes)
    {
        ConfirmTitle = title; ConfirmText = text; ConfirmButton = button;
        confirmed = yes;
        ConfirmVisible = true;
    }

    [RelayCommand] private void ConfirmYes() { var yes = confirmed; ConfirmVisible = false; confirmed = null; yes?.Invoke(); }
    [RelayCommand] private void ConfirmNo() { ConfirmVisible = false; confirmed = null; }

    /// <summary>Asks, then deletes the file of a row: to the Recycle Bin where there is one, else for good (and says which).</summary>
    protected void AskDeleteFile(AppController app, FileDecision d)
    {
        string? path = app.PathOf(d);
        if (path == null || !File.Exists(path)) { ShowToast("The file is not there: " + path); return; }
        // Delete confirmation off: the file is deleted at once
        if (!app.Settings.Opt.ConfirmDelete)
        {
            var (done, text) = app.DeleteFile(d);
            ShowToast(text, done ? 6000 : 10000);
            return;
        }
        bool recoverable = app.CanRecover(d);
        Ask(recoverable ? $"Move this file to the {RecycleBin.Name}?" : "Delete this file permanently?",
            $"{d.FileName}\n{Path.GetDirectoryName(path)}\n\n"
            + (recoverable ? $"You can get it back from the {RecycleBin.Name}." : $"This location has no {RecycleBin.Name}: the file is gone for good."),
            recoverable ? $"Move to {RecycleBin.Name}" : "Delete permanently",
            () => { var (ok, message) = app.DeleteFile(d); ShowToast(message, ok ? 6000 : 10000); });
    }

    /// <summary>Set by the view: opens a folder in the file manager.</summary>
    public Func<string, Task>? OpenFolder { get; set; }

    /// <summary>Opens a file in the program the system uses for it.</summary>
    protected void Play(string? path)
    {
        try
        {
            if (path == null || !File.Exists(path)) { ShowToast("The file is not there: " + path); return; }
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception e) { ShowToast("Could not open the file: " + e.Message); }
    }
}

/// <summary>Small pictures from the thumbnail folder, with a cache.</summary>
public static class Bitmaps
{
    private const int MaxCached = 600;
    private static readonly Dictionary<string, Bitmap?> Cache = new();
    private static readonly Queue<string> Order = new();

    public static Bitmap? Load(string? path)
    {
        if (path == null) return null;
        lock (Cache) { if (Cache.TryGetValue(path, out var hit)) return hit; }
        Bitmap? bitmap = null;
        try
        {
            if (File.Exists(path))
            {
                using var stream = File.OpenRead(path);
                bitmap = Bitmap.DecodeToWidth(stream, 240);
            }
        }
        catch (Exception) { /* unreadable picture: show none */ }
        lock (Cache)
        {
            Cache[path] = bitmap;
            Order.Enqueue(path);
            while (Order.Count > MaxCached) Cache.Remove(Order.Dequeue());
        }
        return bitmap;
    }
}

/// <summary>A video's picture is made when its row is first shown (in the background: making it opens the video).</summary>
public sealed class ThumbLoader
{
    private static readonly SemaphoreSlim Gate = new(2);
    private readonly Func<string?> make;
    private readonly Action changed;
    private bool started;
    private Bitmap? bitmap;

    public ThumbLoader(Func<string?> make, Action changed) { this.make = make; this.changed = changed; }

    public Bitmap? Bitmap
    {
        get
        {
            if (!started) { started = true; _ = LoadAsync(); }
            return bitmap;
        }
    }

    private async Task LoadAsync()
    {
        await Gate.WaitAsync();
        try
        {
            var loaded = await Task.Run(() => Bitmaps.Load(make()));
            if (loaded != null) { bitmap = loaded; Dispatcher.UIThread.Post(changed); }
        }
        finally { Gate.Release(); }
    }
}
