using System.Collections.ObjectModel;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace FileOrganizer.Desktop.ViewModels;

/// <summary>A folder used before, as a button.</summary>
public sealed record RecentFolder(string Path, IRelayCommand Command, IRelayCommand Remove);

/// <summary>A picture of a video that was read during the scan; clicking it plays the video.</summary>
public sealed record ScanThumb(Bitmap Image, string VideoPath, IRelayCommand Play)
{
    /// <summary>The file name, shown when the pointer rests on the picture.</summary>
    public string Name => Path.GetFileName(VideoPath);
}

public partial class ScanViewModel : ViewModelBase
{
    private readonly AppController app;
    private readonly MainViewModel main;
    private int applyQueued;
    private string? lastThumb;

    public ObservableCollection<RecentFolder> Recent { get; } = new();
    public ObservableCollection<ScanThumb> Thumbs { get; } = new();

    /// <summary>Set by the view: lets the user pick a folder (null: cancelled).</summary>
    public Func<Task<string?>>? PickFolder { get; set; }

    [ObservableProperty] private string folder = "";
    [ObservableProperty] private bool isRunning;
    [ObservableProperty] private bool stopping;
    [ObservableProperty] private bool isIdle = true;
    [ObservableProperty] private bool isDone;
    [ObservableProperty] private bool isProblem;
    [ObservableProperty] private string message = "";
    [ObservableProperty] private string error = "";

    [ObservableProperty] private double filePercent;
    [ObservableProperty] private string filePercentText = "";
    [ObservableProperty] private string fileText = "";
    [ObservableProperty] private string left = "";
    [ObservableProperty] private string currentFile = "";
    [ObservableProperty] private string stageText = "";
    [ObservableProperty] private bool showFrames;
    [ObservableProperty] private double framePercent;
    [ObservableProperty] private string frameText = "";
    [ObservableProperty] private string lines = "";
    [ObservableProperty] private string matched = "0";
    [ObservableProperty] private string unmatched = "0";
    [ObservableProperty] private string corrupt = "0";
    [ObservableProperty] private string cacheHits = "0";
    [ObservableProperty] private string ocrDone = "0";
    [ObservableProperty] private string elapsed = "0:00";

    // Beside the progress
    public ObservableCollection<string> Categories { get; } = new();
    [ObservableProperty] private string rulesText = "";
    [ObservableProperty] private string lastScanText = "";
    [ObservableProperty] private string lastScanDetail = "";

    public bool HasError => Error.Length > 0;
    public bool HasThumbs => Thumbs.Count > 0;
    partial void OnErrorChanged(string value) => OnPropertyChanged(nameof(HasError));

    public ScanViewModel(AppController app, MainViewModel main)
    {
        this.app = app;
        this.main = main;
        Folder = app.TargetFolder;
        Thumbs.CollectionChanged += (_, _) => OnPropertyChanged(nameof(HasThumbs));
        app.Engine.ProgressChanged += _ =>
        {
            if (Interlocked.Exchange(ref applyQueued, 1) == 1) return;
            Dispatcher.UIThread.Post(() => { Volatile.Write(ref applyQueued, 0); Apply(app.Engine.Current); });
        };
        app.TargetChanged += () => Dispatcher.UIThread.Post(() => { Folder = app.TargetFolder; RefreshInfo(); });
        app.SettingsChanged += () => Dispatcher.UIThread.Post(RefreshInfo);
        app.ResultChanged += () => Dispatcher.UIThread.Post(RefreshInfo);
        RefreshInfo();
        if (app.Result != null) ShowLastResult();
    }

    /// <summary>The recent folders, the rules in short, and the last result.</summary>
    public void RefreshInfo()
    {
        Recent.Clear();
        foreach (var f in app.RecentFolders) Recent.Add(new RecentFolder(f, new RelayCommand(() => UseRecent(f)), new RelayCommand(() => { app.ForgetFolder(f); RefreshInfo(); })));
        Categories.Clear();
        foreach (var c in app.Settings.Categories) Categories.Add(c.Name);
        int rules = app.Settings.Categories.Sum(c => c.Rules.Count);
        RulesText = app.Settings.Categories.Count == 0
            ? "No categories yet. Add them under Rules."
            : $"{Fmt.Plural(app.Settings.Categories.Count, "category", "categories")} with {Fmt.Plural(rules, "rule", "rules")}.";
        var r = app.Result;
        if (r == null) { LastScanText = "Nothing has been scanned yet."; LastScanDetail = "A scan only reads the files: nothing is moved until you say so."; }
        else
        {
            string name = Path.GetFileName(r.Folder.TrimEnd('/', '\\'));
            LastScanText = $"{Fmt.Ago(r.FinishedUtc)}: {Fmt.Plural(r.Decisions.Count, "file", "files")} in {(name.Length > 0 ? name : r.Folder)}.";
            LastScanDetail = $"{r.Matched.Count():N0} matched, {r.Unmatched.Count():N0} unmatched.";
        }
    }

    private void ShowLastResult()
    {
        var r = app.Result!;
        IsIdle = false; IsDone = true;
        Message = $"Looked at {r.Decisions.Count:N0} files: {r.Matched.Count():N0} matched, {r.Unmatched.Count():N0} unmatched.";
    }

    private void Apply(OrganizeProgress p)
    {
        IsRunning = app.Engine.IsRunning && p.Phase is OrganizePhase.Listing or OrganizePhase.Scanning;
        IsIdle = p.Phase == OrganizePhase.Idle;
        IsDone = p.Phase == OrganizePhase.Done;
        IsProblem = p.Phase is OrganizePhase.Cancelled or OrganizePhase.Failed;
        Message = p.Phase == OrganizePhase.Failed ? "The scan stopped with an error: " + p.Message : p.Message;
        if (p.Phase is OrganizePhase.Done or OrganizePhase.Cancelled or OrganizePhase.Failed) Stopping = false;

        FilePercent = p.Fraction * 100;
        FilePercentText = p.Phase == OrganizePhase.Idle ? "" : (int)Math.Floor(p.Fraction * 100) + "%";
        FileText = p.FilesTotal == 0 ? (p.Phase == OrganizePhase.Listing ? "Listing the files..." : "") : $"{p.FilesDone:N0} of {p.FilesTotal:N0} files";
        Left = IsRunning ? (p.Remaining is TimeSpan r ? Fmt.Approx(r) + " left" : "") : "";
        CurrentFile = p.CurrentFile;
        StageText = p.Stage switch { "Filename" => "Checking the file name", "Tags" => "Reading the title and comment tags", "Reading video" => "Reading text in the video", _ => "" };
        ShowFrames = p.Stage == "Reading video" && p.OcrFrames > 0;
        FramePercent = p.OcrFrames > 0 ? p.OcrFrame * 100.0 / p.OcrFrames : 0;
        FrameText = p.OcrFrames > 0 ? $"frame {p.OcrFrame} of {p.OcrFrames}" : "";
        Lines = string.Join("\n", p.RecentLines);
        Matched = p.Matched.ToString("N0"); Unmatched = p.Unmatched.ToString("N0"); Corrupt = p.Corrupt.ToString("N0");
        CacheHits = p.CacheHits.ToString("N0"); OcrDone = p.OcrDone.ToString("N0");
        Elapsed = Fmt.Time(p.Elapsed);

        if (p.LastThumbnail != null && p.LastThumbnail != lastThumb)
        {
            lastThumb = p.LastThumbnail;
            var bitmap = Bitmaps.Load(p.LastThumbnail);
            string? video = p.LastThumbnailVideo;
            if (bitmap != null && video != null)
            {
                Thumbs.Insert(0, new ScanThumb(bitmap, video, new RelayCommand(() => Play(video))));
                while (Thumbs.Count > 8) Thumbs.RemoveAt(Thumbs.Count - 1);
            }
        }
        if (p.Phase is OrganizePhase.Done or OrganizePhase.Cancelled or OrganizePhase.Failed) RefreshInfo();
    }

    [RelayCommand]
    private async Task Browse()
    {
        if (PickFolder == null) return;
        var chosen = await PickFolder();
        if (chosen != null) { Folder = chosen; Accept(); }
    }

    private void UseRecent(string path) { Folder = path; Accept(); }

    /// <summary>Takes the typed folder. Returns false (and says why) when it is not usable.</summary>
    private bool Accept()
    {
        var problem = app.SetTargetFolder(Folder);
        Error = problem ?? "";
        return problem == null;
    }

    [RelayCommand]
    private void Start()
    {
        if (!Accept()) return;
        var problem = app.StartScan();
        if (problem != null) { Error = problem; return; }
        Thumbs.Clear();
        lastThumb = null;
        IsRunning = true; IsIdle = false; IsDone = false; IsProblem = false;
    }

    [RelayCommand] private void Stop() { Stopping = true; app.StopScan(); }
    [RelayCommand] private void Review() => main.ShowReview();
    [RelayCommand] private void ShowUnmatched() => main.ShowUnmatched();
    [RelayCommand] private void EditRules() => main.ShowRules();
}
