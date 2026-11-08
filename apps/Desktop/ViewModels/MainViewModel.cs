using CommunityToolkit.Mvvm.ComponentModel;

namespace FileOrganizer.Desktop.ViewModels;

public partial class MainViewModel : ObservableObject
{
    public AppController App { get; }
    public ScanViewModel Scan { get; }
    public ReviewViewModel Review { get; }
    public UnmatchedViewModel Unmatched { get; }
    public RulesViewModel Rules { get; }
    public SettingsViewModel Settings { get; }

    public const int ScanTab = 0, ReviewTab = 1, UnmatchedTab = 2, RulesTab = 3, SettingsTab = 4;

    /// <summary>0 = Scan, 1 = Review, 2 = Unmatched, 3 = Rules, 4 = Settings.</summary>
    [ObservableProperty] private int selectedTab;

    // The numbers on the sidebar items
    [ObservableProperty] private string reviewBadge = "";
    [ObservableProperty] private bool hasReviewBadge;
    [ObservableProperty] private string unmatchedBadge = "";
    [ObservableProperty] private bool hasUnmatchedBadge;

    public MainViewModel(AppController app)
    {
        App = app;
        Scan = new ScanViewModel(app, this);
        Review = new ReviewViewModel(app, this);
        Unmatched = new UnmatchedViewModel(app, this);
        Rules = new RulesViewModel(app, this);
        Settings = new SettingsViewModel(app);
        app.ResultChanged += () => Avalonia.Threading.Dispatcher.UIThread.Post(UpdateBadges);
        UpdateBadges();
        // Start where there is something to do
        SelectedTab = app.Result != null ? ReviewTab : app.IsFirstRun || app.Settings.Categories.Count == 0 ? RulesTab : ScanTab;
    }

    public void UpdateBadges()
    {
        int review = App.Result?.Decisions.Count(d => !d.Done && d.EffectiveCategory != null) ?? 0;
        int unmatched = App.Result?.Decisions.Count(d => d.Category == null && d.OverrideCategory == null) ?? 0;
        ReviewBadge = review > 999 ? "999+" : review.ToString(); HasReviewBadge = review > 0;
        UnmatchedBadge = unmatched > 999 ? "999+" : unmatched.ToString(); HasUnmatchedBadge = unmatched > 0;
    }

    public void ShowScan() => SelectedTab = ScanTab;
    public void ShowReview() => SelectedTab = ReviewTab;
    public void ShowUnmatched() => SelectedTab = UnmatchedTab;
    public void ShowRules() => SelectedTab = RulesTab;
    public void ShowSettings() => SelectedTab = SettingsTab;
}
