using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace FileOrganizer.Desktop.ViewModels;

public partial class ExtensionItem : ObservableObject
{
    private readonly Action edited;
    public string Extension { get; }
    [ObservableProperty] private bool isChecked;
    public ExtensionItem(string extension, bool isChecked, Action edited) { Extension = extension; this.isChecked = isChecked; this.edited = edited; }
    partial void OnIsCheckedChanged(bool value) => edited();
}

/// <summary>Which files get the deep scan, and how, plus where the program keeps its files.</summary>
public partial class SettingsViewModel : ViewModelBase
{
    private readonly AppController app;
    private bool loading;

    public ObservableCollection<ExtensionItem> Extensions { get; } = new();

    [ObservableProperty] private bool stopAtFirstMatch = true;
    [ObservableProperty] private string newExtension = "";
    [ObservableProperty] private bool isEdited;
    [ObservableProperty] private string modelsText = "";
    [ObservableProperty] private bool modelsReady;

    public string DataFolder => app.DataFolder;
    public string ModelsFolder => app.ModelsFolder;
    public string ScriptKind => OperatingSystem.IsWindows() ? "PowerShell scripts (.ps1)" : "shell scripts (.sh)";

    public SettingsViewModel(AppController app)
    {
        this.app = app;
        Load();
    }

    partial void OnStopAtFirstMatchChanged(bool value) => Edited();

    private void Edited() { if (!loading) IsEdited = true; }

    private void Load()
    {
        loading = true;
        var chosen = new HashSet<string>(app.Settings.Opt.VideoExtensions ?? Options.DefaultVideoExtensions, StringComparer.OrdinalIgnoreCase);
        Extensions.Clear();
        foreach (var e in Options.DefaultVideoExtensions.Concat(chosen.Except(Options.DefaultVideoExtensions, StringComparer.OrdinalIgnoreCase)))
            Extensions.Add(new ExtensionItem(e, chosen.Contains(e), Edited));
        StopAtFirstMatch = app.Settings.Opt.StopOcrAtFirstMatch;
        loading = false;
        IsEdited = false;
        string models = app.ModelsFolder;
        ModelsReady = Directory.Exists(models) && Directory.GetFiles(models, "*.onnx").Length >= 2;
        ModelsText = ModelsReady ? "The text recognition models are downloaded." : "The text recognition models (about 13 MB) are downloaded the first time a video is read.";
    }

    [RelayCommand]
    private void AddExtension()
    {
        string e = (NewExtension ?? "").Trim().ToLowerInvariant();
        if (e.Length == 0) return;
        if (!e.StartsWith('.')) e = "." + e;
        if (e.Length < 2 || e.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) { ShowToast($"\"{e}\" is not a file type."); return; }
        if (Extensions.Any(x => x.Extension.Equals(e, StringComparison.OrdinalIgnoreCase))) { ShowToast($"{e} is in the list already."); return; }
        Extensions.Add(new ExtensionItem(e, true, Edited));
        NewExtension = "";
        Edited();
    }

    [RelayCommand]
    private void Save()
    {
        var s = app.Settings.Clone();   // the rules stay as they are
        var chosen = Extensions.Where(x => x.IsChecked).Select(x => x.Extension).ToArray();
        bool isDefault = chosen.Length == Options.DefaultVideoExtensions.Length && chosen.All(c => Options.DefaultVideoExtensions.Contains(c, StringComparer.OrdinalIgnoreCase));
        s.Opt.VideoExtensions = isDefault ? null : chosen;
        s.Opt.StopOcrAtFirstMatch = StopAtFirstMatch;
        var problems = app.SaveSettings(s);
        if (problems.Count > 0) { ShowToast(problems[0], 9000); return; }
        IsEdited = false;
        ShowToast("Saved.");
    }

    [RelayCommand] private void Revert() => Load();
    [RelayCommand] private async Task OpenModels() { if (OpenFolder != null) { Directory.CreateDirectory(ModelsFolder); await OpenFolder(ModelsFolder); } }
    [RelayCommand] private async Task OpenData() { if (OpenFolder != null) await OpenFolder(DataFolder); }
}
