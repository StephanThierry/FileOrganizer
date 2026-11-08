using System.Diagnostics;
using Avalonia;
using Avalonia.Headless;
using Avalonia.Threading;
using FileOrganizer;
using FileOrganizer.Desktop;
using FileOrganizer.Desktop.ViewModels;
using FileOrganizer.Desktop.Views;

// Usage: DesktopShots <work folder> <output folder>
// With SHOT_FFMPEG (a folder with ffmpeg) and SHOT_MODELS (a folder with the two OCR .onnx files) it makes real test videos with
// burned-in text, scans them with the real OCR and walks through the screens, saving a screenshot of each.
string work = args[0], output = args[1];
Directory.CreateDirectory(output);
string data = Path.Combine(work, "data"), files = Path.Combine(work, "files");
var size = (Environment.GetEnvironmentVariable("SHOT_SIZE") ?? "1320x860").Split('x');

Directory.CreateDirectory(data); Directory.CreateDirectory(files);
string? ffmpegFolder = Environment.GetEnvironmentVariable("SHOT_FFMPEG");
string? modelsSource = Environment.GetEnvironmentVariable("SHOT_MODELS");
if (modelsSource != null)
{
    Directory.CreateDirectory(Path.Combine(data, "models"));
    foreach (var f in Directory.GetFiles(modelsSource, "*.onnx")) File.Copy(f, Path.Combine(data, "models", Path.GetFileName(f)), true);
}

// --- Test files ---
void Video(string name, string text, int seconds = 5)
{
    string path = Path.Combine(files, name);
    if (File.Exists(path) || ffmpegFolder == null) return;
    var filter = text.Length == 0 ? "null"
        : $"drawtext=fontfile='C\\:/Windows/Fonts/arial.ttf':text='{text}':fontsize=44:fontcolor=white:x=40:y=150";
    var info = new ProcessStartInfo(Path.Combine(ffmpegFolder, "ffmpeg.exe")) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardError = true };
    foreach (var a in new[] { "-hide_banner", "-y", "-f", "lavfi", "-i", $"color=c=0x203050:s=640x360:d={seconds}:r=25", "-vf", filter, "-c:v", "libx264", "-pix_fmt", "yuv420p", path })
        info.ArgumentList.Add(a);
    using var p = Process.Start(info)!;
    p.StandardError.ReadToEnd();
    p.WaitForExit();
}
Video("Python Tutorial 3.mp4", "");
Video("holiday_001.mp4", "Online Course");
Video("clip_0417.mp4", "BarryTV Original");
Video("mystery_a.mp4", "examplesite.com exclusive");
Video("mystery_b.mp4", "examplesite.com premium");
Video("plain.mp4", "");
foreach (var (n, t) in new[] { ("notes.txt", "x"), ("Show S01E02.txt", "x"), ("readme.md", "x") })
    File.WriteAllText(Path.Combine(files, n), t);

if (!SettingsStore.Exists(data))
{
    var s = new Settings
    {
        Categories =
        {
            new Category { Name = "Tutorials", Rules = { new Rule("Contains", "Course"), new Rule("Contains", "Lesson") } },
            new Category { Name = "Barry", IgnoreCategoryInContent = true, Rules = { new Rule("Regex", @"^Barry.*\.mp4$"), new Rule("Contains", "barrytv") } },
            new Category { Name = "Series", IgnoreCategoryInContent = true, Rules = { new Rule("Regex", @"S\d\dE\d\d") } },
        },
    };
    SettingsStore.Save(data, s);
}

FileOrganizer.Desktop.Program.App = AppController.Open(data);
var app = FileOrganizer.Desktop.Program.App;
app.SetTargetFolder(files);

AppBuilder.Configure<FileOrganizer.Desktop.App>()
    .UseSkia().UseHeadless(new AvaloniaHeadlessPlatformOptions { UseHeadlessDrawing = false }).WithInterFont().SetupWithoutStarting();

var window = new MainWindow { Width = int.Parse(size[0]), Height = int.Parse(size[1]) };
var vm = (MainViewModel)window.DataContext!;
window.Show();

void Pump(int ms = 300)
{
    var end = DateTime.UtcNow.AddMilliseconds(ms);
    while (DateTime.UtcNow < end) { Dispatcher.UIThread.RunJobs(); Thread.Sleep(20); }
}
void Shot(string name)
{
    Pump();
    window.CaptureRenderedFrame()?.Save(Path.Combine(output, name + ".png"));
    Console.WriteLine("saved " + name);
}
void WaitFor(Func<bool> condition, int seconds, string what)
{
    var until = DateTime.UtcNow.AddSeconds(seconds);
    while (!condition()) { Pump(100); if (DateTime.UtcNow > until) throw new TimeoutException("Timed out waiting for " + what); }
}

// --- The screens ---
vm.SelectedTab = MainViewModel.RulesTab; vm.Rules.TryText = "Season S02E05 of the show"; Pump(500); Shot("o_rules");
vm.SelectedTab = MainViewModel.SettingsTab; Pump(300); Shot("o_settings");
vm.SelectedTab = MainViewModel.ScanTab; Pump(300); Shot("o_scan_idle");

if (ffmpegFolder != null && modelsSource != null)
{
    vm.Scan.StartCommand.Execute(null);
    WaitFor(() => app.Engine.Current.OcrFrame >= 4, 120, "the first video being read");
    Pump(300); Shot("o_scan_running");
    WaitFor(() => !app.Engine.IsRunning && app.Result != null, 300, "the scan to finish");
    Pump(800); Shot("o_scan_done");
    Console.WriteLine($"result: {app.Result!.Matched.Count()} matched, {app.Result.Unmatched.Count()} unmatched");
    foreach (var d in app.Result.Decisions) Console.WriteLine($"  {d.FileName}: {d.Category ?? "-"} ({d.Source}) {d.Trigger} | text: {string.Join(" / ", d.OcrLines)}");

    vm.SelectedTab = MainViewModel.ReviewTab; Pump(1500); Shot("o_review");
    var first = vm.Review.Rows[0];
    first.Choice = "Series"; Pump(300);
    vm.Review.CreateScriptsCommand.Execute(null); Pump(300); Shot("o_review_scripts");

    vm.SelectedTab = MainViewModel.UnmatchedTab; Pump(1500); Shot("o_unmatched");
    var mystery = vm.Unmatched.Rows.First(r => r.FileName.StartsWith("mystery"));
    Console.WriteLine("text chips: " + string.Join(" | ", mystery.Lines.Select(l => l.Text)));
    mystery.Lines[0].Command.Execute(null);
    vm.Unmatched.RuleText = "examplesite";
    vm.Unmatched.NewCategory = "Example site"; Pump(400); Shot("o_rule_dialog");
    Console.WriteLine("preview: " + vm.Unmatched.RulePreview);
    vm.Unmatched.AddRuleCommand.Execute(null); Pump(1200); Shot("o_unmatched_after");
    Console.WriteLine($"after the rule: {app.Result.Unmatched.Count()} unmatched; toast: {vm.Unmatched.Toast}");

    vm.SelectedTab = MainViewModel.ReviewTab; Pump(1500); Shot("o_review_after");
    vm.Review.AskMoveCommand.Execute(null); Pump(300); Shot("o_review_confirm");
    vm.Review.ConfirmYesCommand.Execute(null); Pump(1500); Shot("o_review_moved");
    Console.WriteLine("folders now: " + string.Join(", ", Directory.GetDirectories(files).Select(Path.GetFileName)));
    Console.WriteLine("toast: " + vm.Review.Toast);

    vm.SelectedTab = MainViewModel.RulesTab; Pump(800); Shot("o_rules_after");
}
app.Dispose();
