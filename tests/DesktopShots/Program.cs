using System.Diagnostics;
using Avalonia;
using Avalonia.Headless;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Avalonia.Controls;
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
vm.SelectedTab = MainViewModel.RulesTab; vm.Rules.FilterText = "Season S02E05 of the show"; Pump(500); Shot("o_rules");
// The regex guide: opens beside a Regex rule, with a test text
vm.Rules.Selected = vm.Rules.Categories.First(c => c.Name == "Series");
vm.Rules.Selected.Rules[0].Activate(); vm.Rules.Guide.Text = "Show S01E07 and s12e10 but not S1E2"; Pump(500); Shot("o_rules_guide");
vm.Rules.Guide.Close();
// A rule with several terms: each Contains or Regex
var allRule = vm.Rules.Selected!.Rules[0];
allRule.Type = RuleItem.AllLabel;
allRule.Terms[0].Type = "Regex"; allRule.Terms[0].Value = "R.?C";
allRule.Terms[1].Value = "hub";
allRule.AddTermCommand.Execute(null); allRule.Terms[2].Value = "community";
allRule.Terms[0].Activate(); vm.Rules.Guide.Text = "REC and R3C and RC but not RXXC"; vm.Rules.FilterText = "REC hub community"; Pump(500); Shot("o_rules_allof");
vm.Rules.Guide.Close();

// Rules > Add: the cursor goes to the name box, so typing names the category (it used to stay "New category" and be saved so)
vm.Rules.RevertCommand.Execute(null); Pump(300);
vm.Rules.AddCategoryCommand.Execute(null); Pump(700);
var newCat = vm.Rules.Selected!;
Console.WriteLine($"[check] after Add the new category has no name yet: {newCat.Name.Length == 0}, list shows '{newCat.DisplayName}'");
foreach (var ch in "Zebra") { window.KeyTextInput(ch.ToString()); Pump(200); }
Console.WriteLine($"[check] typed Zebra right after Add: name = '{newCat.Name}' ({(newCat.Name == "Zebra" ? "OK" : "FAILED")}), still selected = {ReferenceEquals(vm.Rules.Selected, newCat)}");
vm.Rules.RevertCommand.Execute(null); Pump(300);

// Type-ahead in the category list: letters typed with the list selected go to the first category that starts with them
var catList = window.GetVisualDescendants().OfType<ListBox>().First(l => l.Name == "CategoryList");
vm.Rules.Selected = vm.Rules.SortedCategories.First(); Pump(300);
// Like a click on an item: the list item has the keyboard focus
var firstItemBox = catList.ContainerFromItem(vm.Rules.Selected!);
Console.WriteLine($"[check] before typing, the focus is in: {(firstItemBox?.Focus() == true ? "the selected list item" : "NOT the list")}");
Pump(200);
window.KeyTextInput("t"); Pump(250);
Console.WriteLine($"[check] list focused, typed t: selected = '{vm.Rules.Selected?.Name}' ({(vm.Rules.Selected?.Name == "Tutorials" ? "OK" : "FAILED")}), the list shows it selected = {ReferenceEquals(catList.SelectedItem, vm.Rules.Selected)}");
window.KeyTextInput("s"); Pump(250);
Console.WriteLine($"[check] then typed s (within a second: \"ts\" has no match, so a new search): selected = '{vm.Rules.Selected?.Name}' ({(vm.Rules.Selected?.Name == "Series" ? "OK" : "FAILED")})");
Pump(1200);
window.KeyTextInput("b"); Pump(250);
Console.WriteLine($"[check] typed b a second later: selected = '{vm.Rules.Selected?.Name}' ({(vm.Rules.Selected?.Name == "Barry" ? "OK" : "FAILED")})");

// Filter category: the name first, then the strings of the rules
vm.Rules.FilterText = "tut"; Pump(200);
Console.WriteLine($"[check] filter 'tut': {string.Join(", ", vm.Rules.SortedCategories.Select(c => c.Name + (c.HasMatchText ? " [" + c.MatchText + "]" : "")))} | {vm.Rules.FilterSummary}");
vm.Rules.FilterText = "lesson"; Pump(200);
Console.WriteLine($"[check] filter 'lesson': {string.Join(", ", vm.Rules.SortedCategories.Select(c => c.Name + (c.HasMatchText ? " [" + c.MatchText + "]" : "")))} | {vm.Rules.FilterSummary}");
vm.Rules.FilterText = "Season S02E05 of the show"; Pump(200);
Console.WriteLine($"[check] filter with a file name: {string.Join(", ", vm.Rules.SortedCategories.Select(c => c.Name + (c.HasMatchText ? " [" + c.MatchText + "]" : "")))}");
Shot("o_rules_filter");
// The edited category stays selected when the filter hides it
var keepSel = vm.Rules.Selected; vm.Rules.FilterText = "zzzqqq"; Pump(200);
Console.WriteLine($"[check] filter with no hit: list has {vm.Rules.SortedCategories.Count} categories, editor still shows '{vm.Rules.Selected?.Name}' ({(ReferenceEquals(vm.Rules.Selected, keepSel) && vm.Rules.HasSelected ? "OK" : "FAILED")})");
vm.Rules.FilterText = ""; Pump(200);
vm.SelectedTab = MainViewModel.SettingsTab; Pump(300); Shot("o_settings");
vm.SelectedTab = MainViewModel.ScanTab; Pump(300); Shot("o_scan_idle");

// The Unmatched screen with made-up findings (no videos needed): the "make a rule" dialog with the regex guide to its right
{
    var fake = new OrganizeResult { Folder = files, Decisions = { new FileDecision { FileName = "summer clip.mp4", IsVideo = true, OcrLines = { "REC", "hub", "my own summer", "Ppp.tV", "p-p-p.tv", "P-p-p.tv", "P-pp.t", "P-P-p.tv", "Pp-p.t", "Pp-p.tv", "P-P-P.TV", "P-P-PTV", "P-P-P. TV" } } } };
    typeof(AppController).GetProperty("Result")!.SetValue(app, fake);
    vm.SelectedTab = MainViewModel.UnmatchedTab; vm.Unmatched.Reload(); Pump(600);
    var row = vm.Unmatched.Rows[0];
    // Clicking the file name makes a rule from it, like any other text
    row.NameCommand.Execute(null); Pump(300);
    Console.WriteLine($"[check] clicked the file name: dialog open = {vm.Unmatched.RuleVisible}, text = '{vm.Unmatched.RuleText}' ({(vm.Unmatched.RuleText == "summer clip" ? "OK" : "FAILED")}); the regex guide's test box starts with '{vm.Unmatched.Guide.Text}' ({(vm.Unmatched.Guide.Text == "summer clip.mp4" ? "OK" : "FAILED")})");
    vm.Unmatched.CancelRuleCommand.Execute(null); Pump(200);
    row.Lines[0].IsSelected = true; row.Lines[1].IsSelected = true; Pump(300); Shot("o_unmatched_ticked");
    row.MakeAllRuleCommand.Execute(null);
    vm.Unmatched.RuleTerms[0].Type = "Regex"; vm.Unmatched.RuleTerms[0].Value = "R.?C"; vm.Unmatched.Guide.Text = "REC and R3C and RC but not RXXC"; Pump(500); Shot("o_dialog_allof_guide");
    vm.Unmatched.CancelRuleCommand.Execute(null);
    row.Lines[2].Command.Execute(null); vm.Unmatched.RuleType = "Regex"; vm.Unmatched.RuleText = "my.?own"; vm.Unmatched.Guide.Text = "my own summer"; Pump(500); Shot("o_dialog_regex_guide");
    vm.Unmatched.CancelRuleCommand.Execute(null);
    // The text read in many ways: one tag, all the readings in the dialog with what a rule would catch
    var logo = row.Lines.First(l => l.HasVariants && l.Text.Contains("p", StringComparison.OrdinalIgnoreCase) && l.Variants.Count > 3);
    Console.WriteLine("logo tag: " + logo.Text + " | " + logo.VariantsBadge);
    Shot("o_unmatched_variants");
    logo.Command.Execute(null);
    vm.Unmatched.RuleType = "Regex"; vm.Unmatched.RuleText = "p.?p.?p.?tv"; Pump(500); Shot("o_dialog_variants_regex");
    Console.WriteLine("summary: " + vm.Unmatched.VariantSummary);
    vm.Unmatched.RuleType = "Contains"; vm.Unmatched.RuleText = "p-p-p.tv"; Pump(500); Shot("o_dialog_variants_contains");
    Console.WriteLine("summary: " + vm.Unmatched.VariantSummary);
    vm.Unmatched.CancelRuleCommand.Execute(null);
    typeof(AppController).GetProperty("Result")!.SetValue(app, null);
    vm.Unmatched.Reload();
}
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
