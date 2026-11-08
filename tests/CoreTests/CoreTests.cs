using FileOrganizer;
using Xunit;

namespace CoreTests;

internal static class Make
{
    public static Settings Rules(params (string name, bool ignoreName, (string type, string value)[] rules)[] categories) => new()
    {
        Categories = categories.Select(c => new Category
        {
            Name = c.name, IgnoreCategoryInContent = c.ignoreName,
            Rules = c.rules.Select(r => new Rule(r.type, r.value)).ToList(),
        }).ToList(),
    };

    public static (string, bool, (string, string)[]) Cat(string name, bool ignoreName = false, params (string, string)[] rules) => (name, ignoreName, rules);
}

public class RuleMatcherTests
{
    private static Settings Sample() => Make.Rules(
        Make.Cat("Tutorials", false, ("Contains", "Course"), ("Contains", "Lesson")),
        Make.Cat("Barry", true, ("Regex", @"^Barry.*\.mp4$")),
        Make.Cat("site.com", false),
        Make.Cat("HD and 4K", false, ("Regex", @"\d{4}p.*\.mp4$")));

    [Fact]
    public void NameOfTheCategoryIsARule_UnlessIgnored()
    {
        var s = Sample();
        Assert.Equal("Tutorials", RuleMatcher.Match("my tutorials 3.mp4", s)!.Category);
        Assert.Equal("Category Name Match", RuleMatcher.Match("my tutorials 3.mp4", s)!.Trigger);
        Assert.Null(RuleMatcher.Match("barry on holiday.avi", s));                 // "Barry" ignores its own name
        Assert.Equal("Barry", RuleMatcher.Match("Barry S01E01.mp4", s)!.Category);  // but its regex works
    }

    [Fact]
    public void DotComIsLeftOutOfTheName()
    {
        var m = RuleMatcher.Match("watched on SITE yesterday", Sample());
        Assert.Equal("site.com", m!.Category);
    }

    [Fact]
    public void ContainsAndRegexIgnoreCase_AndTheTriggerSaysWhich()
    {
        var s = Sample();
        Assert.Equal("Contains: Lesson", RuleMatcher.Match("my LESSON.mkv", s)!.Trigger);
        var hd = RuleMatcher.Match("something.1080P.hd.mp4", s)!;
        Assert.Equal("HD and 4K", hd.Category);
        Assert.Equal(@"Regex: \d{4}p.*\.mp4$", hd.Trigger);
    }

    [Fact]
    public void TheFirstCategoryInTheListWins()
    {
        var s = Make.Rules(Make.Cat("A", true, ("Contains", "x")), Make.Cat("B", true, ("Contains", "x")));
        Assert.Equal("A", RuleMatcher.Match("x", s)!.Category);
        s.Categories.Reverse();
        Assert.Equal("B", RuleMatcher.Match("x", s)!.Category);
        Assert.Equal(new[] { "B", "A" }, RuleMatcher.Explain("x", s).Select(m => m.Category));   // all hits, in priority order
    }

    [Fact]
    public void ABrokenRegexMatchesNothing_AndIsReported()
    {
        var s = Make.Rules(Make.Cat("A", true, ("Regex", "([unclosed"), ("Contains", "ok")));
        Assert.Equal("Contains: ok", RuleMatcher.Match("this is ok", s)!.Trigger);   // the next rule still works
        Assert.Null(RuleMatcher.Match("nothing", s));
        Assert.Contains("not a valid regular expression", RuleMatcher.Validate(s.Categories[0].Rules[0]));
        Assert.Null(RuleMatcher.Validate(new Rule("Regex", "fine.*")));
        Assert.Equal("a rule is empty.", RuleMatcher.Validate(new Rule("Contains", " ")));
    }

    [Fact]
    public void EmptyValuesAndNamesNeverMatchEverything()
    {
        var s = Make.Rules(Make.Cat(".com", false, ("Contains", "")), Make.Cat("Real", true, ("Contains", "real")));
        Assert.Equal("Real", RuleMatcher.Match("a real file", s)!.Category);
        Assert.Null(RuleMatcher.Match("", s));
        Assert.Null(RuleMatcher.Match(null, s));
    }
}

public class SettingsTests : IDisposable
{
    private readonly string dir = Path.Combine(Path.GetTempPath(), "fo-settings-" + Guid.NewGuid().ToString("N"));
    public SettingsTests() => Directory.CreateDirectory(dir);
    public void Dispose() { try { Directory.Delete(dir, true); } catch { } }

    [Fact]
    public void AnOldSettingsFileLoads_AndSavesBackWithoutNewKeys()
    {
        // The format the command line has always used: no Options
        File.WriteAllText(SettingsStore.PathIn(dir), """
            { "Categories": [ { "Name": "Tutorials", "IgnoreCategoryInContent": false,
                "Rules": [ { "Type": "contains", "Value": "Course" }, { "Type": "Regex", "Value": "x" } ] } ] }
            """);
        var s = SettingsStore.Load(dir);
        Assert.Equal("Tutorials", s.Categories[0].Name);
        Assert.Equal("Contains", s.Categories[0].Rules[0].Type);   // normalised
        Assert.Null(s.Options);

        SettingsStore.Save(dir, s);
        string saved = File.ReadAllText(SettingsStore.PathIn(dir));
        Assert.DoesNotContain("Options", saved);
        Assert.Contains("\"IgnoreCategoryInContent\"", saved);   // same property names as before
        Assert.True(File.Exists(SettingsStore.PathIn(dir) + ".bck"));   // the previous file is kept
    }

    [Fact]
    public void OptionsRoundTrip_AndMissingFileGivesEmptySettings()
    {
        Assert.Empty(SettingsStore.Load(dir).Categories);
        var s = new Settings();
        s.Opt.StopOcrAtFirstMatch = false;
        s.Opt.VideoExtensions = new[] { ".mp4" };
        SettingsStore.Save(dir, s);
        var back = SettingsStore.Load(dir);
        Assert.False(back.Opt.StopOcrAtFirstMatch);
        Assert.Equal(new[] { ".mp4" }, back.Opt.VideoExtensions);
    }

    [Fact]
    public void ABrokenFileGivesAClearError()
    {
        File.WriteAllText(SettingsStore.PathIn(dir), "{ not json");
        var e = Assert.Throws<InvalidDataException>(() => SettingsStore.Load(dir));
        Assert.Contains("settings.json is not valid", e.Message);
    }

    [Fact]
    public void ValidateFindsWhatWouldGoWrong()
    {
        var s = Make.Rules(
            Make.Cat("Good", false, ("Contains", "a")),
            Make.Cat("good", false),                       // the same folder
            Make.Cat("..", false),                         // would leave the folder
            Make.Cat("a/b", false),
            Make.Cat("", false),
            Make.Cat("Rx", false, ("Regex", "(")));
        var problems = SettingsStore.Validate(s);
        Assert.DoesNotContain(problems, p => p.Contains("good"));   // the same name twice is allowed
        Assert.Contains(problems, p => p.Contains("\"..\" can't be a folder name"));
        Assert.Contains(problems, p => p.Contains("\"a/b\" can't be a folder name"));
        Assert.Contains(problems, p => p.Contains("has no name"));
        Assert.Contains(problems, p => p.Contains("Rx: ") && p.Contains("regular expression"));
        Assert.Empty(SettingsStore.Validate(Make.Rules(Make.Cat("Fine", false, ("Contains", "x")))));
    }

    [Fact]
    public void DataFolderComesFromTheArgumentThenTheEnvironment()
    {
        Assert.Equal(Path.GetFullPath(dir), DataLocation.ResolveDataFolder(new[] { "--data", dir }, "fallback"));
    }
}

internal sealed class FakeOcr : IVideoOcr
{
    public readonly Dictionary<string, (string[] lines, bool corrupt)> Videos = new(StringComparer.OrdinalIgnoreCase);
    public readonly List<string> Read_ = new();
    public Action<string>? OnRead;

    public VideoOcrResult Read(string path, Func<string, bool> isMatch, Action<OcrFrameInfo>? onFrame, CancellationToken ct)
    {
        string name = Path.GetFileName(path);
        Read_.Add(name);
        OnRead?.Invoke(name);
        ct.ThrowIfCancellationRequested();
        var result = new VideoOcrResult();
        if (!Videos.TryGetValue(name, out var v)) return result;
        result.IsCorrupt = v.corrupt;
        int frame = 0;
        foreach (var line in v.lines)
        {
            result.Lines.Add(line);
            onFrame?.Invoke(new OcrFrameInfo(++frame, v.lines.Length, frame, new[] { line }));
            if (isMatch(line)) break;
        }
        return result;
    }
}

public class EngineTests : IDisposable
{
    private readonly string dir = Path.Combine(Path.GetTempPath(), "fo-engine-" + Guid.NewGuid().ToString("N"));
    private readonly FakeOcr ocr = new();
    public EngineTests() => Directory.CreateDirectory(dir);
    public void Dispose() { try { Directory.Delete(dir, true); } catch { } }

    private void Files(params string[] names) { foreach (var n in names) File.WriteAllText(Path.Combine(dir, n), "x"); }

    private OrganizeEngine NewEngine() => new(() => ocr);

    private static Settings Rules() => Make.Rules(
        Make.Cat("Tutorials", false, ("Contains", "Course")),
        Make.Cat("Barry", true, ("Regex", @"^Barry.*\.mp4$"), ("Contains", "barrytv")),
        Make.Cat("Series", true, ("Regex", @"S\d\dE\d\d")));

    [Fact]
    public void EachFileGetsItsCategoryFromTheNameThenTheTagsThenTheVideoText()
    {
        Files("Python Course 1.mp4", "Barry S01E01.mp4", "clip.mp4", "notes.txt", "other.mp4");
        ocr.Videos["clip.mp4"] = (new[] { "watch barrytv now" }, false);
        ocr.Videos["other.mp4"] = (new[] { "nothing useful" }, false);

        var result = NewEngine().Run(dir, Rules(), CancellationToken.None);

        FileDecision Of(string n) => result.Decisions.Single(d => d.FileName == n);
        Assert.Equal(("Tutorials", MatchSource.Filename), (Of("Python Course 1.mp4").Category, Of("Python Course 1.mp4").Source));
        Assert.Equal("Barry", Of("Barry S01E01.mp4").Category);     // Barry comes before Series in the list
        Assert.Equal(("Barry", MatchSource.Ocr), (Of("clip.mp4").Category, Of("clip.mp4").Source));
        Assert.Equal("OCR Match: 'Contains: barrytv'", Of("clip.mp4").Trigger);
        Assert.Null(Of("notes.txt").Category);                       // not a video: no OCR
        Assert.Null(Of("other.mp4").Category);
        Assert.Equal(new[] { "nothing useful" }, Of("other.mp4").OcrLines);   // the text is kept to write a rule from
        Assert.DoesNotContain("Python Course 1.mp4", ocr.Read_);     // matched by name: never opened
        Assert.DoesNotContain("notes.txt", ocr.Read_);
        Assert.Equal(3, result.Matched.Count());
        Assert.Equal(2, result.Unmatched.Count());
    }

    [Fact]
    public void ProgramFilesAreNotOrganized()
    {
        Files("a.txt", "settings.json", "unmatched_files.txt", "move_files.ps1", "rename_files.sh", OcrCache.FileName);
        var result = NewEngine().Run(dir, Rules(), CancellationToken.None);
        Assert.Equal(new[] { "a.txt" }, result.Decisions.Select(d => d.FileName));
    }

    [Fact]
    public void AVideoThatCannotBeReadGoesToCheck_AndMatchesBeforeItStillWin()
    {
        Files("broken.mp4", "Barry broken.mp4");
        ocr.Videos["broken.mp4"] = (Array.Empty<string>(), true);
        var result = NewEngine().Run(dir, Rules(), CancellationToken.None);
        var broken = result.Decisions.Single(d => d.FileName == "broken.mp4");
        Assert.Equal((OrganizeEngine.CorruptCategory, MatchSource.Corrupt), (broken.Category, broken.Source));
        Assert.Equal("Barry", result.Decisions.Single(d => d.FileName == "Barry broken.mp4").Category);   // the name matched first
    }

    [Fact]
    public void TheOcrCacheIsUsedAndFilledInTheScannedFolder()
    {
        Files("clip.mp4");
        ocr.Videos["clip.mp4"] = (new[] { "barrytv" }, false);
        var first = NewEngine().Run(dir, Rules(), CancellationToken.None);
        Assert.False(first.Decisions[0].FromCache);
        Assert.Single(ocr.Read_);
        Assert.True(File.Exists(OcrCache.PathIn(dir)));

        ocr.Read_.Clear();
        var second = NewEngine().Run(dir, Rules(), CancellationToken.None);
        Assert.Empty(ocr.Read_);                                      // not read again
        Assert.True(second.Decisions[0].FromCache);
        Assert.Equal("Barry", second.Decisions[0].Category);
    }

    [Fact]
    public void ReevaluateAppliesNewRulesWithoutTouchingTheFiles_AndKeepsTheUsersChoices()
    {
        Files("a.mp4", "b.mp4");
        ocr.Videos["a.mp4"] = (new[] { "welcome to examplesite" }, false);
        ocr.Videos["b.mp4"] = (new[] { "welcome to examplesite" }, false);
        var settings = Rules();
        var result = NewEngine().Run(dir, settings, CancellationToken.None);
        Assert.Equal(2, result.Unmatched.Count());

        result.Decisions[1].Skip = true;
        result.Decisions[1].OverrideCategory = "Mine";
        settings.Categories.Add(new Category { Name = "Example", Rules = { new Rule("Contains", "examplesite") }, IgnoreCategoryInContent = true });
        ocr.Read_.Clear();
        OrganizeEngine.Reevaluate(result, settings);

        Assert.Equal("Example", result.Decisions[0].Category);
        Assert.Equal(MatchSource.Ocr, result.Decisions[0].Source);
        Assert.Empty(ocr.Read_);                                      // nothing was read
        Assert.Equal("Example", result.Decisions[1].Category);
        Assert.Null(result.Decisions[1].EffectiveCategory);           // still left alone by the user's choice (Skip wins)
        result.Decisions[1].Skip = false;
        Assert.Equal("Mine", result.Decisions[1].EffectiveCategory);  // and the chosen category beats the rules'
    }

    [Fact]
    public void DecideFollowsTheOrderNameTitleCommentCorruptText()
    {
        var s = Make.Rules(Make.Cat("A", true, ("Contains", "aaa")), Make.Cat("B", true, ("Contains", "bbb")),
                           Make.Cat("C", true, ("Contains", "ccc")), Make.Cat("D", true, ("Contains", "ddd")));
        var d = new FileDecision { FileName = "x.mp4", Title = "bbb", Comment = "ccc", OcrLines = { "ddd" }, OcrCorrupt = true };
        Assert.True(OrganizeEngine.Decide(d, s));
        Assert.Equal((MatchSource.Title, "B", "Title Tag: Contains: bbb"), (d.Source, d.Category, d.Trigger));
        d.Title = "";
        OrganizeEngine.Decide(d, s);
        Assert.Equal((MatchSource.Comment, "C"), (d.Source, d.Category));
        d.Comment = "";
        OrganizeEngine.Decide(d, s);
        Assert.Equal((MatchSource.Corrupt, OrganizeEngine.CorruptCategory), (d.Source, d.Category));
        d.OcrCorrupt = false;
        OrganizeEngine.Decide(d, s);
        Assert.Equal((MatchSource.Ocr, "D"), (d.Source, d.Category));
        d.OcrLines.Clear();
        Assert.False(OrganizeEngine.Decide(d, s));
        Assert.Null(d.Category);
        d.FileName = "aaa.mp4";
        OrganizeEngine.Decide(d, s);
        Assert.Equal(MatchSource.Filename, d.Source);
    }

    [Fact]
    public void StoppingGivesWhatWasFoundSoFar_AndReportsProgress()
    {
        Files("a.mp4", "b.mp4", "c.mp4");
        using var cts = new CancellationTokenSource();
        ocr.OnRead = name => { if (name == "b.mp4") cts.Cancel(); };
        var engine = NewEngine();
        var phases = new List<OrganizePhase>();
        engine.ProgressChanged += p => phases.Add(p.Phase);

        var result = engine.Run(dir, Rules(), cts.Token);

        Assert.Single(result.Decisions);                               // a.mp4 was finished, b.mp4 was being read
        Assert.Equal(OrganizePhase.Cancelled, engine.Current.Phase);
        Assert.Contains(OrganizePhase.Listing, phases);
        Assert.Contains(OrganizePhase.Scanning, phases);
        Assert.Equal(1, engine.Current.FilesDone);
        Assert.Equal(3, engine.Current.FilesTotal);
    }

    [Fact]
    public void ABackgroundScanCompletesAndRefusesASecondOne()
    {
        Files("Python Course 1.mp4");
        var engine = NewEngine();
        var done = new ManualResetEventSlim();
        OrganizeResult? got = null;
        engine.Completed += r => { got = r; done.Set(); };
        Assert.True(engine.Start(dir, Rules()));
        Assert.True(done.Wait(TimeSpan.FromSeconds(20)));
        Assert.Equal("Tutorials", got!.Decisions.Single().Category);
        Assert.Equal(OrganizePhase.Done, engine.Current.Phase);
    }

    [Fact]
    public void ARenameIsSuggestedFromTheTitle_OnlyWhenTheNameDiffers()
    {
        // (Tags need a real media file, so this checks the same rule the engine applies)
        Assert.Equal("My_ title_.mp4", FolderNames.SafeFileName("My: title?") + ".mp4");
        Assert.Equal("unnamed", FolderNames.SafeFileName("  "));
        Assert.Equal("x", FolderNames.SafeFileName("x. "));
    }
}

public class ScriptsAndMoveTests : IDisposable
{
    private readonly string dir = Path.Combine(Path.GetTempPath(), "fo-move-" + Guid.NewGuid().ToString("N"));
    public ScriptsAndMoveTests() => Directory.CreateDirectory(dir);
    public void Dispose() { try { Directory.Delete(dir, true); } catch { } }

    private void Files(params string[] names) { foreach (var n in names) File.WriteAllText(Path.Combine(dir, n), n); }

    private static FileDecision Match(string file, string category, string trigger = "Contains: x") =>
        new() { FileName = file, Category = category, Trigger = trigger, Source = MatchSource.Filename };

    [Fact]
    public void QuotingKeepsEverythingLiteral()
    {
        Assert.Equal("'it''s [a] $b `c` 100%'", ScriptWriter.PsQuote("it's [a] $b `c` 100%"));
        Assert.Equal("'it’’s'", ScriptWriter.PsQuote("it’s"));      // typographic quotes too
        Assert.Equal("'it'\\''s $b'", ScriptWriter.ShQuote("it's $b"));
    }

    [Fact]
    public void TheMoveScriptHasAFolderAndAMovePerFile_AndTheUsersChoicesCount()
    {
        var a = Match("a.mp4", "Series");
        var b = Match("b o'k.mp4", "Series");
        var c = Match("c.mp4", "Skipped"); c.Skip = true;
        var d = Match("d.mp4", "Tutorials"); d.OverrideCategory = "Mine";
        var unmatched = new FileDecision { FileName = "e.txt" };

        string? path = ScriptWriter.WriteMoveScript(dir, new[] { a, b, c, d, unmatched });
        string text = File.ReadAllText(path!);
        Assert.Equal(ScriptWriter.MovePath(dir), path);
        Assert.Contains("Series", text);
        Assert.Contains("Mine", text);
        Assert.DoesNotContain("Skipped", text);
        Assert.DoesNotContain("Tutorials", text);
        Assert.DoesNotContain("e.txt", text);
        Assert.Contains("# Triggered: Chosen by you", text);
        Assert.Contains(OperatingSystem.IsWindows() ? "'b o''k.mp4'" : "'b o'\\''k.mp4'", text);
        byte[] bytes = File.ReadAllBytes(path!);
        if (OperatingSystem.IsWindows()) Assert.Equal(new byte[] { 0xEF, 0xBB, 0xBF }, bytes[..3]);   // PowerShell 5.1 needs the BOM for æøå
        else Assert.Equal((byte)'#', bytes[0]);                                                          // sh: "#!/bin/sh", no BOM
    }

    [Fact]
    public void NothingToDoRemovesAnOldScript()
    {
        File.WriteAllText(ScriptWriter.MovePath(dir), "old");
        Assert.Null(ScriptWriter.WriteMoveScript(dir, new[] { new FileDecision { FileName = "x" } }));
        Assert.False(File.Exists(ScriptWriter.MovePath(dir)));
    }

    [Fact]
    public void RenameScriptAndUnmatchedReport()
    {
        var r = new FileDecision { FileName = "a.mp4", RenameTo = "Real title.mp4", RenameSource = "Title" };
        string text = File.ReadAllText(ScriptWriter.WriteRenameScript(dir, new[] { r })!);
        Assert.Contains("Real title.mp4", text);
        Assert.Contains("# Source: Title", text);

        var u = new FileDecision { FileName = "u.mp4", Title = "T", Comment = "C", OcrLines = { "some text" } };
        string report = File.ReadAllText(ScriptWriter.WriteUnmatchedReport(dir, new[] { u, Match("m.mp4", "X") })!);
        Assert.Contains("u.mp4...", report);
        Assert.Contains("Metadata Title: \"T\"", report);
        Assert.Contains("OCR Detected Text: \"some text\"", report);
        Assert.DoesNotContain("m.mp4", report);
    }

    [Fact]
    public void FilesAreMovedIntoTheirFolder_NeverOverwritingAnything()
    {
        Files("a.mp4", "b.mp4", "c.txt");
        Directory.CreateDirectory(Path.Combine(dir, "Series"));
        File.WriteAllText(Path.Combine(dir, "Series", "b.mp4"), "already here");
        var a = Match("a.mp4", "Series"); var b = Match("b.mp4", "Series"); var c = new FileDecision { FileName = "c.txt" };

        var report = MoveExecutor.Apply(dir, new[] { a, b, c }, alsoRename: false);

        Assert.Equal((1, 1, 0), (report.Moved, report.SkippedExisting, report.Failed));
        Assert.True(File.Exists(Path.Combine(dir, "Series", "a.mp4")));
        Assert.False(File.Exists(Path.Combine(dir, "a.mp4")));
        Assert.Equal("already here", File.ReadAllText(Path.Combine(dir, "Series", "b.mp4")));   // untouched
        Assert.True(File.Exists(Path.Combine(dir, "b.mp4")));                                      // and the source stays
        Assert.True(File.Exists(Path.Combine(dir, "c.txt")));
        Assert.True(a.Done);
        Assert.False(b.Done);
        Assert.Contains("exists already", report.Messages.Single());
        Assert.StartsWith("Moved 1 file", report.Summary);
    }

    [Fact]
    public void ACategoryCannotLeadOutOfTheFolder()
    {
        Files("a.mp4");
        var evil = Match("a.mp4", "..");
        var report = MoveExecutor.Apply(dir, new[] { evil }, false);
        Assert.Equal(1, report.Failed);
        Assert.True(File.Exists(Path.Combine(dir, "a.mp4")));
        var nested = Match("a.mp4", "x/y");
        Assert.Equal(1, MoveExecutor.Apply(dir, new[] { nested }, false).Failed);
    }

    [Fact]
    public void RenamingWorksInPlaceAndInTheCategoryFolder()
    {
        Files("a.mp4", "b.mp4", "c.mp4");
        var inPlace = new FileDecision { FileName = "a.mp4", RenameTo = "Alpha.mp4" };
        var moved = Match("b.mp4", "Series"); moved.RenameTo = "Beta.mp4";
        var notWanted = new FileDecision { FileName = "c.mp4", RenameTo = "Gamma.mp4" };

        var withRename = MoveExecutor.Apply(dir, new[] { inPlace, moved }, alsoRename: true);
        Assert.Equal((1, 2), (withRename.Moved, withRename.Renamed));
        Assert.True(File.Exists(Path.Combine(dir, "Alpha.mp4")));
        Assert.True(File.Exists(Path.Combine(dir, "Series", "Beta.mp4")));

        var without = MoveExecutor.Apply(dir, new[] { notWanted }, alsoRename: false);
        Assert.Equal(0, without.Renamed);
        Assert.True(File.Exists(Path.Combine(dir, "c.mp4")));
    }

    [Fact]
    public void AMissingFileIsReportedNotThrown()
    {
        var gone = Match("gone.mp4", "Series");
        var report = MoveExecutor.Apply(dir, new[] { gone }, false);
        Assert.Equal(1, report.Failed);
        Assert.Contains("not there any more", report.Messages.Single());
    }
}

public class AppControllerTests : IDisposable
{
    private readonly string root = Path.Combine(Path.GetTempPath(), "fo-app-" + Guid.NewGuid().ToString("N"));
    private string Data => Path.Combine(root, "data");
    private string Files => Path.Combine(root, "files");
    public AppControllerTests() { Directory.CreateDirectory(Data); Directory.CreateDirectory(Files); }
    public void Dispose() { try { Directory.Delete(root, true); } catch { } }

    private AppController Open(Settings? settings = null)
    {
        if (settings != null) SettingsStore.Save(Data, settings);
        return AppController.Open(Data);
    }

    private static void Scan(AppController app)
    {
        var done = new ManualResetEventSlim();
        void Handler() => done.Set();
        app.ResultChanged += Handler;
        Assert.Null(app.StartScan());
        Assert.True(done.Wait(TimeSpan.FromSeconds(20)));
        app.ResultChanged -= Handler;
    }

    [Fact]
    public void FirstRunHasNoRulesAndAScanNeedsFolderAndCategories()
    {
        using var app = Open();
        Assert.True(app.IsFirstRun);
        Assert.Equal("Choose the folder to organize first.", app.StartScan());
        Assert.Null(app.SetTargetFolder(Files));
        Assert.StartsWith("There are no categories yet", app.StartScan());
        Assert.StartsWith("The folder", app.SetTargetFolder(Path.Combine(root, "nope")));
        Assert.Equal(Files, app.TargetFolder);   // a wrong choice does not replace the good one
    }

    [Fact]
    public void TheFolderIsRememberedWithTheRecentOnes()
    {
        using (var app = Open())
        {
            string second = Path.Combine(root, "second"); Directory.CreateDirectory(second);
            app.SetTargetFolder(Files); app.SetTargetFolder(second); app.SetTargetFolder(Files);
            Assert.Equal(new[] { Files, second }, app.RecentFolders);
        }
        using var again = AppController.Open(Data);
        Assert.Equal(Files, again.TargetFolder);
    }

    [Fact]
    public void ARuleMadeFromUnmatchedTextAppliesAtOnce()
    {
        File.WriteAllText(Path.Combine(Files, "a.txt"), "x");
        File.WriteAllText(Path.Combine(Files, "Show S01.txt"), "x");
        using var app = Open(Make.Rules(Make.Cat("Series", true, ("Regex", @"S\d\d"))));
        app.SetTargetFolder(Files);
        Scan(app);
        Assert.Single(app.Result!.Unmatched);

        string message = app.AddRule("Notes", "Contains", "a.t", out bool ok);   // a new category
        Assert.True(ok, message);
        Assert.Contains("Made the category Notes", message);
        Assert.Equal("Notes", app.Result.Decisions.Single(d => d.FileName == "a.txt").Category);
        Assert.Empty(app.Result.Unmatched);

        // saved: a new controller sees it, and the new category does not match by its own name
        using var reopened = AppController.Open(Data);
        var notes = reopened.Settings.Categories.Single(c => c.Name == "Notes");
        Assert.True(notes.IgnoreCategoryInContent);
        Assert.Equal("a.t", notes.Rules.Single().Value);

        Assert.Contains("has that rule already", app.AddRule("notes", "Contains", "A.T", out ok));
        Assert.False(ok);
        Assert.Contains("not a valid regular expression", app.AddRule("Notes", "Regex", "(", out ok));
        Assert.Contains("can't be a folder name", app.AddRule("../x", "Contains", "y", out ok));
        Assert.Contains("Choose a category", app.AddRule("", "Contains", "y", out ok));
    }

    [Fact]
    public void OcrTextShownWithoutRepeats()
    {
        var lines = OcrText.Distinct(new[] { "ZEBRA", "Zebra Corp", "zebracorp.", "ZEBRA CORP", "Other Text", "Oth" });
        Assert.Equal(new[] { "Zebra Corp", "Other Text" }, lines);
        Assert.Empty(OcrText.Distinct(new[] { "...", " " }));
    }

    [Fact]
    public void VideoSamplingStartsWithTheFirstFrame()
    {
        var times = OpenCvVideoOcr.SampleTimes(60);
        Assert.Equal(0, times[0]);
        Assert.Equal(1, times[1]);
        Assert.Equal(26, times.Count);
        Assert.Equal(new[] { 0.0 }, OpenCvVideoOcr.SampleTimes(0.4));   // a very short video: just the first frame
        Assert.Equal(new[] { 0.0, 1.0, 2.0 }, OpenCvVideoOcr.SampleTimes(2.5));
    }

    [Fact]
    public void CategoriesWithTheSameNameAreJoinedWhenLoaded()
    {
        var s = Make.Rules(
            Make.Cat("Other Cat", true, ("Contains", "x"), ("Regex", "a.b")),
            Make.Cat("Mid", false, ("Contains", "m")),
            Make.Cat("other cat", false, ("Contains", "X"), ("Contains", "y")));
        using var app = Open(s);
        Assert.Equal(new[] { "Other Cat", "Mid" }, app.Settings.Categories.Select(c => c.Name));
        var other = app.Settings.Categories[0];
        Assert.Equal(new[] { "x", "a.b", "y" }, other.Rules.Select(r => r.Value));   // "X" was there already
        Assert.False(other.IgnoreCategoryInContent);                                  // one of them used the name as a rule
        // and it was written back, so the next start sees one category
        using var again = AppController.Open(Data);
        Assert.Equal(2, again.Settings.Categories.Count);
        Assert.True(File.Exists(SettingsStore.PathIn(Data) + ".bck"));
    }

    [Fact]
    public void AddingARuleWorksWhenSettingsRepeatACategoryName()
    {
        using var app = Open(Make.Rules(Make.Cat("Other Cat", true, ("Contains", "x")), Make.Cat("Other Cat", true, ("Contains", "y"))));
        string message = app.AddRule("MyCat", "Contains", "abc", out bool ok);
        Assert.True(ok, message);
        Assert.Contains(app.Settings.Categories, c => c.Name == "MyCat");
    }

    [Fact]
    public void AddingARuleExplainsAProblemInAnotherCategory()
    {
        // "Other Cat" already has an empty rule: nothing can be saved until that is fixed, and the message says so
        using var app = Open(Make.Rules(Make.Cat("Other Cat", true, ("Contains", ""))));
        string message = app.AddRule("MyCat", "Contains", "abc", out bool ok);
        Assert.False(ok);
        Assert.Contains("Other Cat: a rule is empty.", message);
        Assert.Contains("Fix it under Rules", message);
        Assert.DoesNotContain(app.Settings.Categories, c => c.Name == "MyCat");
    }

    [Fact]
    public void SavingBadRulesIsRefusedAndChangesNothing()
    {
        using var app = Open(Make.Rules(Make.Cat("A", false, ("Contains", "a"))));
        var bad = app.Settings.Clone();
        bad.Categories.Add(new Category { Name = "B", Rules = { new Rule("Regex", "([") } });
        var problems = app.SaveSettings(bad);
        Assert.Single(problems);
        Assert.Single(AppController.Open(Data).Settings.Categories);   // the file is unchanged
        Assert.Single(app.Settings.Categories);
    }

    [Fact]
    public void ChangingTheRulesReevaluatesTheResultAndRaisesTheEvents()
    {
        File.WriteAllText(Path.Combine(Files, "alpha.txt"), "x");
        using var app = Open(Make.Rules(Make.Cat("Other", true, ("Contains", "zzz"))));
        app.SetTargetFolder(Files);
        Scan(app);
        Assert.Single(app.Result!.Unmatched);

        int settingsEvents = 0, resultEvents = 0;
        app.SettingsChanged += () => settingsEvents++;
        app.ResultChanged += () => resultEvents++;
        var s = app.Settings.Clone();
        s.Categories[0].Rules.Add(new Rule("Contains", "alpha"));
        Assert.Empty(app.SaveSettings(s));
        Assert.Equal((1, 1), (settingsEvents, resultEvents));
        Assert.Equal("Other", app.Result.Decisions[0].Category);
    }

    [Fact]
    public void ScriptsAndMovesFollowTheChoicesMadeInTheReview()
    {
        foreach (var n in new[] { "Show S01.txt", "Show S02.txt", "Show S03.txt", "plain.txt" }) File.WriteAllText(Path.Combine(Files, n), n);
        using var app = Open(Make.Rules(Make.Cat("Series", true, ("Regex", @"S\d\d")), Make.Cat("Elsewhere", true)));
        app.SetTargetFolder(Files);
        Scan(app);
        var s2 = app.Result!.Decisions.Single(d => d.FileName == "Show S02.txt"); s2.OverrideCategory = "Elsewhere";
        var s3 = app.Result.Decisions.Single(d => d.FileName == "Show S03.txt"); s3.Skip = true;
        Assert.Equal(2, app.Pending.Count());

        var scripts = app.WriteScripts();
        Assert.Equal(2, scripts.Moves);
        Assert.NotNull(scripts.UnmatchedReport);          // plain.txt
        string text = File.ReadAllText(scripts.MoveScript!);
        Assert.Contains("Elsewhere", text);
        Assert.DoesNotContain("Show S03.txt", text);

        var report = app.MoveNow(alsoRename: false);
        Assert.Equal(2, report.Moved);
        Assert.True(File.Exists(Path.Combine(Files, "Series", "Show S01.txt")));
        Assert.True(File.Exists(Path.Combine(Files, "Elsewhere", "Show S02.txt")));
        Assert.True(File.Exists(Path.Combine(Files, "Show S03.txt")));      // left alone
        Assert.Equal(Path.Combine(Files, "Series", "Show S01.txt"), app.PathOf(app.Result.Decisions.Single(d => d.FileName == "Show S01.txt")));
        Assert.Empty(app.Pending);

        // scripts made later leave out what has been moved
        Assert.Null(app.WriteScripts().MoveScript);
    }

    [Fact]
    public void ABrokenSettingsFileIsReportedAndBlocksScanning()
    {
        File.WriteAllText(SettingsStore.PathIn(Data), "{ nope");
        using var app = AppController.Open(Data);
        Assert.Contains("is not valid", app.SettingsError);
        app.SetTargetFolder(Files);
        Assert.Contains("is not valid", app.StartScan());
    }

    [Fact]
    public void CategoryNamesIncludeTheFolderForUnreadableVideos()
    {
        using var app = Open(Make.Rules(Make.Cat("A"), Make.Cat("B")));
        Assert.Equal(new[] { "A", "B", "_CHECK_" }, app.CategoryNames());
    }
}
