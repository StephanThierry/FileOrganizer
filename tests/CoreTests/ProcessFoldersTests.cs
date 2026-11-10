using FileOrganizer;
using Xunit;

namespace CoreTests;

public class ProcessFoldersTests : IDisposable
{
    private readonly string dir = Path.Combine(Path.GetTempPath(), "fo-folders-" + Guid.NewGuid().ToString("N"));
    private readonly FakeOcr ocr = new();
    public ProcessFoldersTests() => Directory.CreateDirectory(dir);
    public void Dispose() { try { Directory.Delete(dir, true); } catch { } }

    private void File_(string relative, string content = "x")
    {
        string path = Path.Combine(dir, relative);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, content);
    }

    private static Settings Rules(bool processFolders)
    {
        var s = Make.Rules(
            Make.Cat("President Curtis", false),
            Make.Cat("Barry", true, ("Regex", @"^Barry.*\.mp4$")));
        s.Opt.ProcessFolders = processFolders;
        return s;
    }

    private OrganizeResult Scan(bool processFolders) => new OrganizeEngine(() => ocr).Run(dir, Rules(processFolders), CancellationToken.None);

    private void Layout()
    {
        File_("loose.txt");
        File_(Path.Combine("President Curtis S01E12 x265", "episode.mkv"));
        File_(Path.Combine("President Curtis S01E12 x265", "RARBG.txt"));
        File_(Path.Combine("President Curtis S01E12 x265", "Subs", "en.srt"));
        File_(Path.Combine("President Curtis S01E12 x265", "Sample", "deep", "clip.mkv"));
        File_(Path.Combine("Holiday photos", "img.jpg"));                    // its name fits no category
    }

    [Fact]
    public void Off_OnlyTheFilesDirectlyInTheFolderAreLookedAt()
    {
        Layout();
        var result = Scan(false);
        Assert.Equal(new[] { "loose.txt" }, result.Decisions.Select(d => d.FileName));
        Assert.All(result.Decisions, d => Assert.Null(d.FolderName));
    }

    [Fact]
    public void On_AFolderWhoseNameFitsACategory_IsOpened_AtAnyDepth_AndItsFilesGoWithIt()
    {
        Layout();
        var result = Scan(true);

        var inside = result.Decisions.Where(d => d.FolderName != null).OrderBy(d => d.RelativePath, StringComparer.OrdinalIgnoreCase).ToList();
        Assert.Equal(new[] { "episode.mkv", "RARBG.txt", "clip.mkv", "en.srt" }.OrderBy(x => x), inside.Select(d => d.FileName).OrderBy(x => x));
        Assert.All(inside, d =>
        {
            Assert.Equal("President Curtis", d.Category);
            Assert.Equal(MatchSource.Folder, d.Source);
            Assert.Equal("President Curtis S01E12 x265", d.FolderName);
            Assert.StartsWith("Folder: President Curtis S01E12 x265", d.Trigger);
        });
        Assert.Equal(Path.Combine("President Curtis S01E12 x265", "Sample", "deep"), inside.Single(d => d.FileName == "clip.mkv").SubFolder);
        Assert.Equal(Path.Combine("President Curtis S01E12 x265", "Subs", "en.srt"), inside.Single(d => d.FileName == "en.srt").RelativePath);

        Assert.DoesNotContain(result.Decisions, d => d.RelativePath.StartsWith("Holiday photos"));   // a folder that fits nothing is left alone
        Assert.Null(result.Decisions.Single(d => d.FileName == "loose.txt").Category);                 // as before
        Assert.Empty(ocr.Read_);                                                                       // the videos in the folder were not read: the folder says where they go
    }

    [Fact]
    public void AFilesOwnMatchWins_TheFolderIsOnlyTheFallback()
    {
        File_(Path.Combine("President Curtis S01E12 x265", "Barry S01E01.mp4"));
        File_(Path.Combine("President Curtis S01E12 x265", "plain.mp4"));
        var result = Scan(true);
        var own = result.Decisions.Single(d => d.FileName == "Barry S01E01.mp4");
        Assert.Equal(("Barry", MatchSource.Filename), (own.Category, own.Source));
        Assert.Equal("President Curtis", result.Decisions.Single(d => d.FileName == "plain.mp4").Category);
    }

    [Fact]
    public void ACategorysOwnFolder_AndTheCheckFolder_AreNeverOpened()
    {
        File_(Path.Combine("President Curtis", "organized.mkv"));      // the category's folder: organized already
        File_(Path.Combine("_CHECK_", "broken.mkv"));
        File_(Path.Combine("President Curtis 2019", "new.mkv"));       // a different name that fits: opened
        var result = Scan(true);
        Assert.Equal(new[] { "new.mkv" }, result.Decisions.Select(d => d.FileName));
    }

    [Fact]
    public void TheFilesAreMovedUpFlatIntoTheCategoryFolder_AClashStaysWhereItIs()
    {
        File_(Path.Combine("President Curtis A", "sample.mkv"), "a");
        File_(Path.Combine("President Curtis B", "sample.mkv"), "b");
        File_(Path.Combine("President Curtis B", "Subs", "en.srt"));
        var result = Scan(true);

        var report = MoveExecutor.Apply(dir, result.Decisions, alsoRename: false);

        Assert.Equal(2, report.Moved);                 // the first sample.mkv and en.srt
        Assert.Equal(1, report.SkippedExisting);       // the second sample.mkv: the name is taken in the category folder, nothing is overwritten
        Assert.True(File.Exists(Path.Combine(dir, "President Curtis", "en.srt")));
        string moved = File.ReadAllText(Path.Combine(dir, "President Curtis", "sample.mkv"));
        Assert.Equal(1, new[] { "a", "b" }.Count(c => c == moved));
        Assert.Contains(report.Messages, m => m.Contains("exists already") && m.Contains("sample.mkv"));
        string stays = moved == "a" ? Path.Combine(dir, "President Curtis B", "sample.mkv") : Path.Combine(dir, "President Curtis A", "sample.mkv");
        Assert.True(File.Exists(stays));
        Assert.False(File.Exists(Path.Combine(dir, "President Curtis A", "sample.mkv")) && File.Exists(Path.Combine(dir, "President Curtis B", "sample.mkv")));
    }

    [Fact]
    public void WhatIsLeftInTheOpenedFolders_IsSummarizedAfterTheMoves()
    {
        File_(Path.Combine("President Curtis S01E12 x265", "ep.mkv"));
        File_(Path.Combine("President Curtis S01E12 x265", "notes.txt"), "keep me");
        File_(Path.Combine("President Curtis S01E12 x265", "Subs", "en.srt"));
        File_(Path.Combine("President Curtis S02E01", "second.mkv"));  // moves too: nothing will be left in this folder
        File_(Path.Combine("President Curtis S03", "x.mkv"));
        var result = Scan(true);
        result.Decisions.Single(d => d.FileName == "notes.txt").Choose(FileDecision.LeaveHere);     // the user leaves this one
        MoveExecutor.Apply(dir, result.Decisions, alsoRename: false);
        Directory.Delete(Path.Combine(dir, "President Curtis S03"), true);                             // deleted afterwards

        var names = result.Decisions.Where(d => d.FolderName != null).Select(d => d.FolderName!).Distinct().ToList();
        var report = FolderLeftovers.Build(dir, names);

        var first = report.Folders.Single(f => f.Name == "President Curtis S01E12 x265");
        Assert.Equal(new[] { "notes.txt" }, first.Files.Select(f => f.Path));
        Assert.Equal(new[] { "Subs" }, first.EmptyFolders);             // the subtitle moved out and left an empty folder
        Assert.False(first.IsEmpty);
        Assert.True(report.Folders.Single(f => f.Name == "President Curtis S02E01").IsEmpty);
        Assert.False(report.Folders.Single(f => f.Name == "President Curtis S03").Exists);

        Assert.Equal(1, report.EmptyCount);
        Assert.Equal(1, report.FilesLeft);
        Assert.Equal("Of the 3 folders that were opened: 1 folder is empty; 1 file (0 KB) left in 1 folder; 1 folder was deleted.", report.Headline);
        Assert.StartsWith("President Curtis S01E12 x265: 1 file left (0 KB): notes.txt (0 KB).", report.Lines()[0]);
    }

    [Fact]
    public void ChangingTheRulesLater_StillFollowsTheFolder()
    {
        var d = new FileDecision { FileName = "episode.mkv", FolderName = "President Curtis S01E12 x265", SubFolder = "President Curtis S01E12 x265" };
        Assert.True(OrganizeEngine.Decide(d, Rules(true)));
        Assert.Equal("President Curtis", d.Category);

        var other = Make.Rules(Make.Cat("Somebody else", true, ("Contains", "somebody")));
        Assert.False(OrganizeEngine.Decide(d, other));                 // no rule fits the folder's name any more: the file has no category
        Assert.Null(d.Category);
    }

    [Fact]
    public void TheScriptsUseThePathInsideTheFolder()
    {
        var d = new FileDecision { FileName = "en.srt", SubFolder = Path.Combine("President Curtis S01E12 x265", "Subs"), Category = "President Curtis", Source = MatchSource.Folder, Trigger = "Folder: x" };
        string? path = ScriptWriter.WriteMoveScript(dir, new[] { d });
        Assert.NotNull(path);
        string script = File.ReadAllText(path!);
        string relative = Path.Combine("President Curtis S01E12 x265", "Subs", "en.srt");
        Assert.Contains(relative, script);                              // the source is the path inside the folder
        Assert.Contains("en.srt", script);
        Assert.Contains("President Curtis", script);                    // the target is the category folder
    }
}
