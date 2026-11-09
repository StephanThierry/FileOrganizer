using FileOrganizer;
using Xunit;

namespace CoreTests;

public class RegexGuideTests
{
    [Fact]
    public void ThePatternsYourRulesUseComeFirst_MostUsedFirst()
    {
        var s = Make.Rules(Make.Cat("A", true,
            ("Regex", "night.?drive"), ("Regex", "real.?life"), ("Regex", @"AH\d{3,4}"), ("Regex", "^start"), ("Contains", "ignored.?")));
        var tips = RegexGuide.Tips(s);

        Assert.Equal(".?", tips[0].Pattern);                    // used twice
        Assert.Equal(2, tips[0].UsedBy);
        int firstUnused = tips.FindIndex(t => t.UsedBy == 0);
        Assert.True(firstUnused > 0);
        Assert.All(tips.Take(firstUnused), t => Assert.True(t.UsedBy > 0));
        Assert.All(tips.Skip(firstUnused), t => Assert.Equal(0, t.UsedBy));   // used ones first, then the extra ones
        Assert.Contains(tips, t => t.Pattern == "(a|b)" && t.UsedBy == 0);
        Assert.Contains(tips, t => t.Pattern == "^" && t.UsedBy == 1);
        Assert.Contains(tips, t => t.Pattern == @"\d{3,4}" && t.UsedBy == 1);
    }

    [Fact]
    public void ARulesFile_WithoutRegexRules_StillGivesTheGuide()
    {
        var tips = RegexGuide.Tips(new Settings());
        Assert.True(tips.Count >= 10);
        Assert.All(tips, t => { Assert.Equal(0, t.UsedBy); Assert.NotEmpty(t.Meaning); Assert.NotEmpty(t.Example); });
    }

    [Fact]
    public void TheTestTextHighlightsWhatThePatternSelects()
    {
        var (pieces, message) = RegexGuide.Highlight("night.?drive", "A Night-Drive and a nightdrive");
        Assert.Equal(new[] { "A ", "Night-Drive", " and a ", "nightdrive" }, pieces!.Select(p => p.Text));
        Assert.Equal(new[] { false, true, false, true }, pieces!.Select(p => p.IsMatch));
        Assert.StartsWith("Matches 2 parts", message);
    }

    [Fact]
    public void NoMatchBrokenPatternAndEmptyInput_AreSaidInWords()
    {
        var (none, noMatch) = RegexGuide.Highlight("zzz", "nothing here");
        Assert.Equal("nothing here", Assert.Single(none!).Text);
        Assert.StartsWith("No match", noMatch);

        var (broken, why) = RegexGuide.Highlight("([unclosed", "text");
        Assert.Null(broken);
        Assert.Contains("not a valid regular expression", why);

        Assert.Null(RegexGuide.Highlight("", "text").Pieces);
        Assert.Empty(RegexGuide.Highlight("abc", "").Pieces!);
    }

    [Fact]
    public void AnEmptyMatchSelectsNothing()
    {
        var (pieces, message) = RegexGuide.Highlight("x*", "abc");
        Assert.All(pieces!, p => Assert.False(p.IsMatch));
        Assert.Contains("nothing is selected", message);
    }

    // What the guide says each pattern finds really is what it finds
    [Theory]
    [InlineData("night.?drive", "nightdrive", true)]
    [InlineData("night.?drive", "night-drive", true)]
    [InlineData("night.?drive", "night--drive", false)]
    [InlineData("real.life", "reallife", false)]
    [InlineData("real.life", "real_life", true)]
    [InlineData("^bang", "bang bus 12.mp4", true)]
    [InlineData("^bang", "the big bang.mp4", false)]
    [InlineData(@"AH\d{3,4}", "AH1234", true)]
    [InlineData(@"AH\d{3,4}", "AH12", false)]
    [InlineData("^drops$", "drops", true)]
    [InlineData("^drops$", "drops 2", false)]
    [InlineData("premium.*party", "premiumparty", true)]
    [InlineData(@"18[ \.]+only", "18 . only", true)]
    [InlineData(@"\.mp4$", "film.mp4", true)]
    [InlineData(@"\.mp4$", "filmxmp4", false)]
    [InlineData("(WoW|Minecraft)", "Minecraft 3.mp4", true)]
    [InlineData(@"\bart\b", "my art video", true)]
    [InlineData(@"\bart\b", "party", false)]
    [InlineData(@"S\d\dE\d\d", "Show S01E07", true)]
    [InlineData(@"(19|20)\d\d", "Trip 2019", true)]
    [InlineData("^(?!.*trailer)", "plain video", true)]
    [InlineData("^(?!.*trailer)x", "x trailer", false)]
    [InlineData("(?-i)^AB", "ab 1.mp4", false)]
    public void TheExamplesInTheGuideHold(string pattern, string text, bool matches)
    {
        Assert.Null(RuleMatcher.Validate(new Rule("Regex", pattern)));
        Assert.Equal(matches, RuleMatcher.Match(text, Make.Rules(Make.Cat("X", true, ("Regex", pattern)))) != null);
    }
}

public class OcrLinesTests
{
    private static OcrWord W(string t, float left, float right, float top, int line) => new(t, left, right, top, top + 40, line);

    [Fact]
    public void WordsOnTheSameRow_BecomeOneFinding_AndRowsStaySeparate()
    {
        var lines = OcrLines.Join(new[]
        {
            W("my", 38, 116, 118, 0), W("own", 139, 235, 121, 0), W("summer", 261, 465, 121, 0),
            W("Holiday", 36, 221, 204, 1), W("Films", 243, 378, 206, 1),
        });
        Assert.Equal(new[] { "my own summer", "Holiday Films" }, lines);
    }

    [Fact]
    public void TheWordsAreJoinedLeftToRight_WhateverTheOrderTheyCameIn()
    {
        Assert.Equal(new[] { "my own summer" }, OcrLines.Join(new[] { W("summer", 261, 465, 121, 0), W("my", 38, 116, 118, 0), W("own", 139, 235, 121, 0) }));
    }

    [Fact]
    public void TextFarApartOnTheSameRow_IsTwoFindings()
    {
        // A logo at the left and a caption at the right of the picture: more than three text heights apart
        Assert.Equal(new[] { "site.com", "my summer" }, OcrLines.Join(new[] { W("site.com", 10, 150, 100, 0), W("my", 560, 600, 100, 0), W("summer", 610, 700, 100, 0) }));
    }

    [Fact]
    public void AWordWithoutARow_StandsAlone()
    {
        Assert.Equal(new[] { "a", "b" }, OcrLines.Join(new[] { new OcrWord("a", 0, 0, 0, 0, -1), new OcrWord("b", 0, 0, 0, 0, -2) }));
    }
}

public class AllOfRuleTests
{
    private static Rule All(params (string type, string value)[] terms) => Rule.AllOf(terms.Select(t => new RuleTerm(t.type, t.value)));
    private static Settings WithAll(Rule rule, string category = "Hub") => new() { Categories = { new Category { Name = category, IgnoreCategoryInContent = true, Rules = { rule } } } };

    [Fact]
    public void EveryTermMustBeFound_InAnyOfTheTexts_ContainsAndRegexMixed()
    {
        var rule = All(("Regex", "R.?C"), ("Contains", "hub"), ("Contains", "community"));
        // REC, R3C and RC all fit the pattern; the terms come from different findings
        foreach (var first in new[] { "REC", "R3C", "RC" })
            Assert.NotNull(RuleMatcher.MatchVideo(new[] { first, "hub", "the community" }, new[] { first, "hub", "the community" }, WithAll(rule)));
        Assert.Null(RuleMatcher.MatchVideo(new[] { "REC", "hub" }, new[] { "REC", "hub" }, WithAll(rule)));              // community is missing
        Assert.Null(RuleMatcher.MatchVideo(new[] { "RXXC", "hub", "community" }, new[] { "RXXC", "hub", "community" }, WithAll(rule)));
    }

    [Fact]
    public void TheTermsMayBeSpreadOverTheFileNameTagsAndVideoText()
    {
        var rule = All(("Contains", "summer"), ("Contains", "hub"));
        var m = RuleMatcher.MatchVideo(new[] { "hub" }, new[] { "my summer.mp4", "", "", "hub" }, WithAll(rule));
        Assert.True(m!.IsAll);
        Assert.Equal("All of: summer + hub", m.Trigger);
    }

    [Fact]
    public void CategoriesAreCheckedInPriorityOrder_AnAllRuleBeatsALaterSingleRule()
    {
        var s = Make.Rules(Make.Cat("Later", true, ("Contains", "hub")));
        s.Categories.Insert(0, new Category { Name = "First", IgnoreCategoryInContent = true, Rules = { All(("Contains", "rec"), ("Contains", "hub")) } });
        var lines = new[] { "rec", "hub" };
        Assert.Equal("First", RuleMatcher.MatchVideo(lines, lines, s)!.Category);
        Assert.Equal("Later", RuleMatcher.MatchVideo(new[] { "hub" }, new[] { "hub" }, s)!.Category);
    }

    [Fact]
    public void AnAllRule_IsIgnoredWhenOneTextIsMatched_AndValidated()
    {
        var s = WithAll(All(("Contains", "a"), ("Contains", "b")));
        Assert.Null(RuleMatcher.Match("a b", s));                                    // single texts never fit an All rule
        Assert.Single(RuleMatcher.Explain("a and b", s));                             // but "try some text" treats it as one text
        Assert.Contains("at least two", RuleMatcher.Validate(All(("Contains", "a"))));
        Assert.Contains("is empty", RuleMatcher.Validate(All(("Contains", "a"), ("Contains", " "))));
        Assert.Contains("not a valid regular expression", RuleMatcher.Validate(All(("Contains", "a"), ("Regex", "([x"))));
        Assert.Null(RuleMatcher.Validate(All(("Contains", "a"), ("Regex", "R.?C"))));
    }

    [Fact]
    public void TheDecisionUsesIt_AndTheRuleSurvivesSavingAndLoading()
    {
        var s = WithAll(All(("Regex", "R.?C"), ("Contains", "hub")));
        var d = new FileDecision { FileName = "clip.mp4", OcrLines = { "R3C", "Hub" } };
        Assert.True(OrganizeEngine.Decide(d, s));
        Assert.Equal("Hub", d.Category);
        Assert.Equal(MatchSource.Ocr, d.Source);
        Assert.Equal("OCR Match: 'All of: /R.?C/ + hub'", d.Trigger);

        var copy = s.Clone();
        copy.Normalize();
        Assert.Equal("All", copy.Categories[0].Rules[0].Type);
        Assert.Equal(new[] { "Regex:R.?C", "Contains:hub" }, copy.Categories[0].Rules[0].Terms!.Select(t => t.Type + ":" + t.Value));
        // Other rules keep no terms, so the file stays as it was
        Assert.DoesNotContain("Terms", System.Text.Json.JsonSerializer.Serialize(new Rule("Contains", "x"), SettingsStore.Json));
    }

    [Fact]
    public void TwoAllRulesWithTheSameTermsAreTheSameRule_AndOthersAreNot()
    {
        Assert.True(All(("Contains", "A"), ("Regex", "b")).SameAs(All(("Regex", "b"), ("Contains", "a"))));
        Assert.False(All(("Contains", "a"), ("Contains", "b")).SameAs(All(("Contains", "a"), ("Contains", "c"))));
        Assert.False(All(("Contains", "a"), ("Contains", "b")).SameAs(new Rule("Contains", "")));
    }
}

public class OcrFindingsTests
{
    private static readonly string[] Logo =
    {
        "Ppp.tV", "p-p-p.tv", "P-p-p.tv", "P-pp.t", "P-P-p.tv", "© R", "Pp-p.t", "Pp-p.tv", "P-P-P.TV", "P-P-PTV", "P-P-P. TV", "SLUT",
    };

    [Fact]
    public void TheReadingsOfOneText_AreOneFinding_WithEveryDifferentReadingKept()
    {
        var findings = OcrText.Findings(Logo);
        Assert.Equal(3, findings.Count);                                    // the logo, "© R" and SLUT
        var logo = findings[0];
        // The same text in other capitals is one reading: P-P-p.tv, P-p-p.tv and p-p-p.tv are the same to a rule (capitals don't matter)
        Assert.Equal(new[] { "Ppp.tV", "p-p-p.tv", "P-pp.t", "Pp-p.t", "Pp-p.tv", "P-P-PTV", "P-P-P. TV" }.OrderBy(x => x, StringComparer.OrdinalIgnoreCase),
                     logo.Variants.Select(v => v.Text).OrderBy(x => x, StringComparer.OrdinalIgnoreCase));
        Assert.Equal(Logo.Length - 2, logo.Variants.Sum(v => v.Count));      // every reading of it (all but "© R" and SLUT) was counted once
        Assert.Equal("p-p-p.tv", logo.Best, ignoreCase: true);               // the most complete: the most letters, then the most often seen
        Assert.Equal("SLUT", findings[2].Best);
    }

    [Fact]
    public void ReadingsAreCountedAndTheMostSeenComesFirst()
    {
        var f = OcrText.Findings(new[] { "Zebra", "Zebra Corp", "ZEBRA CORP", "zebracorp.", "Zebra Corp" }).Single();
        Assert.Equal("Zebra Corp", f.Best);
        Assert.Equal(new[] { ("Zebra Corp", 3), ("zebracorp.", 1), ("Zebra", 1) }, f.Variants.Select(v => (v.Text, v.Count)));
    }

    [Fact]
    public void TextWithoutLettersOrDigitsIsLeftOut_AndTheOrderIsTheOrderSeen()
    {
        Assert.Empty(OcrText.Findings(new[] { "...", " ", "" }));
        Assert.Equal(new[] { "Other Text", "Zebra Corp" }, OcrText.Findings(new[] { "Other Text", "Oth", "Zebra", "Zebra Corp" }).Select(f => f.Best));
    }
}
