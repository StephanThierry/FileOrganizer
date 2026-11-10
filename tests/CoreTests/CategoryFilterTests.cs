using FileOrganizer;
using Xunit;

namespace CoreTests;

public class CategoryFilterTests
{
    private static List<Category> Sample() => Make.Rules(
        Make.Cat("Tutorials", false, ("Contains", "Course"), ("Contains", "Lesson")),
        Make.Cat("Barry", true, ("Regex", @"^Barry.*\.mp4$"), ("Contains", "barrytv")),
        Make.Cat("Series", true, ("Regex", @"S\d\dE\d\d")),
        Make.Cat("Night drives", true, ("Regex", "night.?drive")),
        Make.Cat("Zebra tutorials", true, ("Contains", "zebra"))).Categories;

    [Fact]
    public void NoText_IsEveryCategory_Alphabetical()
    {
        var hits = CategoryFilter.Apply(Sample(), "  ");
        Assert.Equal(new[] { "Barry", "Night drives", "Series", "Tutorials", "Zebra tutorials" }, hits.Select(h => h.Category.Name));
        Assert.All(hits, h => Assert.Null(h.Reason));
    }

    [Fact]
    public void TheNameComesFirst_NamesStartingWithTheTextBeforeOthers_ThenTheRulesStrings()
    {
        // "tut": Tutorials starts with it, Zebra tutorials has it inside; no other category has a rule with it
        var hits = CategoryFilter.Apply(Sample(), "tut");
        Assert.Equal(new[] { "Tutorials", "Zebra tutorials" }, hits.Select(h => h.Category.Name));
        Assert.All(hits, h => Assert.Null(h.Reason));                         // both by name

        // "bar": Barry by its name; nothing else
        Assert.Equal(new[] { "Barry" }, CategoryFilter.Apply(Sample(), "bar").Select(h => h.Category.Name));

        // "course" is in no name: the category whose rule has it comes, with that rule as the reason
        var byRule = CategoryFilter.Apply(Sample(), "course").Single();
        Assert.Equal(("Tutorials", "Contains: Course"), (byRule.Category.Name, byRule.Reason));

        // A name hit and a rule hit together: the name first, even though the other is earlier in the alphabet
        var both = CategoryFilter.Apply(Sample(), "zebra");
        Assert.Equal(new[] { "Zebra tutorials" }, both.Select(h => h.Category.Name));   // its name; and its rule is the same text: still one hit
    }

    [Fact]
    public void ARulesPatternWithTheTextInIt_Fits_AndSoDoesARuleThatMatchesTheText()
    {
        var hits = CategoryFilter.Apply(Sample(), "drive");
        Assert.Equal(new[] { "Night drives" }, hits.Select(h => h.Category.Name));       // by its name

        // "night" is in the name too; "S02E05" is no name or pattern text, but the Series rule matches it as a file name
        var series = CategoryFilter.Apply(Sample(), "Season S02E05 of the show").Single();
        Assert.Equal(("Series", @"Regex: S\d\dE\d\d"), (series.Category.Name, series.Reason));

        // The pattern's own text is searched too: "^barry" is in Barry's rule string (and not in any name)
        var barry = CategoryFilter.Apply(Sample(), "^Barry").Single();
        Assert.Equal(("Barry", @"Regex: ^Barry.*\.mp4$"), (barry.Category.Name, barry.Reason));
    }

    [Fact]
    public void RuleHitsFollowTheNameHits_EachGroupAlphabetical_AndAllOfRulesAreSearched()
    {
        var cats = Make.Rules(
            Make.Cat("Zed", true, ("Contains", "lisbon trip")),
            Make.Cat("Alpha", true, ("Contains", "my lisbon video")),
            Make.Cat("Lisbon films", true, ("Contains", "nothing"))).Categories;
        cats.Add(new Category { Name = "Mixed", IgnoreCategoryInContent = true, Rules = { Rule.AllOf(new[] { new RuleTerm("Contains", "rec"), new RuleTerm("Regex", "lisb.n") }) } });

        var hits = CategoryFilter.Apply(cats, "lisbon");
        Assert.Equal(new[] { "Lisbon films", "Alpha", "Zed" }, hits.Select(h => h.Category.Name));    // the name first, then the others alphabetical
        Assert.Null(hits[0].Reason);
        Assert.Equal("Contains: my lisbon video", hits[1].Reason);

        Assert.Equal("All of: rec + /lisb.n/", CategoryFilter.Apply(cats, "rec").Single().Reason);   // a term of an All of rule has the text
        Assert.Empty(CategoryFilter.Apply(cats, "zzzqqq"));
    }
}
