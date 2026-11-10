namespace FileOrganizer;

/// <summary>A category that fits the filter, and why when it is not by its name (the rule string that fits).</summary>
public sealed record CategoryHit(Category Category, string? Reason);

/// <summary>
/// The "Filter category" box on the Rules screen. Without text: every category, alphabetical. With text: first the categories whose
/// name has the text in it (the ones that start with it first), then the ones with a rule that fits, each with the rule as the reason.
/// Each group is alphabetical. A rule fits when its text or pattern has the typed text in it ("night" finds <c>night.?drive</c>),
/// or when the rule matches the typed text itself (a file name or a line from a video, to see which category it would go to).
/// </summary>
public static class CategoryFilter
{
    public static List<CategoryHit> Apply(IEnumerable<Category> categories, string? text)
    {
        var order = StringComparer.CurrentCultureIgnoreCase;
        var all = categories.OrderBy(c => c.Name ?? "", order).ToList();
        text = (text ?? "").Trim();
        if (text.Length == 0) return all.Select(c => new CategoryHit(c, null)).ToList();

        var byName = all.Where(c => (c.Name ?? "").Contains(text, StringComparison.CurrentCultureIgnoreCase))
            .OrderBy(c => (c.Name ?? "").StartsWith(text, StringComparison.CurrentCultureIgnoreCase) ? 0 : 1)   // stable: alphabetical inside each
            .ToList();
        var hits = byName.Select(c => new CategoryHit(c, null)).ToList();

        var named = byName.ToHashSet();
        foreach (var category in all.Where(c => !named.Contains(c)))
            if (RuleThatFits(category, text) is { } reason) hits.Add(new CategoryHit(category, reason));
        return hits;
    }

    // The first rule of the category that fits, as text ("Contains: lesson", "Regex: S\d\dE\d\d", "All of: rec + hub"); null: none does
    private static string? RuleThatFits(Category category, string text)
    {
        foreach (var rule in category.Rules ?? new())
        {
            bool has = rule.IsAll
                ? (rule.Terms ?? new()).Any(t => (t.Value ?? "").Contains(text, StringComparison.CurrentCultureIgnoreCase))
                : (rule.Value ?? "").Contains(text, StringComparison.CurrentCultureIgnoreCase);
            if (has) return Describe(rule);
        }
        // No rule has the text in it: does a rule match the text (as a file name, a tag or a line of text from a video)?
        var settings = new Settings { Categories = { category } };
        foreach (var hit in RuleMatcher.Explain(text, settings))
            if (hit.Trigger != "Category Name Match") return hit.Trigger;
        return null;
    }

    private static string Describe(Rule rule) => rule.IsAll ? rule.Describe() : $"{(rule.Type.Equals("Regex", StringComparison.OrdinalIgnoreCase) ? "Regex" : "Contains")}: {rule.Value}";
}
