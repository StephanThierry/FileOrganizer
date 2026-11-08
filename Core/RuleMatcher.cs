using System.Collections.Concurrent;
using System.Text.RegularExpressions;

namespace FileOrganizer;

/// <summary>A rule that matched: the category and what matched (the text the scripts show).</summary>
public sealed record RuleMatch(string Category, string Trigger);

/// <summary>Matching text against the categories. The first category (in the order of the list) with a match wins.</summary>
public static class RuleMatcher
{
    private static readonly ConcurrentDictionary<string, Regex?> Cache = new();

    private static Regex? Compile(string pattern) => Cache.GetOrAdd(pattern, p =>
    {
        try { return new Regex(p, RegexOptions.IgnoreCase, TimeSpan.FromSeconds(2)); }
        catch (ArgumentException) { return null; }
    });

    /// <summary>
    /// The first match: a category whose name is found in the text (unless it ignores its name), or one of its rules.
    /// A broken regular expression matches nothing (Validate tells what is wrong with it).
    /// </summary>
    public static RuleMatch? Match(string? input, Settings settings)
    {
        if (string.IsNullOrEmpty(input)) return null;
        foreach (var category in settings.Categories)
        {
            foreach (var hit in Hits(input, category))
                return new RuleMatch(category.Name, hit);
        }
        return null;
    }

    /// <summary>Every rule of every category that matches, in priority order (for the "try some text" box). The first one wins.</summary>
    public static List<RuleMatch> Explain(string? input, Settings settings)
    {
        var all = new List<RuleMatch>();
        if (string.IsNullOrEmpty(input)) return all;
        foreach (var category in settings.Categories)
            foreach (var hit in Hits(input, category))
                all.Add(new RuleMatch(category.Name, hit));
        return all;
    }

    private static IEnumerable<string> Hits(string input, Category category)
    {
        // The name as a rule: ".com" is left out ("site.com" matches "site")
        string searchName = category.Name.Replace(".com", "", StringComparison.OrdinalIgnoreCase);
        if (!category.IgnoreCategoryInContent && searchName.Length > 0 && input.Contains(searchName, StringComparison.OrdinalIgnoreCase))
            yield return "Category Name Match";

        foreach (var rule in category.Rules)
        {
            if (string.IsNullOrEmpty(rule.Value)) continue;
            if (rule.Type.Equals("Regex", StringComparison.OrdinalIgnoreCase))
            {
                if (IsRegexMatch(input, rule.Value)) yield return $"Regex: {rule.Value}";
            }
            else if (input.Contains(rule.Value, StringComparison.OrdinalIgnoreCase))
            {
                yield return $"Contains: {rule.Value}";
            }
        }
    }

    private static bool IsRegexMatch(string input, string pattern)
    {
        var regex = Compile(pattern);
        if (regex == null) return false;
        try { return regex.IsMatch(input); }
        catch (RegexMatchTimeoutException) { return false; }
    }

    /// <summary>What is wrong with a rule (null: nothing).</summary>
    public static string? Validate(Rule rule)
    {
        if (string.IsNullOrWhiteSpace(rule.Value)) return "a rule is empty.";
        if (!rule.Type.Equals("Regex", StringComparison.OrdinalIgnoreCase)) return null;
        try { _ = new Regex(rule.Value, RegexOptions.IgnoreCase, TimeSpan.FromSeconds(2)); return null; }
        catch (ArgumentException e) { return $"\"{rule.Value}\" is not a valid regular expression ({e.Message})."; }
    }
}
