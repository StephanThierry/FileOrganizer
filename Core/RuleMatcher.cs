using System.Collections.Concurrent;
using System.Text.RegularExpressions;

namespace FileOrganizer;

/// <summary>A rule that matched: the category and what matched (the text the scripts show).</summary>
public sealed record RuleMatch(string Category, string Trigger, bool IsAll = false);

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
        var texts = new[] { input };
        foreach (var category in settings.Categories)
        {
            foreach (var hit in Hits(input, category))
                all.Add(new RuleMatch(category.Name, hit));
            foreach (var rule in category.Rules.Where(r => AllOfMatches(r, texts)))
                all.Add(new RuleMatch(category.Name, rule.Describe(), true));
        }
        return all;
    }

    /// <summary>
    /// The first match for a video: walking the categories in priority order, a category matches when one of the lines read from
    /// the video fits one of its rules, or one of its "All" rules finds every term somewhere in <paramref name="texts"/> (the file
    /// name, the tags and the lines read from the video together).
    /// </summary>
    public static RuleMatch? MatchVideo(IReadOnlyList<string> lines, IReadOnlyList<string> texts, Settings settings)
    {
        foreach (var category in settings.Categories)
        {
            foreach (var line in lines)
                foreach (var hit in Hits(line, category))
                    return new RuleMatch(category.Name, hit);
            foreach (var rule in category.Rules)
                if (AllOfMatches(rule, texts)) return new RuleMatch(category.Name, rule.Describe(), true);
        }
        return null;
    }

    /// <summary>An "All" rule: every term is found in at least one of the texts. Other rules, and rules without terms, never match here.</summary>
    public static bool AllOfMatches(Rule rule, IReadOnlyList<string> texts)
    {
        if (!rule.IsAll || rule.Terms == null || rule.Terms.Count == 0) return false;
        return rule.Terms.All(t => !string.IsNullOrEmpty(t.Value) && texts.Any(text => TermMatches(t, text)));
    }

    private static bool TermMatches(RuleTerm term, string text) =>
        !string.IsNullOrEmpty(text) && (term.IsRegex ? IsRegexMatch(text, term.Value) : text.Contains(term.Value, StringComparison.OrdinalIgnoreCase));

    private static IEnumerable<string> Hits(string input, Category category)
    {
        // The name as a rule: ".com" is left out ("site.com" matches "site")
        string searchName = category.Name.Replace(".com", "", StringComparison.OrdinalIgnoreCase);
        if (!category.IgnoreCategoryInContent && searchName.Length > 0 && input.Contains(searchName, StringComparison.OrdinalIgnoreCase))
            yield return "Category Name Match";

        foreach (var rule in category.Rules)
        {
            if (rule.IsAll || string.IsNullOrEmpty(rule.Value)) continue;   // "All" rules look at the whole file, see MatchVideo
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
        if (rule.IsAll)
        {
            var terms = rule.Terms ?? new();
            if (terms.Count < 2) return "a rule with several terms needs at least two.";
            foreach (var t in terms)
            {
                if (string.IsNullOrWhiteSpace(t.Value)) return "a term of a rule with several terms is empty.";
                if (!t.IsRegex) continue;
                try { _ = new Regex(t.Value, RegexOptions.IgnoreCase, TimeSpan.FromSeconds(2)); }
                catch (ArgumentException e) { return $"\"{t.Value}\" is not a valid regular expression ({e.Message})."; }
            }
            return null;
        }
        if (string.IsNullOrWhiteSpace(rule.Value)) return "a rule is empty.";
        if (!rule.Type.Equals("Regex", StringComparison.OrdinalIgnoreCase)) return null;
        try { _ = new Regex(rule.Value, RegexOptions.IgnoreCase, TimeSpan.FromSeconds(2)); return null; }
        catch (ArgumentException e) { return $"\"{rule.Value}\" is not a valid regular expression ({e.Message})."; }
    }
}
