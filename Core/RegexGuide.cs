using System.Text.RegularExpressions;

namespace FileOrganizer;

/// <summary>One pattern in the regex guide: what to type, what it means and an example that says what it finds.</summary>
public sealed record RegexTip(string Pattern, string Meaning, string Example, int UsedBy);

/// <summary>A piece of the test text: a match is shown highlighted.</summary>
public sealed record TextPiece(string Text, bool IsMatch);

/// <summary>The help shown beside a Regex rule: the patterns your rules use today, a few more that are useful, and a test box.</summary>
public static class RegexGuide
{
    // A tip, and how to see that a rule's pattern uses it (Detect is looked for in the pattern text)
    private sealed record Entry(string Pattern, string Meaning, string Example, string? Detect);

    private static readonly Entry[] Entries =
    {
        new(".?", "Any one character, or none. Good for a space, a dash or a dot that is sometimes there.",
            "night.?drive finds nightdrive, night drive and night-drive, but not night--drive.", @"\.\?"),
        new(".", "Any one character, exactly one. Good for a separator that can be a space, a dot or an underscore.",
            "real.life finds real life, real.life and real_life, but not reallife.", @"(?<!\\)\.(?![?*+{])"),
        new("^", "The start of the text: it must begin with this.",
            "^bang finds bang bus 12.mp4, but not the big bang.mp4.", @"\^"),
        new(@"\d{3,4}", @"Digits: \d is one digit, {3,4} says three or four of them ({4} is exactly four, {2,} is two or more).",
            @"AH\d{3,4} finds AH123 and AH1234, but not AH12.", @"\\d|\{\d+(,\d*)?\}"),
        new("$", "The end of the text: it must finish with this. Use ^ and $ together for an exact match.",
            @"^drops$ finds drops, but not drops 2.   \.mp4$ finds names ending in .mp4.", @"\$"),
        new(".*", "Anything in between, any number of characters (also none).",
            "premium.*party finds premium big party and premiumparty.", @"\.\*"),
        new("[ .]+", "One or more of the characters inside the brackets. Good for a separator of any length.",
            @"18[ \.]+only finds 18 only, 18.only and 18 . only.", @"\[[^\]]*\]"),
        new(@"\.", @"A real dot. A plain . means any character, so a dot in a file name is written \.",
            @"\.mp4$ finds names that end in .mp4, but not names that end in xmp4.", @"\\\."),
        new("x*", "Careful: * repeats only the character just before it, so shack* means shac, shack, shackk... To say \"anything\" write .* instead.",
            "angelo.*shack finds angelo big shack. (angelo*shack does not.)", @"[A-Za-z]\*"),

        // Not used by your rules yet
        new("(a|b)", "Either one or the other. Put the choices in brackets, with | between them.",
            "(WoW|Minecraft) finds WoW 12.mp4 and Minecraft 3.mp4.", null),
        new(@"\bword\b", @"A whole word: \b is the edge of a word, so the text is not part of a longer word.",
            @"\bart\b finds my art video, but not party.", null),
        new(@"S\d\dE\d\d", "A season and episode number: S, two digits, E, two digits.",
            "S\\d\\dE\\d\\d finds Show S01E07 and S12E10.", null),
        new(@"\d+", "One or more digits, of any length.",
            @"ep\d+ finds ep7 and ep104.", null),
        new("(19|20)\\d\\d", "A year: 19 or 20 followed by two digits.",
            @"(19|20)\d\d finds Trip 2019 and Film 1987.", null),
        new("[ ._-]?", "An optional separator: a space, a dot, an underscore or a dash, or nothing.",
            "night[ ._-]?drive finds night drive, night_drive and nightdrive.", null),
        new("^(?!.*word)", "Does NOT contain a word: the start, then a check that the word does not follow anywhere.",
            "^(?!.*trailer) finds every name without the word trailer.", null),
        new("(?-i)", "Capitals matter. Rules ignore capitals by default; put this at the start to make them count.",
            "(?-i)^AB finds AB 1.mp4, but not ab 1.mp4.", null),
    };

    /// <summary>
    /// The tips, the ones your regex rules use today first (the most used first), then the others.
    /// UsedBy says how many of your rules use each pattern (0 for the extra ones).
    /// </summary>
    public static List<RegexTip> Tips(Settings settings)
    {
        var patterns = settings.Categories.SelectMany(c => c.Rules)
            .Where(r => r.Type.Equals("Regex", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(r.Value))
            .Select(r => r.Value).ToList();
        var tips = Entries.Select(e => new RegexTip(e.Pattern, e.Meaning, e.Example,
            e.Detect == null ? 0 : patterns.Count(p => Regex.IsMatch(p, e.Detect)))).ToList();
        // Stable: used ones by use (most first), then the unused ones in the order above
        return tips.Select((t, i) => (t, i)).OrderBy(x => x.t.UsedBy > 0 ? 0 : 1).ThenByDescending(x => x.t.UsedBy).ThenBy(x => x.i).Select(x => x.t).ToList();
    }

    /// <summary>
    /// The test text cut into pieces, the parts the pattern selects marked. The text is null when the pattern is
    /// empty or broken (the message says why).
    /// </summary>
    public static (List<TextPiece>? Pieces, string Message) Highlight(string pattern, string text)
    {
        if (string.IsNullOrEmpty(pattern)) return (null, "Type a pattern in the rule to test it here.");
        Regex regex;
        try { regex = new Regex(pattern, RegexOptions.IgnoreCase, TimeSpan.FromSeconds(2)); }
        catch (ArgumentException e) { return (null, $"This is not a valid regular expression: {e.Message}"); }

        var pieces = new List<TextPiece>();
        if (string.IsNullOrEmpty(text)) return (pieces, "Type some text below to see what the pattern selects.");
        int at = 0, matches = 0;
        try
        {
            foreach (Match m in regex.Matches(text))
            {
                if (m.Length == 0) continue;   // an empty match selects nothing
                if (m.Index > at) pieces.Add(new TextPiece(text[at..m.Index], false));
                pieces.Add(new TextPiece(m.Value, true));
                at = m.Index + m.Length;
                matches++;
            }
        }
        catch (RegexMatchTimeoutException) { return (null, "This pattern takes too long on this text."); }
        if (at < text.Length) pieces.Add(new TextPiece(text[at..], false));
        if (matches == 0) pieces = new List<TextPiece> { new(text, false) };

        bool any = regex.IsMatch(text);
        string message = matches > 0 ? $"Matches {matches} {(matches == 1 ? "part" : "parts")} of the text: this rule would catch it."
            : any ? "The pattern matches, but only an empty part, so nothing is selected. The rule would catch this text."
            : "No match: this rule would not catch this text.";
        return (pieces, message);
    }
}
