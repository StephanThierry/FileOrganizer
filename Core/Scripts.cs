using System.Text;
using System.Text.RegularExpressions;

namespace FileOrganizer;

/// <summary>
/// Writes the move and rename scripts: PowerShell (.ps1) on Windows, shell (.sh) on Linux and macOS. Not .bat: cmd.exe can't use an
/// unmapped UNC path like \\nas\share as its current folder. Nothing is ever overwritten (File.Move and mv -n leave an existing file).
/// </summary>
public static class ScriptWriter
{
    private static readonly bool IsWindows = OperatingSystem.IsWindows();
    public static string ScriptExtension => IsWindows ? ".ps1" : ".sh";

    public static string MovePath(string folder) => Path.Combine(folder, "move_files" + ScriptExtension);
    public static string RenamePath(string folder) => Path.Combine(folder, "rename_files" + ScriptExtension);
    public static string UnmatchedPath(string folder) => Path.Combine(folder, "unmatched_files.txt");

    /// <summary>move_files: a folder per category, and a move for each file that has one (the user's choice counts). Null when nothing moves.</summary>
    public static string? WriteMoveScript(string folder, IEnumerable<FileDecision> decisions)
    {
        var moves = decisions.Where(d => d.EffectiveCategory != null).ToList();
        string path = MovePath(folder);
        if (moves.Count == 0) { TryDelete(path); return null; }

        using var writer = CreateScript(path);
        foreach (var category in moves.Select(d => d.EffectiveCategory!).Distinct())
        {
            writer.WriteLine(IsWindows
                ? $"[System.IO.Directory]::CreateDirectory((P {PsQuote(category)})) | Out-Null"
                : $"mkdir -p -- {ShQuote(category)}");
        }
        writer.WriteLine();
        writer.WriteLine("# --- Moving Files ---");
        foreach (var d in moves)
        {
            string category = d.EffectiveCategory!;
            writer.WriteLine(IsWindows
                ? $"[System.IO.File]::Move((P {PsQuote(d.FileName)}), (P {PsQuote(category)} {PsQuote(d.FileName)}))"
                : $"mv -n -- {ShQuote(d.FileName)} {ShQuote(category + "/")}");
            string why = d.OverrideCategory != null ? "Chosen by you" : d.Trigger;
            writer.WriteLine($"# Triggered: {why.Replace('\r', ' ').Replace('\n', ' ')}");
        }
        MakeExecutable(path);
        return path;
    }

    /// <summary>rename_files: a new name from the title or comment tag, for the files that have one. Null when there is none.</summary>
    public static string? WriteRenameScript(string folder, IEnumerable<FileDecision> decisions)
    {
        var renames = decisions.Where(d => d.RenameTo != null).ToList();
        string path = RenamePath(folder);
        if (renames.Count == 0) { TryDelete(path); return null; }

        using var writer = CreateScript(path);
        writer.WriteLine("# === Auto-generated Rename Script (Title/Comment) ===");
        writer.WriteLine();
        foreach (var d in renames)
        {
            writer.WriteLine(IsWindows
                ? $"[System.IO.File]::Move((P {PsQuote(d.FileName)}), (P {PsQuote(d.RenameTo!)}))"
                : $"mv -n -- {ShQuote(d.FileName)} {ShQuote(d.RenameTo!)}");
            writer.WriteLine($"# Source: {d.RenameSource}");
            writer.WriteLine();
        }
        MakeExecutable(path);
        return path;
    }

    /// <summary>unmatched_files.txt: the files nothing matched, with the tags and the text that was found, to help writing rules.</summary>
    public static string? WriteUnmatchedReport(string folder, IEnumerable<FileDecision> decisions)
    {
        var unmatched = decisions.Where(d => d.Category == null).ToList();
        string path = UnmatchedPath(folder);
        if (unmatched.Count == 0) { TryDelete(path); return null; }

        using var writer = new StreamWriter(path);
        foreach (var d in unmatched)
        {
            writer.WriteLine(d.FileName + "...");
            if (!string.IsNullOrWhiteSpace(d.Title)) writer.WriteLine($"    Metadata Title: \"{d.Title}\"");
            if (!string.IsNullOrWhiteSpace(d.Comment)) writer.WriteLine($"    Metadata Comment: \"{d.Comment}\"");
            foreach (var line in d.OcrLines) writer.WriteLine($"    OCR Detected Text: \"{line}\"");
            writer.WriteLine();
        }
        return path;
    }

    private static void TryDelete(string path)
    {
        // An older script would otherwise be run by mistake instead of "nothing to do"
        try { if (File.Exists(path)) File.Delete(path); } catch { }
    }

    private static StreamWriter CreateScript(string path)
    {
        // .ps1: UTF-8 with BOM, or Windows PowerShell 5.1 reads æ/ø/å and ' - as ANSI. .sh: no BOM, LF.
        var writer = new StreamWriter(path, false, new UTF8Encoding(IsWindows)) { NewLine = IsWindows ? "\r\n" : "\n" };
        if (IsWindows)
        {
            writer.WriteLine($"# Run with:  powershell -ExecutionPolicy Bypass -File \"{Path.GetFileName(path)}\"");
            writer.WriteLine("$ErrorActionPreference = 'Continue'");
            writer.WriteLine("# Paths are relative to this script's folder - works for UNC paths (\\\\server\\share) too");
            writer.WriteLine("$here = $PSScriptRoot");
            writer.WriteLine("function P([string]$a, [string]$b = '') { if ($b) { [System.IO.Path]::Combine($here, $a, $b) } else { [System.IO.Path]::Combine($here, $a) } }");
        }
        else
        {
            writer.WriteLine("#!/bin/sh");
            writer.WriteLine("cd \"$(dirname \"$0\")\" || exit 1");
        }
        writer.WriteLine();
        return writer;
    }

    // Single quotes are literal in PowerShell ($ ` [ ] % are not special). PowerShell also treats the typographic quotes
    // as single quotes, so each of them is doubled like ' to stay literal.
    public static string PsQuote(string s) => "'" + Regex.Replace(s, "['\u2018\u2019\u201A\u201B]", m => m.Value + m.Value) + "'";

    // Single quotes keep everything literal in sh; a ' inside is written as '\''
    public static string ShQuote(string s) => "'" + s.Replace("'", "'\\''") + "'";

    private static void MakeExecutable(string path)
    {
        if (OperatingSystem.IsWindows()) return;
        try
        {
            File.SetUnixFileMode(path, File.GetUnixFileMode(path) | UnixFileMode.UserExecute | UnixFileMode.GroupExecute | UnixFileMode.OtherExecute);
        }
        catch (Exception)
        {
            // Not fatal: run it with sh
        }
    }
}

/// <summary>What happened when the moves were applied.</summary>
public sealed record MoveReport(int Moved, int Renamed, int SkippedExisting, int Failed, List<string> Messages)
{
    public string Summary
    {
        get
        {
            var parts = new List<string>();
            if (Moved > 0) parts.Add($"moved {Fmt.Plural(Moved, "file", "files")}");
            if (Renamed > 0) parts.Add($"renamed {Fmt.Plural(Renamed, "file", "files")}");
            if (SkippedExisting > 0) parts.Add($"left {Fmt.Plural(SkippedExisting, "file", "files")} because the target exists already");
            if (Failed > 0) parts.Add($"{Fmt.Plural(Failed, "file", "files")} failed");
            return parts.Count == 0 ? "Nothing to do." : char.ToUpper(parts[0][0]) + parts[0][1..] + (parts.Count > 1 ? ", " + string.Join(", ", parts.Skip(1)) : "") + ".";
        }
    }
}

/// <summary>Moves (and renames) the files now, instead of writing a script. Never overwrites a file.</summary>
public static class MoveExecutor
{
    /// <param name="alsoRename">Files that have a new name from their tags get it (in the category folder, or in place when they stay).</param>
    public static MoveReport Apply(string folder, IEnumerable<FileDecision> decisions, bool alsoRename, IProgress<int>? progress = null)
    {
        int moved = 0, renamed = 0, skipped = 0, failed = 0, n = 0;
        var messages = new List<string>();
        foreach (var d in decisions.ToList())
        {
            progress?.Report(++n);
            if (d.Done) continue;
            string? category = d.EffectiveCategory;
            string newName = alsoRename && d.RenameTo != null ? d.RenameTo : d.FileName;
            bool renames = !string.Equals(newName, d.FileName, StringComparison.Ordinal);
            if (category == null && !renames) continue;

            string source = Path.Combine(folder, d.FileName);
            try
            {
                if (!File.Exists(source)) { failed++; messages.Add($"{d.FileName}: the file is not there any more."); continue; }
                string targetFolder = folder;
                if (category != null)
                {
                    // A category is a folder inside the scanned folder: never a path that leads out of it
                    if (!FolderNames.IsSafe(category)) { failed++; messages.Add($"{d.FileName}: \"{category}\" is not a usable folder name."); continue; }
                    targetFolder = Path.Combine(folder, category);
                    Directory.CreateDirectory(targetFolder);
                }
                if (!FolderNames.IsSafe(newName)) { failed++; messages.Add($"{d.FileName}: \"{newName}\" is not a usable file name."); continue; }
                string target = Path.Combine(targetFolder, newName);
                if (File.Exists(target)) { skipped++; messages.Add($"{d.FileName}: {Path.Combine(category ?? "", newName)} exists already, so it stays."); continue; }

                File.Move(source, target);
                d.Done = true;
                if (category != null) moved++; else renamed++;
                if (category != null && renames) renamed++;
            }
            catch (Exception e)
            {
                failed++;
                messages.Add($"{d.FileName}: {e.Message}");
            }
        }
        return new MoveReport(moved, renamed, skipped, failed, messages);
    }
}
