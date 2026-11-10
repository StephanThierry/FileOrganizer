namespace FileOrganizer;

/// <summary>A file that is still in a folder that was opened ("Process folders"), with its path inside that folder.</summary>
public sealed record LeftoverFile(string Path, long Bytes);

/// <summary>What is left in one folder that was opened: the files (any depth), and the sub-folders that hold no file at all.</summary>
public sealed record FolderLeftover(string Name, bool Exists, IReadOnlyList<LeftoverFile> Files, IReadOnlyList<string> EmptyFolders)
{
    public long Bytes => Files.Sum(f => f.Bytes);

    /// <summary>The folder is there and no file is left in it (it could be deleted).</summary>
    public bool IsEmpty => Exists && Files.Count == 0;

    /// <summary>One line for the folder, with the first <paramref name="show"/> files that are left.</summary>
    public string Line(int show = 6)
    {
        if (!Exists) return $"{Name}: gone (it was deleted).";
        string empties = EmptyFolders.Count == 0 ? "" : $" {Fmt.Plural(EmptyFolders.Count, "empty sub-folder", "empty sub-folders")}.";
        if (Files.Count == 0) return $"{Name}: nothing left, the folder can be deleted.{empties}";
        string list = string.Join(", ", Files.Take(show).Select(f => $"{f.Path} ({Fmt.Size(f.Bytes)})"));
        string more = Files.Count > show ? $" and {Files.Count - show:N0} more" : "";
        return $"{Name}: {Fmt.Plural(Files.Count, "file", "files")} left ({Fmt.Size(Bytes)}): {list}{more}.{empties}";
    }
}

/// <summary>What is left, after the moves, in the folders that were opened because their name fits a category.</summary>
public sealed record LeftoverReport(string Root, IReadOnlyList<FolderLeftover> Folders)
{
    public int EmptyCount => Folders.Count(f => f.IsEmpty);
    public int GoneCount => Folders.Count(f => !f.Exists);
    public int FilesLeft => Folders.Sum(f => f.Files.Count);
    public long BytesLeft => Folders.Sum(f => f.Bytes);

    /// <summary>The whole in one sentence.</summary>
    public string Headline
    {
        get
        {
            var parts = new List<string>();
            int withFiles = Folders.Count(f => f.Exists && f.Files.Count > 0);
            if (EmptyCount > 0) parts.Add($"{Fmt.Plural(EmptyCount, "folder is", "folders are")} empty");
            if (withFiles > 0) parts.Add($"{Fmt.Plural(FilesLeft, "file", "files")} ({Fmt.Size(BytesLeft)}) left in {Fmt.Plural(withFiles, "folder", "folders")}");
            if (GoneCount > 0) parts.Add($"{Fmt.Plural(GoneCount, "folder was", "folders were")} deleted");
            string opened = $"Of the {Fmt.Plural(Folders.Count, "folder", "folders")} that {(Folders.Count == 1 ? "was" : "were")} opened: ";
            return opened + (parts.Count == 0 ? "nothing is left." : string.Join("; ", parts) + ".");
        }
    }

    /// <summary>A line for each folder: the folders with files left first, then the empty ones.</summary>
    public IReadOnlyList<string> Lines(int show = 6) =>
        Folders.OrderBy(f => f.Exists && f.Files.Count > 0 ? 0 : f.Exists ? 1 : 2).ThenBy(f => f.Name, StringComparer.OrdinalIgnoreCase).Select(f => f.Line(show)).ToList();
}

public static class FolderLeftovers
{
    /// <summary>Looks in the folders (names inside <paramref name="root"/>) and lists what is still there. Only reads.</summary>
    public static LeftoverReport Build(string root, IEnumerable<string> folderNames)
    {
        var options = new EnumerationOptions { RecurseSubdirectories = true, IgnoreInaccessible = true, AttributesToSkip = 0 };
        var folders = new List<FolderLeftover>();
        foreach (var name in folderNames.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            string dir = Path.Combine(root, name);
            if (!Directory.Exists(dir)) { folders.Add(new FolderLeftover(name, false, Array.Empty<LeftoverFile>(), Array.Empty<string>())); continue; }
            var files = new List<LeftoverFile>();
            var empty = new List<string>();
            try
            {
                foreach (var file in Directory.EnumerateFiles(dir, "*", options).OrderBy(f => f, StringComparer.OrdinalIgnoreCase))
                {
                    long length = 0;
                    try { length = new FileInfo(file).Length; } catch (Exception) { }
                    files.Add(new LeftoverFile(Path.GetRelativePath(dir, file), length));
                }
                foreach (var sub in Directory.EnumerateDirectories(dir, "*", options).OrderBy(d => d, StringComparer.OrdinalIgnoreCase))
                    if (!Directory.EnumerateFiles(sub, "*", options).Any()) empty.Add(Path.GetRelativePath(dir, sub));
            }
            catch (Exception) { /* unreadable: what was seen is shown */ }
            folders.Add(new FolderLeftover(name, true, files, empty));
        }
        return new LeftoverReport(root, folders);
    }
}
