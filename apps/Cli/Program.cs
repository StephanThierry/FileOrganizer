using FileOrganizer;

// Usage: FileOrganizer [folder] [--data <folder with settings.json>]
//   folder  the folder with the files to organize (default: the current folder)
//   --data  where settings.json is (or FILEORG_DATA). Default: the current folder or the program's folder if they
//           hold a settings.json, otherwise the user's application data folder.
//
// Looks at every file in the folder and writes move_files / rename_files scripts (.ps1 on Windows, .sh elsewhere) and
// unmatched_files.txt. Nothing is moved. The desktop and web apps show the same result and can also move the files.

string? targetArg = null;
for (int i = 0; i < args.Length; i++)
{
    if (args[i] == "--data") { i++; continue; }
    targetArg ??= args[i];
}
string targetFolder = targetArg ?? Directory.GetCurrentDirectory();
string dataFolder = DataLocation.ResolveDataFolder(args, null) ?? DataLocation.DefaultDataFolder();

if (!Directory.Exists(targetFolder))
{
    Console.WriteLine($"Error: The target directory '{targetFolder}' does not exist!");
    return 1;
}
if (!SettingsStore.Exists(dataFolder))
{
    Console.WriteLine($"Error: {SettingsStore.PathIn(dataFolder)} not found! (Copy settings.json-template there, or use the desktop app to make one.)");
    return 1;
}

Settings settings;
try { settings = SettingsStore.Load(dataFolder); }
catch (InvalidDataException e) { Console.WriteLine("Error: " + e.Message); return 1; }
foreach (var problem in SettingsStore.Validate(settings)) Console.WriteLine("Warning: " + problem);

using var cts = new CancellationTokenSource();
Console.CancelKeyPress += (_, e) => { e.Cancel = true; cts.Cancel(); Console.WriteLine("\nStopping..."); };

using var engine = new OrganizeEngine(() => new OpenCvVideoOcr(DataLocation.ModelsFolder(dataFolder)));
engine.Log += Console.WriteLine;

Console.WriteLine($"Starting scan in: {targetFolder}");
var result = engine.Run(targetFolder, settings, cts.Token);

var move = ScriptWriter.WriteMoveScript(targetFolder, result.Decisions);
var rename = ScriptWriter.WriteRenameScript(targetFolder, result.Decisions);
ScriptWriter.WriteUnmatchedReport(targetFolder, result.Decisions);
if (move != null) Console.WriteLine($"Generated '{Path.GetFileName(move)}' with {result.Matched.Count()} moves.");
if (rename != null) Console.WriteLine($"Generated '{Path.GetFileName(rename)}' with {result.Renames.Count()} commands.");

Console.WriteLine("\nProcessing Complete.");
return 0;
