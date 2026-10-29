using System;
using System.IO;

// 1. Setup paths
string targetFolder = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

Environment.SetEnvironmentVariable("OPENCV_FFMPEG_LOGLEVEL", "-8"); 
Environment.SetEnvironmentVariable("OPENCV_LOG_LEVEL", "OFF");
Environment.SetEnvironmentVariable("OPENCV_VIDEOIO_DEBUG", "0");

if (!Directory.Exists(targetFolder))
{
    Console.WriteLine($"Error: The target directory '{targetFolder}' does not exist!");
    return;
}

string appFolder = AppDomain.CurrentDomain.BaseDirectory;
string settingsFile = Path.Combine(appFolder, "settings.json");

if (!File.Exists(settingsFile))
{
    Console.WriteLine($"Error: {settingsFile} not found!");
    return;
}

// 2. Initialize and Run Processor
using var processor = new FileProcessor(targetFolder, settingsFile);

Console.WriteLine($"Starting scan in: {targetFolder}");
processor.Run();

Console.WriteLine("\nProcessing Complete.");