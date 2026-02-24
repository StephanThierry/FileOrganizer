# Video & File Organizer (Auto-Tagger & OCR)

## What it does
This program is an automated, non-destructive file sorting and renaming utility. It analyzes files in a target directory and categorizes them based on filenames, internal metadata (Title/Comments), and **Optical Character Recognition (OCR) extracted directly from video frames**. 

Instead of moving or renaming files automatically, it generates Windows Batch (`.bat`) files. This allows you to review the proposed changes before executing them.

## Features
- Multi-stage matching priority:  
  1. Filename  
  2. Metadata title/comment  
  3. OCR text from selected video frames (deep scan)  
- Detects likely corrupt/unreadable videos (Put into category `_CHECK_`)  

## How It Works
The processor evaluates every file in the target folder using a 3-step deep scan:
1. **Filename Matching:** Checks if the file's name matches your predefined text or Regex rules.
2. **Metadata Extraction:** Reads media tags (Title, Comments) using TagLib. If tags exist, it proposes a file rename based on the tags and checks the tags against your rules.
3. **Video OCR Deep Scan:** For video files (`.mp4`, `.mkv`, etc.), it extracts sample frames across the video, applies image processing (bilateral filtering, CLAHE) to clean up noise, and reads on-screen text using native Windows OCR. The extracted text is then checked against your rules.

**Outputs Generated:**
* `move_files.bat`: A script to move files into categorized subfolders.
* `rename_files.bat`: A script to rename files based on discovered Title/Comment metadata.
* `unmatched_files.txt`: A report of files that matched zero rules, including the text the OCR engine found, so you can update your rules.
* `_FileProcessor_VideoOcrResult.json`: A local cache to speed up future runs by remembering previously scanned video OCR data.

---

## Setup & Configuration

### 1. Prerequisites
* **OS:** Windows 10 or Windows 11 (required for the native `Windows.Media.Ocr` engine).

### 2. Creating `settings.json`
The program requires a `settings.json` file located in the same directory as the application executable. This file defines your categories and the rules used to match files.

Create a `settings.json` file using the following structure:

The project includes: `settings.json-template` that demonstrate the format:  

**RegEx used in this example:**    
* `^` → Start of the string
* `.` → Any 1 character
* `.?` → Empty or 1 characters
* `.*` → Any number of characters
* `\\.` → Literal dot (.) - the regex is `\.` but must be escaped inside the json
* `\\d{4}` → any 4 digits 
* `$` → End of string

Combined example:
* `\\d{4}p.*\\.mp4$` → Would match: "Something.1080p.HD.mp4" or "Any title.2160p.BluRay.mp4"

This rule is last so it works as a catch-all of  HD/4K videos not being mapped by any previous rules.  

```json
{
  "Categories": [
    {
      "Name": "Tutorials",
      "IgnoreCategoryInContent": false,
      "Rules": [
        { "Type": "Contains", "Value": "Tutorial" },
        { "Type": "Contains", "Value": "Course" },
        { "Type": "Contains", "Value": "Lesson" }
      ]
    },
    {
      "Name": "Gameplay_Recordings",
      "IgnoreCategoryInContent": false,
      "Rules": [
        { "Type": "Regex", "Value": "^(WoW|Minecraft|Halo).*\\.mp4$" }
      ]
    },
    {
      "Name": "Last Week Tonight With John Oliver",
      "IgnoreCategoryInContent": false,
      "Rules": [
        { "Type": "Regex", "Value": "Last.?Week.?Tonight" }
      ]
    },
    {
      "Name": "Barry",
      "IgnoreCategoryInContent": true,
      "Rules": [
        { "Type": "Regex", "Value": "^Barry.*\\.mp4$" }
      ]
    },
    {
      "Name": "American Dad",
      "IgnoreCategoryInContent": false,
      "Rules": [ 
        { "Type": "Regex", "Value": "American.?Dad" }
      ]
    },
    {
      "Name": "HD and 4K",
      "IgnoreCategoryInContent": false,
      "Rules": [ 
        { "Type": "Regex", "Value": "\\d{4}p.*\\.mp4$" }
      ]
    }
  ]
}
```

**Settings Breakdown:**
* **`Name`**: The name of the category (this will become the destination folder name).
* **`IgnoreCategoryInContent`**: If `false`, the program will automatically trigger a match if the word in `Name` is found anywhere in the file/OCR.   
Use `true` if the category name is a general word (like `Barry`) that should not trigger a match. Use the RegEx rules to trigger the matches.
* **`Rules`**: A list of rules. 
  * `Type` can be either `"Contains"` (simple text match) or `"Regex"` (Regular Expression match).
  * `Value` is the text or regex pattern to look for.

---

## How to Run

1. Open a Command Prompt or PowerShell window.
2. Build and execute the compiled program. Or use the dotnet cli to run. Pass the target folder as an argument.

**To scan a specific folder:**
Build exe and run:  
```cmd
dotnet build  
FileOrganizer.exe "C:\Path\To\Your\Files"
```
After build the exe will be in `.\bin\Debug\net10.0-windows10.0.19041.0`  

Alternatevely, build and run using donet-cli:  
```cmd
dotnet run -- "C:\Path\To\Your\Files"
```

### Applying the Changes
Once the program finishes running, look in your target folder.
1. Review `unmatched_files.txt` to see what was missed and adjust your `settings.json` if needed.
2. First review and remove all non-relevat entries. Run `rename_files.bat` to apply metadata-based file renaming.
3. Run `move_files.bat` to create the category folders and move your files into them.  

```cmd
move_files.bat
```