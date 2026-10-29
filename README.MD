# Video & File Organizer (Auto-Tagger & OCR)

## What it does
This program is an automated, non-destructive file sorting and renaming utility. It analyzes files in a target directory and categorizes them based on filenames, internal metadata (Title/Comments), and **Optical Character Recognition (OCR) extracted directly from video frames**. 

Instead of moving or renaming files automatically, it generates scripts: PowerShell (`.ps1`) scripts on Windows and shell (`.sh`) scripts on Linux/macOS. This allows you to review the proposed changes before executing them.

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
3. **Video OCR Deep Scan:** For video files (`.mp4`, `.mkv`, etc.), it extracts sample frames across the video and reads on-screen text (logos, website watermarks, titles) using PaddleOCR's PP-OCRv5 models running locally on ONNX Runtime ([RapidOCRSharpOnnx](https://github.com/meloht/RapidOCRSharpOnnx)). This works on Windows, Linux and macOS. The extracted text is then checked against your rules.

**Outputs Generated:**
* `move_files.ps1` / `move_files.sh`: A script to move files into categorized subfolders.
* `rename_files.ps1` / `rename_files.sh`: A script to rename files based on discovered Title/Comment metadata.
* `unmatched_files.txt`: A report of files that matched zero rules, including the text the OCR engine found, so you can update your rules.
* `_FileProcessor_VideoOcrResult.json`: A local cache to speed up future runs by remembering previously scanned video OCR data.

---

## Setup & Configuration

### 1. Prerequisites
* **OS:** Windows 10/11 or Linux x64 (e.g. Ubuntu 22.04/24.04).
* **.NET 10 SDK** (or runtime, if you use a pre-built binary).
* **At least 2 CPU cores.** On a 1-core VM or container, set the environment variable `DOTNET_PROCESSOR_COUNT=2`.
* **Linux only:** OpenCV's native library needs GTK 3, which headless servers may not have:
  ```bash
  sudo apt install libgtk-3-0t64   # Ubuntu 24.04 (use libgtk-3-0 on Ubuntu 22.04)
  ```
* **OCR models:** On the first video scan, the PP-OCRv5 models (~13 MB) are downloaded to a `models` folder next to the executable and verified with SHA-256. To run offline, download these two files yourself and put them in that folder:
  * [ch_PP-OCRv5_det_mobile.onnx](https://www.modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.10.0/onnx/PP-OCRv5/det/ch_PP-OCRv5_det_mobile.onnx) (text detection)
  * [latin_PP-OCRv5_rec_mobile.onnx](https://www.modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.10.0/onnx/PP-OCRv5/rec/latin_PP-OCRv5_rec_mobile.onnx) (text recognition, English and accented Latin letters such as æ, ø, å)

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

1. Open a terminal (Command Prompt/PowerShell on Windows, any shell on Linux).
2. Build and execute the compiled program. Or use the dotnet cli to run. Pass the target folder as an argument.

**Windows:**  
Build exe and run:  
```cmd
dotnet build  
FileOrganizer.exe "C:\Path\To\Your\Files"
```
After build the exe will be in `.\bin\Debug\net10.0`  

Alternatively, build and run using dotnet-cli:  
```cmd
dotnet run -- "C:\Path\To\Your\Files"
```

**Linux:**  
```bash
dotnet build
./bin/Debug/net10.0/FileOrganizer "/path/to/your/files"
# or
dotnet run -- "/path/to/your/files"
```

### Applying the Changes
Once the program finishes running, look in your target folder.
1. Review `unmatched_files.txt` to see what was missed and adjust your `settings.json` if needed.
2. First review and remove all non-relevant entries. Run `rename_files.ps1` (Windows) or `./rename_files.sh` (Linux) to apply metadata-based file renaming.
3. Run `move_files.ps1` (Windows) or `./move_files.sh` (Linux) to create the category folders and move your files into them.  

**Windows** (PowerShell scripts are blocked by default, so start them like this):
```powershell
powershell -ExecutionPolicy Bypass -File .\move_files.ps1
```
**Linux:**
```bash
./move_files.sh
```
The scripts never overwrite an existing file. The Windows scripts are PowerShell rather than `.bat`, because `.bat` files can't run from a network folder that isn't mapped to a drive letter (`\\nas\share\...`). They also handle file names with `%`, `[ ]`, `'`, `’` and `æøå`.