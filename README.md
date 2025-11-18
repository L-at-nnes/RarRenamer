# RAR Renamer

A simple Windows tool to rename RAR archives based on the first folder name inside them.

## Features

- **Dark Mode GUI** - Modern VS Code-inspired dark theme interface
- **Automatic Detection** - Scans RAR files and detects the first folder inside
- **Batch Renaming** - Processes multiple RAR files at once
- **Smart Selection** - Automatically selects files ready to be renamed
- **7-Zip Integration** - Works with 7-Zip installed in standard or custom locations
- **Clean Interface** - Simple and intuitive workflow

## Requirements

- Windows 7 or later
- PowerShell 5.1 or later (included in Windows)
- [7-Zip](https://www.7-zip.org/) installed

## Installation

1. Download or clone this repository
2. Ensure 7-Zip is installed (the tool will detect it automatically)
3. No additional installation required

## Usage

Simply run `RarRenamerGUI.ps1`:

```powershell
.\RarRenamerGUI.ps1
```

**Steps:**
1. Click **Browse** to select a folder containing RAR files (defaults to script folder)
2. Click **Scan Archives** to analyze the archives
3. Review the proposed names in the grid
4. Select the files you want to rename (or keep the auto-selection)
5. Click **Rename All** to apply changes

**Status Indicators:**
- **Ready to rename** - File can be renamed safely
- **Already correct** - File already has the correct name
- **No folder found** - Archive doesn't contain a top-level folder
- **Target exists** - A file with the target name already exists

## How It Works

1. **Scans** each RAR file using 7-Zip
2. **Analyzes** the archive structure to find the first top-level folder
3. **Determines** if renaming is needed:
   - Folder found inside archive
   - Current name differs from folder name
   - Target filename doesn't already exist
4. **Renames** the RAR file to match the internal folder name

## Examples

**Before:**
```
Archive-test-xyz-1.2.3.4.rar  →  Contains folder: "Example"
```

**After:**
```
Example.rar
```

## Troubleshooting

**7-Zip not found**
- Install 7-Zip from [7-zip.org](https://www.7-zip.org)
- The tool automatically detects 7-Zip via Windows Registry or system PATH

**PowerShell execution policy error**
```powershell
# Run with bypass policy
powershell -ExecutionPolicy Bypass -File .\RarRenamerGUI.ps1
```

**Files not renaming**
- Ensure you have write permissions in the target folder
- Check that proposed filenames don't already exist
- Verify RAR files contain at least one top-level folder

**GUI doesn't appear**
- Make sure PowerShell 5.1+ is installed
- Check Windows version (Windows 7 or later required)

## Project Files

- **RarRenamerGUI.ps1** - Graphical WPF interface (dark theme)
- **README.md** - This documentation file

## Technical Details

- **Language**: PowerShell 5.1+
- **UI Framework**: WPF (Windows Presentation Foundation)
- **Archive Tool**: 7-Zip command-line interface
- **Color Scheme**: VS Code-inspired dark theme (#1E1E1E background, #F1F1F1 text, #007ACC accents, #4EC9B0 success)

## License

Free to use and modify for personal and commercial purposes.
