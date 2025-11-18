# RAR Renamer - Enhanced Version

A PowerShell GUI tool to rename RAR archives based on their internal folder structure, with advanced features for selection, logging, and custom prefixes/suffixes.

## Features

### ✅ Checkbox Selection
- Individual checkbox for each file in the grid
- Select All / Deselect All buttons for quick selection
- Precise control over which files to rename

### ✅ Logging & Rollback System
- **Automatic Logging**: All rename operations are logged to `RarRenamer_Log.json`
- **Detailed Log Entries**: Each entry includes:
  - Timestamp
  - Old file path and name
  - New file path and name
  - Success/failure status
  - Error messages (if applicable)
- **Undo Last Operation**: Rollback the most recent batch of renames with one click
- **Persistent Log**: The log file is saved in the same directory as the script

### ✅ Prefix & Suffix System
- **Prefix Input**: Add text before the folder name (e.g., "Portable-MyApp.rar")
- **Suffix Input**: Add text after the folder name (e.g., "MyApp-v2.rar")
- **Automatic Naming**: Format is `Prefix-FolderName-Suffix.rar`
- **Examples**:
  - No prefix/suffix: `MyApp.rar`
  - Prefix "Portable": `Portable-MyApp.rar`
  - Suffix "v2": `MyApp-v2.rar`
  - Both: `Portable-MyApp-v2.rar`

## Requirements

- Windows 10/11
- PowerShell 5.1 or later
- [7-Zip](https://www.7-zip.org/) installed

## Usage

1. **Launch the script**: Run `RarRenamerGUI.ps1`
2. **Select folder**: Click "Browse" or use the default (script location)
3. **Configure prefix/suffix** (optional):
   - Enter prefix in the first text box
   - Enter suffix in the second text box
4. **Scan archives**: Click "Scan Archives"
5. **Select files**: 
   - Use checkboxes to select/deselect individual files
   - Use "Select All" or "Deselect All" buttons
6. **Rename**: Click "Rename Selected"
7. **Undo if needed**: Click "Undo Last Operation" to rollback changes

## How It Works

1. **Scans** each RAR file using 7-Zip
2. **Analyzes** the archive structure to find the first top-level folder
3. **Applies** optional prefix/suffix to the folder name
4. **Determines** if renaming is needed:
   - Folder found inside archive
   - Current name differs from proposed name
   - Target filename doesn't already exist
5. **Renames** the RAR file to match the pattern

## Examples

**Without prefix/suffix:**
```
Archive: test_v1.2.3.rar
Contains folder: "MyApp"
Result: MyApp.rar
```

**With suffix "Portable":**
```
Archive: test_v1.2.3.rar
Contains folder: "MyApp"
Suffix: "Portable"
Result: MyApp-Portable.rar
```

**With prefix "P" and suffix "x64":**
```
Archive: test_v1.2.3.rar
Contains folder: "MyApp"
Prefix: "P"
Suffix: "x64"
Result: P-MyApp-x64.rar
```

## Log File Format

The `RarRenamer_Log.json` file stores all operations:

```json
[
  {
    "Timestamp": "2025-11-18 14:30:45",
    "OldPath": "C:\\Apps\\app123.rar",
    "NewPath": "C:\\Apps\\MyApp-Portable.rar",
    "OldName": "app123.rar",
    "NewName": "MyApp-Portable.rar",
    "Success": true
  }
]
```

## Tips

- The log file grows over time - you can archive or delete it periodically
- Undo only works for the most recent batch of operations in the current session
- You can rescan with different prefixes/suffixes to preview different naming schemes
- Leave prefix and suffix empty for basic renaming (folder name only)

## Troubleshooting

- **7-Zip not found**: Install from https://www.7-zip.org/
- **Undo doesn't work**: Ensure you haven't closed the application since renaming
- **Log file error**: Check file permissions in the script directory

## Version History

- **v2.0** (2025-11-18): Added checkboxes, logging/rollback, and prefix/suffix system
- **v1.0**: Initial release with basic rename functionality

## License

Free to use and modify.
