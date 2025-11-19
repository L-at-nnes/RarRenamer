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
- **Undo Last Operation**: Rollback the most recent rename operation
- **Undo All Operations**: Rollback ALL successful renames at once (bulk undo)
- **Persistent Log**: Works across application restarts

### ✅ Prefix & Suffix System
- **Prefix Input**: Add text before the folder name
- **Suffix Input**: Add text after the folder name
- **Direct Concatenation**: Text is added exactly as typed (no automatic dashes or spaces)
  - Use "-Portable" → `FolderName-Portable.rar`
  - Use " Portable" → `FolderName Portable.rar`
  - Use "Portable" → `FolderNamePortable.rar`
  - Use "P-" → `P-FolderName.rar`
- **Apply Button**: Recalculate all proposed names with new prefix/suffix values
- **Examples**:
  - No prefix/suffix: `MyApp.rar`
  - Prefix "P-": `P-MyApp.rar`
  - Suffix " Portable": `MyApp Portable.rar`
  - Suffix "-v2": `MyApp-v2.rar`
  - Both prefix "P-" and suffix "-x64": `P-MyApp-x64.rar`

## Requirements

- **Windows 7** or later (tested on Windows 7 & 10)
- **PowerShell 3.0+** (built-in on Windows 7+)
- **7-Zip** installed ([Download](https://www.7-zip.org/))

## Compatibility Notes

### Windows 7
- Fully compatible with Windows 7 SP1
- PowerShell 3.0+ recommended (update via Windows Management Framework)
- If you encounter errors, ensure PowerShell execution policy allows scripts:
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```

### Windows 10/11
- No additional configuration needed
- Runs out-of-the-box

## Usage

1. **Launch the script**: Run `RarRenamerGUI.ps1`
2. **Select folder**: Click "Browse" or use the default (script location)
4. **Configure prefix/suffix** (optional):
   - Enter prefix in the first text box (e.g., "P" or "MyApp ")
   - Enter suffix in the second text box (e.g., "v2" or " Portable")
   - Spaces are preserved: " Portable" will add a space before the word
5. **Scan archives**: Click "Scan Archives"
6. **Select files**: 
   - Use checkboxes to select/deselect individual files
   - Use "Select All" or "Deselect All" buttons
7. **Rename**: Click "Rename Selected"
8. **Undo if needed**: 
   - Click "Undo Last Operation" to choose between:
     - **YES**: Undo ALL recent operations
     - **NO**: Undo only the last one
     - **CANCEL**: Abort

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

**With suffix " Portable" (with leading space):**
```
Archive: test_v1.2.3.rar
Contains folder: "MyApp"
Suffix: " Portable"
Result: MyApp Portable.rar
```

**With prefix "P-" and suffix "-x64":**
```
Archive: test_v1.2.3.rar
Contains folder: "MyApp"
Prefix: "P-"
Suffix: "-x64"
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

- **Direct Concatenation**: Prefix/suffix is added exactly as typed - include your own dashes or spaces
  - Example: `-Portable` → `FolderName-Portable.rar`
  - Example: ` Portable` (with space) → `FolderName Portable.rar`
  - Example: `P-` → `P-FolderName.rar`
- **Apply Button**: After changing prefix/suffix, click "Apply" to see updated names without scanning again
- **Bulk Undo**: Use "Undo All" option to revert multiple renames at once (useful after testing)
- **Persistent Log**: The log file is saved and works across restarts - you can undo even after closing the program
- The log file grows over time - you can delete it periodically
- You can rescan with different prefixes/suffixes to preview different naming schemes
- Leave prefix and suffix empty for basic renaming (folder name only)

## Troubleshooting

### Windows 7 Issues
**Error: "Join-Path : Cannot bind argument to parameter 'Path'"**
- Update PowerShell to version 3.0 or later
- Download: [Windows Management Framework 3.0](https://www.microsoft.com/en-us/download/details.aspx?id=34595)

**Error: "Execution Policy"**
- Run PowerShell as Administrator
- Execute: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

### Common Issues
**"7-Zip not found"**
- Install 7-Zip from [7-zip.org](https://www.7-zip.org/)
- Ensure it's installed in the default location or available in PATH

**"No folder found" status**
- RAR archive doesn't contain a top-level folder
- Archive structure is flat (files only, no folders)

**Undo doesn't work**
- Check if `RarRenamer_Log.json` exists in the script folder
- Verify files haven't been manually moved/renamed after the operation
- Ensure the log file has valid JSON format

## Version History

- **v2.1** (2025-11-19): Windows 7 compatibility, space preservation in prefix/suffix, bulk undo functionality
- **v2.0** (2025-11-18): Added checkboxes, logging/rollback, and prefix/suffix system
- **v1.0**: Initial release with basic rename functionality

## License

Free to use and modify.
