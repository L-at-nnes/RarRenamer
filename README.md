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
- **Selective Undo**: Choose which operations to undo with checkboxes
  - Click "Undo Last Operation" to open selection window
  - Select specific operations or use "Select All"/"Deselect All"
  - Click "Undo Selected" to revert chosen operations
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

- **Windows 7** or later (tested on Windows 7 & 11)
- **PowerShell 3.0+** (built-in on Windows 7+)
- **.NET Framework**:
  - **.NET 4.0+** for modern UI with DataGrid (Windows 10/11 default)
  - **.NET 3.5+** for legacy UI mode (Windows 7 default)
  - **Automatic UI selection**: The script detects your .NET version and chooses the appropriate interface
- **7-Zip** installed ([Download](https://www.7-zip.org/))

## Compatibility Notes

### Windows 7

- **Fully compatible with Windows 7 SP1**
- **Automatic Legacy Mode**: 
  - If .NET Framework 3.5 is detected (default on Windows 7), the script automatically uses a simplified interface
  - Legacy mode uses checkboxes instead of DataGrid for full .NET 3.5 compatibility
  - All features work identically - only the visual presentation differs
  - **No installation required** - works out of the box with Windows 7's default .NET Framework
- **Optional Modern Mode**:
  - Install .NET Framework 4.5+ to unlock the modern DataGrid interface
  - Download from [Microsoft](https://www.microsoft.com/en-us/download/details.aspx?id=42643)
  - Requires a system reboot after installation
- PowerShell 3.0+ recommended (update via Windows Management Framework)
- If you encounter errors, ensure PowerShell execution policy allows scripts:
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```

### Windows 10/11

- **No configuration needed** - runs with modern UI mode automatically
- Uses .NET Framework 4.x (pre-installed)
- Full DataGrid interface with all features

## Usage

1. **Launch the script**: Run `RarRenamerGUI.ps1`
2. **Select folder**: Click "Browse" or use the default (script location)
3. **Scan archives**: Click "Scan Archives"
4. **Configure prefix/suffix** (optional):
   - Enter prefix in the first text box (e.g., "P-" or "MyApp ")
   - Enter suffix in the second text box (e.g., "-v2" or " Portable")
   - Text is concatenated exactly as typed - add your own dashes/spaces
   - Click **"Preview"** to recalculate all proposed names
5. **Select files**: 
   - Use checkboxes to select/deselect individual files
   - Use "Select All" or "Deselect All" buttons
6. **Rename**: Click "Rename Selected"
7. **Undo if needed**: 
   - Click "Undo Last Operation" to open the undo window
   - Check/uncheck specific operations you want to undo
   - Use "Select All" or "Deselect All" for bulk selection
   - Click "Undo Selected" to revert chosen operations

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
- **Preview Button**: After changing prefix/suffix, click "Preview" to see updated names without scanning again
- **Selective Undo**: Open undo window to choose exactly which operations to revert
- **Persistent Log**: The log file is saved and works across restarts - you can undo even after closing the program
- The log file grows over time - you can delete it periodically
- You can rescan with different prefixes/suffixes to preview different naming schemes
- Leave prefix and suffix empty for basic renaming (folder name only)

## Troubleshooting

### Windows 7 Issues

**Script works but interface looks different**
- This is normal! The script automatically detects .NET 3.5 and uses legacy mode
- Legacy mode = checkboxes in a list instead of DataGrid table
- All features work exactly the same - only visual presentation differs
- To get the modern DataGrid interface, install .NET Framework 4.5+ (optional)

**Want to upgrade to modern UI?**
- Install .NET Framework 4.5+ manually: [Download here](https://www.microsoft.com/en-us/download/details.aspx?id=42643)
- **Reboot required** after installation
- Script will automatically detect and switch to modern UI mode

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

- **v2.2** (2025-11-24): Automatic UI mode detection - legacy mode for .NET 3.5 (Windows 7), modern mode for .NET 4.0+
- **v2.1** (2025-11-19): Windows 7 compatibility, space preservation in prefix/suffix, bulk undo functionality
- **v2.0** (2025-11-18): Added checkboxes, logging/rollback, and prefix/suffix system
- **v1.0**: Initial release with basic rename functionality

## License

Free to use and modify.
