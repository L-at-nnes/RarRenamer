# RAR Renamer - GUI Version
# Dark mode WPF interface for renaming RAR files
# Features: Filters, Logging with rollback, Custom suffixes
# Automatic legacy mode for .NET 3.5 (Windows 7)

# Check .NET Framework version and determine UI mode
$script:useLegacyUI = $false
try {
    $netVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo([System.Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory() + "mscorlib.dll").ProductVersion
    
    # If .NET < 4.0, use legacy UI (no DataGrid)
    if ($netVersion -lt "4.0") {
        $script:useLegacyUI = $true
        Write-Host "Detected .NET Framework $netVersion - using legacy UI mode" -ForegroundColor Yellow
    }
    else {
        Write-Host "Detected .NET Framework $netVersion - using modern UI mode" -ForegroundColor Green
    }
}
catch {
    # If we can't check, assume modern UI and try to continue
    Write-Host "Unable to detect .NET version - trying modern UI mode" -ForegroundColor Yellow
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Find 7-Zip executable
$script:sevenZip = $null

# Check registry first
$regExePath = (Get-ItemProperty -Path "HKLM:\SOFTWARE\7-Zip" -Name "Path" -ErrorAction SilentlyContinue).Path
if ($regExePath) {
    $fullPath = Join-Path -Path $regExePath -ChildPath "7z.exe"
    if (Test-Path -Path $fullPath) {
        $script:sevenZip = $fullPath
    }
}

# Try to find in PATH
if (-not $script:sevenZip) {
    $pathSevenZip = Get-Command "7z.exe" -ErrorAction SilentlyContinue
    if ($pathSevenZip) {
        $script:sevenZip = $pathSevenZip.Source
    }
}

if (-not $script:sevenZip) {
    [System.Windows.MessageBox]::Show("7-Zip not found. Please install 7-Zip from https://www.7-zip.org/", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    exit 1
}

# Initialize log file path (Windows 7 compatible)
if ($PSScriptRoot) {
    $script:logFilePath = Join-Path -Path $PSScriptRoot -ChildPath "RarRenamer_Log.json"
}
else {
    # Fallback for Windows 7 / PowerShell 2.0
    $script:logFilePath = Join-Path -Path (Split-Path -Parent $MyInvocation.MyCommand.Path) -ChildPath "RarRenamer_Log.json"
}

# Load existing log if available
if (Test-Path -Path $script:logFilePath) {
    try {
        $script:allLogs = @(Get-Content -Path $script:logFilePath -Raw | ConvertFrom-Json)
        if (-not $script:allLogs) {
            $script:allLogs = @()
        }
    }
    catch {
        $script:allLogs = @()
    }
}
else {
    $script:allLogs = @()
}

# XAML for dark mode GUI (Modern with DataGrid - .NET 4.0+)
$xamlModern = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="RAR Renamer - Enhanced" Height="650" Width="1100"
        Background="#1E1E1E" WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#007ACC"/>
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                BorderThickness="0" 
                                CornerRadius="3"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#1C97EA"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#3E3E42"/>
                    <Setter Property="Foreground" Value="#656565"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#2D2D30"/>
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="BorderBrush" Value="#3F3F46"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="#1E1E1E"/>
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="BorderBrush" Value="#3F3F46"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="RowBackground" Value="#1E1E1E"/>
            <Setter Property="AlternatingRowBackground" Value="#1E1E1E"/>
            <Setter Property="GridLinesVisibility" Value="None"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="CanUserAddRows" Value="False"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#2D2D30"/>
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="BorderBrush" Value="#3F3F46"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
        <Style TargetType="DataGridRow">
            <Setter Property="Background" Value="#1E1E1E"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#2A2D2E"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#094771"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="DataGridCell">
                        <Border Padding="{TemplateBinding Padding}" Background="{TemplateBinding Background}">
                            <ContentPresenter VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Folder selection -->
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,15">
            <Label Content="Folder:" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <TextBox x:Name="txtFolder" Width="680" VerticalAlignment="Center" IsReadOnly="True"/>
            <Button x:Name="btnBrowse" Content="Browse" Margin="10,0,0,0" Width="100"/>
        </StackPanel>
        
        <!-- Prefix/Suffix configuration -->
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,15">
            <Label Content="Prefix:" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <TextBox x:Name="txtPrefix" Width="200" VerticalAlignment="Center" />
            <Label Content="Suffix:" VerticalAlignment="Center" Margin="30,0,10,0"/>
            <TextBox x:Name="txtSuffix" Width="200" VerticalAlignment="Center" />
            <Button x:Name="btnApplyPrefixSuffix" Content="Preview" Width="100" Margin="30,0,0,0" IsEnabled="False"/>
        </StackPanel>
        
        <!-- Scan button -->
        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,15">
            <Button x:Name="btnScan" Content="Scan Archives" Width="150" IsEnabled="False"/>
            <Label x:Name="lblStatus" Content="" Margin="15,0,0,0" Foreground="#4EC9B0"/>
        </StackPanel>
        
        <!-- Results grid -->
        <DataGrid Grid.Row="3" x:Name="dgResults" Margin="0,0,0,15">
            <DataGrid.Columns>
                <DataGridCheckBoxColumn Header="Select" Binding="{Binding IsSelected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Width="60"/>
                <DataGridTextColumn Header="Current Name" Binding="{Binding CurrentName}" Width="2*"/>
                <DataGridTextColumn Header="New Name" Binding="{Binding NewName}" Width="2*"/>
                <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="*"/>
            </DataGrid.Columns>
        </DataGrid>
        
        <!-- Action buttons -->
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnSelectAll" Content="Select All" Width="120" Margin="0,0,10,0"/>
            <Button x:Name="btnDeselectAll" Content="Deselect All" Width="120" Margin="0,0,10,0"/>
            <Button x:Name="btnRenameAll" Content="Rename Selected" Width="150" IsEnabled="False"/>
            <Button x:Name="btnUndo" Content="Undo Last Operation" Width="180" Margin="10,0,0,0" Background="#D17000"/>
        </StackPanel>
    </Grid>
</Window>
"@

# XAML for legacy mode (ListBox - .NET 3.5+)
$xamlLegacy = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="RAR Renamer - Legacy Mode" Height="650" Width="1100"
        Background="#1E1E1E" WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#007ACC"/>
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#2D2D30"/>
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="BorderBrush" Value="#3F3F46"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
        <Style TargetType="ListBox">
            <Setter Property="Background" Value="#1E1E1E"/>
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="BorderBrush" Value="#3F3F46"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Margin" Value="5,2"/>
        </Style>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Folder selection -->
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,15">
            <Label Content="Folder:" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <TextBox x:Name="txtFolder" Width="680" VerticalAlignment="Center" IsReadOnly="True"/>
            <Button x:Name="btnBrowse" Content="Browse" Margin="10,0,0,0" Width="100"/>
        </StackPanel>
        
        <!-- Prefix/Suffix configuration -->
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,15">
            <Label Content="Prefix:" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <TextBox x:Name="txtPrefix" Width="200" VerticalAlignment="Center" />
            <Label Content="Suffix:" VerticalAlignment="Center" Margin="30,0,10,0"/>
            <TextBox x:Name="txtSuffix" Width="200" VerticalAlignment="Center" />
            <Button x:Name="btnApplyPrefixSuffix" Content="Preview" Width="100" Margin="30,0,0,0" IsEnabled="False"/>
        </StackPanel>
        
        <!-- Scan button -->
        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,15">
            <Button x:Name="btnScan" Content="Scan Archives" Width="150" IsEnabled="False"/>
            <Label x:Name="lblStatus" Content="" Margin="15,0,0,0" Foreground="#4EC9B0"/>
        </StackPanel>
        
        <!-- Results list (legacy ListBox with checkboxes) -->
        <ScrollViewer Grid.Row="3" VerticalScrollBarVisibility="Auto" Margin="0,0,0,15">
            <StackPanel x:Name="spResults" Background="#1E1E1E"/>
        </ScrollViewer>
        
        <!-- Action buttons -->
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnSelectAll" Content="Select All" Width="120" Margin="0,0,10,0"/>
            <Button x:Name="btnDeselectAll" Content="Deselect All" Width="120" Margin="0,0,10,0"/>
            <Button x:Name="btnRenameAll" Content="Rename Selected" Width="150" IsEnabled="False"/>
            <Button x:Name="btnUndo" Content="Undo Last Operation" Width="180" Margin="10,0,0,0" Background="#D17000"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Choose XAML based on .NET version
$xaml = if ($script:useLegacyUI) { $xamlLegacy } else { $xamlModern }

# Load XAML (Windows 7 compatible)
$stringReader = New-Object System.IO.StringReader($xaml)
$reader = [System.Xml.XmlReader]::Create($stringReader)
$window = [System.Windows.Markup.XamlReader]::Load($reader)
$reader.Close()
$stringReader.Close()

# Get controls
$txtFolder = $window.FindName("txtFolder")
$btnBrowse = $window.FindName("btnBrowse")
$btnScan = $window.FindName("btnScan")
$lblStatus = $window.FindName("lblStatus")
$btnRenameAll = $window.FindName("btnRenameAll")
$btnUndo = $window.FindName("btnUndo")
$btnSelectAll = $window.FindName("btnSelectAll")
$btnDeselectAll = $window.FindName("btnDeselectAll")
$txtPrefix = $window.FindName("txtPrefix")
$txtSuffix = $window.FindName("txtSuffix")
$btnApplyPrefixSuffix = $window.FindName("btnApplyPrefixSuffix")

# Data collection
$script:results = New-Object System.Collections.ObjectModel.ObservableCollection[Object]

# Initialize results display based on UI mode
if ($script:useLegacyUI) {
    # Legacy mode: use StackPanel with checkboxes
    $spResults = $window.FindName("spResults")
}
else {
    # Modern mode: use DataGrid
    $dgResults = $window.FindName("dgResults")
    $dgResults.ItemsSource = $script:results
}

# Function to save log
function SaveLog {
    param (
        [array]$logEntries
    )
    
    if ($logEntries.Count -eq 0) {
        return
    }
    
    # Add to all logs
    $script:allLogs = @($script:allLogs) + @($logEntries)
    
    # Save to file
    try {
        $script:allLogs | ConvertTo-Json -Depth 10 | Set-Content -Path $script:logFilePath -Encoding UTF8
    }
    catch {
        [System.Windows.MessageBox]::Show("Error saving log: $_", "Log Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
    }
}

# Function to refresh legacy UI display
function RefreshLegacyDisplay {
    if (-not $script:useLegacyUI) { return }
    
    $spResults.Children.Clear()
    
    foreach ($item in $script:results) {
        # Create a checkbox for each item
        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.IsChecked = $item.IsSelected
        $cb.Tag = $item
        
        # Format display text
        $statusColor = switch ($item.Status) {
            "Ready to rename" { "#4EC9B0" }
            "Already correct" { "#808080" }
            "No folder found" { "#F48771" }
            "Target exists" { "#F48771" }
            "Renamed" { "#4EC9B0" }
            default { "#F48771" }
        }
        
        $displayText = "[{0}] {1} → {2} | {3}" -f `
            $(if ($item.IsSelected) { "X" } else { " " }), `
            $item.CurrentName.PadRight(40), `
            $item.NewName.PadRight(40), `
            $item.Status
        
        $cb.Content = $displayText
        $cb.Foreground = $statusColor
        $cb.IsEnabled = $item.CanRename
        
        # Handle checkbox state changes
        $cb.Add_Checked({
            $this.Tag.IsSelected = $true
        })
        $cb.Add_Unchecked({
            $this.Tag.IsSelected = $false
        })
        
        $spResults.Children.Add($cb) | Out-Null
    }
}

# Function to refresh modern UI display  
function RefreshModernDisplay {
    if ($script:useLegacyUI) { return }
    $dgResults.Items.Refresh()
}

# Set default folder on load
if ($PSScriptRoot) {
    $scriptFolder = $PSScriptRoot
}
else {
    # Fallback for Windows 7 / PowerShell 2.0
    $scriptFolder = Split-Path -Parent $MyInvocation.MyCommand.Path
}
$txtFolder.Text = $scriptFolder
$btnScan.IsEnabled = $true

# Browse button click
$btnBrowse.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select folder containing RAR files"
    $folderBrowser.SelectedPath = $txtFolder.Text
    
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtFolder.Text = $folderBrowser.SelectedPath
        $btnScan.IsEnabled = $true
        $lblStatus.Content = ""
        $script:results.Clear()
        $btnRenameAll.IsEnabled = $false
    }
})

# Scan button click
$btnScan.Add_Click({
    $scanFolder = $txtFolder.Text
    if (-not $scanFolder) { return }
    
    $script:results.Clear()
    $lblStatus.Content = "Scanning..."
    $lblStatus.Foreground = "#F1F1F1"
    
    $rarFiles = Get-ChildItem -Path $scanFolder -Filter "*.rar" -File -ErrorAction SilentlyContinue
    
    if ($rarFiles.Count -eq 0) {
        $lblStatus.Content = "No RAR files found"
        $lblStatus.Foreground = "#F48771"
        return
    }
    
    $readyCount = 0
    # Get prefix/suffix and preserve ALL spaces (including leading/trailing)
    $prefix = if ($txtPrefix.Text) { $txtPrefix.Text } else { "" }
    $suffix = if ($txtSuffix.Text) { $txtSuffix.Text } else { "" }
    
    foreach ($rarFile in $rarFiles) {
        $output = & $script:sevenZip l $rarFile.FullName
        
        $firstFolder = $null
        foreach ($line in $output) {
            if ($line -match '^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\s+D....\s+\d+\s+\d+\s+(.+)') {
                $filePath = $matches[1].Trim()
                if (-not $filePath.Contains('\') -and -not $filePath.Contains('/')) {
                    $firstFolder = $filePath
                    break
                }
            }
        }
        
        if (-not $firstFolder) {
            $script:results.Add([PSCustomObject]@{
                CurrentName = $rarFile.Name
                NewName = ""
                Status = "No folder found"
                FullPath = $rarFile.FullName
                CanRename = $false
                IsSelected = $false
            })
            continue
        }
        
        # Build proposed name with optional prefix/suffix (no automatic dash/space)
        $proposedName = $firstFolder
        if ($prefix -ne "") {
            $proposedName = "$prefix$proposedName"
        }
        if ($suffix -ne "") {
            $proposedName = "$proposedName$suffix"
        }
        $proposedName = "$proposedName.rar"
        
        if ($rarFile.Name -eq $proposedName) {
            $script:results.Add([PSCustomObject]@{
                CurrentName = $rarFile.Name
                NewName = $proposedName
                Status = "Already correct"
                FullPath = $rarFile.FullName
                CanRename = $false
                IsSelected = $false
            })
            continue
        }
        
        $targetPath = Join-Path -Path $scanFolder -ChildPath $proposedName
        
        if (Test-Path -Path $targetPath) {
            $script:results.Add([PSCustomObject]@{
                CurrentName = $rarFile.Name
                NewName = $proposedName
                Status = "Target exists"
                FullPath = $rarFile.FullName
                CanRename = $false
                IsSelected = $false
            })
            continue
        }
        
        $script:results.Add([PSCustomObject]@{
            CurrentName = $rarFile.Name
            NewName = $proposedName
            Status = "Ready to rename"
            FullPath = $rarFile.FullName
            CanRename = $true
            IsSelected = $true
        })
        $readyCount++
    }
    
    $lblStatus.Content = "Found $($rarFiles.Count) RAR file(s) - $readyCount ready to rename"
    $lblStatus.Foreground = "#4EC9B0"
    
    # Refresh display
    if ($script:useLegacyUI) {
        RefreshLegacyDisplay
    } else {
        RefreshModernDisplay
    }
    
    if ($readyCount -gt 0) {
        $btnRenameAll.IsEnabled = $true
        $btnApplyPrefixSuffix.IsEnabled = $true
    }
})

# Rename All button click
$btnRenameAll.Add_Click({
    $itemsToRename = $script:results | Where-Object { $_.IsSelected -and $_.CanRename }
    
    if ($itemsToRename.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No files selected for renaming", "Info", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    $result = [System.Windows.MessageBox]::Show("Rename $($itemsToRename.Count) selected file(s)?", "Confirm", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    
    if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
        return
    }
    
    $renamed = 0
    $errors = 0
    $sessionLog = @()
    
    foreach ($item in $itemsToRename) {
        try {
            $oldPath = $item.FullPath
            $newPath = Join-Path -Path (Split-Path -Parent $oldPath) -ChildPath $item.NewName
            
            Rename-Item -Path $oldPath -NewName $item.NewName -ErrorAction Stop
            
            # Log the operation
            $logEntry = @{
                Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                OldPath = $oldPath
                NewPath = $newPath
                OldName = $item.CurrentName
                NewName = $item.NewName
                Success = $true
            }
            $sessionLog += $logEntry
            
            $item.Status = "Renamed"
            $item.CurrentName = $item.NewName
            $item.FullPath = $newPath
            $item.CanRename = $false
            $item.IsSelected = $false
            $renamed++
        }
        catch {
            $logEntry = @{
                Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                OldPath = $item.FullPath
                NewPath = ""
                OldName = $item.CurrentName
                NewName = $item.NewName
                Success = $false
                Error = $_.Exception.Message
            }
            $sessionLog += $logEntry
            
            $item.Status = "Error: $_"
            $errors++
        }
    }
    
    # Save log
    SaveLog -logEntries $sessionLog
    
    # Refresh display
    if ($script:useLegacyUI) {
        RefreshLegacyDisplay
    } else {
        RefreshModernDisplay
    }
    
    $lblStatus.Content = "Renamed: $renamed | Errors: $errors"
    $lblStatus.Foreground = if ($errors -eq 0) { "#4EC9B0" } else { "#F48771" }
    $btnRenameAll.IsEnabled = ($script:results | Where-Object { $_.CanRename }).Count -gt 0
})

# Select All button
$btnSelectAll.Add_Click({
    foreach ($item in $script:results) {
        if ($item.CanRename) {
            $item.IsSelected = $true
        }
    }
    if ($script:useLegacyUI) {
        RefreshLegacyDisplay
    } else {
        RefreshModernDisplay
    }
})

# Deselect All button
$btnDeselectAll.Add_Click({
    foreach ($item in $script:results) {
        $item.IsSelected = $false
    }
    if ($script:useLegacyUI) {
        RefreshLegacyDisplay
    } else {
        RefreshModernDisplay
    }
})

# Apply Prefix/Suffix button click
$btnApplyPrefixSuffix.Add_Click({
    if ($script:results.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No files scanned. Please scan archives first.", "Info", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    # Get prefix/suffix and preserve ALL spaces (including leading/trailing)
    $prefix = if ($txtPrefix.Text) { $txtPrefix.Text } else { "" }
    $suffix = if ($txtSuffix.Text) { $txtSuffix.Text } else { "" }
    
    # Recalculate all proposed names
    foreach ($item in $script:results) {
        if ($item.Status -eq "No folder found") {
            continue
        }
        
        # Try to extract the original folder name by checking the archive
        $output = & $script:sevenZip l $item.FullPath
        $firstFolder = $null
        foreach ($line in $output) {
            if ($line -match '^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\s+D....\s+\d+\s+\d+\s+(.+)') {
                $filePath = $matches[1].Trim()
                if (-not $filePath.Contains('\') -and -not $filePath.Contains('/')) {
                    $firstFolder = $filePath
                    break
                }
            }
        }
        
        if (-not $firstFolder) {
            continue
        }
        
        # Build new proposed name (no automatic dash/space)
        $proposedName = $firstFolder
        if ($prefix -ne "") {
            $proposedName = "$prefix$proposedName"
        }
        if ($suffix -ne "") {
            $proposedName = "$proposedName$suffix"
        }
        $proposedName = "$proposedName.rar"
        
        # Update the item
        if ($item.CurrentName -eq $proposedName) {
            $item.NewName = $proposedName
            $item.Status = "Already correct"
            $item.CanRename = $false
            $item.IsSelected = $false
        }
        else {
            $targetPath = Join-Path -Path (Split-Path -Parent $item.FullPath) -ChildPath $proposedName
            if (Test-Path -Path $targetPath) {
                $item.NewName = $proposedName
                $item.Status = "Target exists"
                $item.CanRename = $false
                $item.IsSelected = $false
            }
            else {
                $item.NewName = $proposedName
                $item.Status = "Ready to rename"
                $item.CanRename = $true
                $item.IsSelected = $true
            }
        }
    }
    
    # Refresh the display
    if ($script:useLegacyUI) {
        RefreshLegacyDisplay
    } else {
        RefreshModernDisplay
    }
    
    # Update status and button state
    $readyCount = ($script:results | Where-Object { $_.CanRename }).Count
    $lblStatus.Content = "Prefix/Suffix applied - $readyCount file(s) ready to rename"
    $lblStatus.Foreground = "#4EC9B0"
    
    if ($readyCount -gt 0) {
        $btnRenameAll.IsEnabled = $true
    }
    else {
        $btnRenameAll.IsEnabled = $false
    }
})

# Undo button click
$btnUndo.Add_Click({
    # Check if there are logs to undo
    if ($script:allLogs.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No operations to undo", "Info", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    # Get only successful operations from the log
    $successfulLogs = $script:allLogs | Where-Object { $_.Success -eq $true }
    
    if ($successfulLogs.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No successful operations to undo", "Info", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    # Create undo selection window (modern or legacy based on UI mode)
    if ($script:useLegacyUI) {
        # Legacy mode: StackPanel with CheckBoxes
        $undoXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select Operations to Undo - Legacy Mode" Height="500" Width="800"
        Background="#1E1E1E" WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#007ACC"/>
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Margin" Value="5,2"/>
        </Style>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <Label Grid.Row="0" Content="Select operations to undo (most recent first):" Margin="0,0,0,10"/>
        
        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="0,0,0,15">
            <StackPanel x:Name="spUndoOps" Background="#1E1E1E"/>
        </ScrollViewer>
        
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnUndoSelectAll" Content="Select All" Width="120" Margin="0,0,10,0"/>
            <Button x:Name="btnUndoDeselectAll" Content="Deselect All" Width="120" Margin="0,0,10,0"/>
            <Button x:Name="btnUndoSelected" Content="Undo Selected" Width="150" Margin="0,0,10,0"/>
            <Button x:Name="btnUndoCancel" Content="Cancel" Width="100" Background="#6E6E6E"/>
        </StackPanel>
    </Grid>
</Window>
"@
    }
    else {
        # Modern mode: DataGrid
        $undoXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select Operations to Undo" Height="500" Width="800"
        Background="#1E1E1E" WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#007ACC"/>
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                BorderThickness="0" 
                                CornerRadius="3"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#1C97EA"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#3E3E42"/>
                    <Setter Property="Foreground" Value="#656565"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="#1E1E1E"/>
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="BorderBrush" Value="#3F3F46"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="RowBackground" Value="#1E1E1E"/>
            <Setter Property="AlternatingRowBackground" Value="#1E1E1E"/>
            <Setter Property="GridLinesVisibility" Value="None"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="CanUserAddRows" Value="False"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#2D2D30"/>
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="BorderBrush" Value="#3F3F46"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
        <Style TargetType="DataGridRow">
            <Setter Property="Background" Value="#1E1E1E"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#2A2D2E"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#094771"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="DataGridCell">
                        <Border Padding="{TemplateBinding Padding}" Background="{TemplateBinding Background}">
                            <ContentPresenter VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="#F1F1F1"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <Label Grid.Row="0" Content="Select operations to undo (most recent first):" Margin="0,0,0,10"/>
        
        <DataGrid Grid.Row="1" x:Name="dgUndoOps" Margin="0,0,0,15">
            <DataGrid.Columns>
                <DataGridCheckBoxColumn Header="Select" Binding="{Binding IsSelected}" Width="60"/>
                <DataGridTextColumn Header="Timestamp" Binding="{Binding Timestamp}" Width="150" IsReadOnly="True"/>
                <DataGridTextColumn Header="Old Name" Binding="{Binding OldName}" Width="*" IsReadOnly="True"/>
                <DataGridTextColumn Header="New Name" Binding="{Binding NewName}" Width="*" IsReadOnly="True"/>
            </DataGrid.Columns>
        </DataGrid>
        
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnUndoSelectAll" Content="Select All" Width="120" Margin="0,0,10,0"/>
            <Button x:Name="btnUndoDeselectAll" Content="Deselect All" Width="120" Margin="0,0,10,0"/>
            <Button x:Name="btnUndoSelected" Content="Undo Selected" Width="150" Margin="0,0,10,0"/>
            <Button x:Name="btnUndoCancel" Content="Cancel" Width="100" Background="#6E6E6E"/>
        </StackPanel>
    </Grid>
</Window>
"@
    }
    
    # Load undo window
    $undoStringReader = New-Object System.IO.StringReader($undoXaml)
    $undoReader = [System.Xml.XmlReader]::Create($undoStringReader)
    $undoWindow = [System.Windows.Markup.XamlReader]::Load($undoReader)
    $undoReader.Close()
    $undoStringReader.Close()
    
    # Get controls
    $btnUndoSelectAll = $undoWindow.FindName("btnUndoSelectAll")
    $btnUndoDeselectAll = $undoWindow.FindName("btnUndoDeselectAll")
    $btnUndoSelected = $undoWindow.FindName("btnUndoSelected")
    $btnUndoCancel = $undoWindow.FindName("btnUndoCancel")
    
    # Populate operations list (most recent first)
    $undoCollection = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
    foreach ($log in ($successfulLogs | Sort-Object -Property Timestamp -Descending)) {
        $undoCollection.Add([PSCustomObject]@{
            Timestamp = $log.Timestamp
            OldName = $log.OldName
            NewName = $log.NewName
            OldPath = $log.OldPath
            NewPath = $log.NewPath
            IsSelected = $true
        })
    }
    
    if ($script:useLegacyUI) {
        # Legacy mode: populate StackPanel with CheckBoxes
        $spUndoOps = $undoWindow.FindName("spUndoOps")
        
        foreach ($item in $undoCollection) {
            $cb = New-Object System.Windows.Controls.CheckBox
            $cb.IsChecked = $true
            $cb.Tag = $item
            
            $displayText = "[{0}] {1} → {2}" -f `
                $item.Timestamp, `
                $item.OldName.PadRight(35), `
                $item.NewName
            
            $cb.Content = $displayText
            $cb.Foreground = "#4EC9B0"
            
            $cb.Add_Checked({
                $this.Tag.IsSelected = $true
            })
            $cb.Add_Unchecked({
                $this.Tag.IsSelected = $false
            })
            
            $spUndoOps.Children.Add($cb) | Out-Null
        }
    }
    else {
        # Modern mode: bind to DataGrid
        $dgUndoOps = $undoWindow.FindName("dgUndoOps")
        $dgUndoOps.ItemsSource = $undoCollection
    }
    
    # Select All button
    $btnUndoSelectAll.Add_Click({
        foreach ($item in $undoCollection) {
            $item.IsSelected = $true
        }
        if ($script:useLegacyUI) {
            foreach ($cb in $spUndoOps.Children) {
                $cb.IsChecked = $true
            }
        } else {
            $dgUndoOps.Items.Refresh()
        }
    })
    
    # Deselect All button
    $btnUndoDeselectAll.Add_Click({
        foreach ($item in $undoCollection) {
            $item.IsSelected = $false
        }
        if ($script:useLegacyUI) {
            foreach ($cb in $spUndoOps.Children) {
                $cb.IsChecked = $false
            }
        } else {
            $dgUndoOps.Items.Refresh()
        }
    })
    
    # Cancel button
    $btnUndoCancel.Add_Click({
        $undoWindow.Close()
    })
    
    # Undo Selected button
    $btnUndoSelected.Add_Click({
        $selectedOps = $undoCollection | Where-Object { $_.IsSelected }
        
        if ($selectedOps.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No operations selected", "Info", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            return
        }
        
        $result = [System.Windows.MessageBox]::Show("Undo $($selectedOps.Count) selected operation(s)?", "Confirm Undo", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        
        if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
            return
        }
        
        $undoneCount = 0
        $undoErrors = 0
        
        foreach ($op in $selectedOps) {
            try {
                if (Test-Path -Path $op.NewPath) {
                    Rename-Item -Path $op.NewPath -NewName $op.OldName -ErrorAction Stop
                    
                    # Remove this operation from the log
                    $script:allLogs = @($script:allLogs | Where-Object { 
                        -not ($_.OldPath -eq $op.OldPath -and $_.NewPath -eq $op.NewPath -and $_.Timestamp -eq $op.Timestamp)
                    })
                    
                    # Update the item in the results if it exists
                    $item = $script:results | Where-Object { $_.FullPath -eq $op.NewPath }
                    if ($item) {
                        $item.CurrentName = $op.OldName
                        $item.FullPath = $op.OldPath
                        $item.Status = "Ready to rename"
                        $item.CanRename = $true
                        $item.IsSelected = $true
                        
                        # Recalculate the proposed new name with current prefix/suffix
                        $prefix = if ($txtPrefix.Text) { $txtPrefix.Text } else { "" }
                        $suffix = if ($txtSuffix.Text) { $txtSuffix.Text } else { "" }
                        
                        # Extract folder name from old name (remove .rar extension)
                        $folderName = [System.IO.Path]::GetFileNameWithoutExtension($op.OldName)
                        
                        # Rebuild proposed name (no automatic dash/space)
                        $proposedName = $folderName
                        if ($prefix -ne "") {
                            $proposedName = "$prefix$proposedName"
                        }
                        if ($suffix -ne "") {
                            $proposedName = "$proposedName$suffix"
                        }
                        $item.NewName = "$proposedName.rar"
                    }
                    
                    $undoneCount++
                }
                else {
                    $undoErrors++
                }
            }
            catch {
                $undoErrors++
            }
        }
        
        # Save updated log
        try {
            if ($script:allLogs.Count -gt 0) {
                $script:allLogs | ConvertTo-Json -Depth 10 | Set-Content -Path $script:logFilePath -Encoding UTF8
            }
            else {
                # Delete log file if empty
                if (Test-Path -Path $script:logFilePath) {
                    Remove-Item -Path $script:logFilePath -Force
                }
            }
        }
        catch {
            # Log save error, but undo was successful
        }
        
        # Refresh main display
        if ($script:useLegacyUI) {
            RefreshLegacyDisplay
        } else {
            RefreshModernDisplay
        }
        
        # Re-enable rename button if there are files that can be renamed
        $canRenameCount = ($script:results | Where-Object { $_.CanRename }).Count
        if ($canRenameCount -gt 0) {
            $btnRenameAll.IsEnabled = $true
        }
        
        # Show results
        if ($undoneCount -gt 0) {
            $lblStatus.Content = "Undone: $undoneCount | Errors: $undoErrors | Remaining in log: $($script:allLogs.Count)"
            $lblStatus.Foreground = if ($undoErrors -eq 0) { "#4EC9B0" } else { "#F48771" }
            
            [System.Windows.MessageBox]::Show("Successfully undone $undoneCount operation(s)", "Undo Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
        else {
            $lblStatus.Content = "Undo failed - no files were restored"
            $lblStatus.Foreground = "#F48771"
            [System.Windows.MessageBox]::Show("No files were restored. Files may have been moved or renamed manually.", "Undo Failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        }
        
        # Close undo window
        $undoWindow.Close()
    })
    
    # Show undo window
    $undoWindow.ShowDialog() | Out-Null
})

# Add required assembly for folder browser
Add-Type -AssemblyName System.Windows.Forms

# Show window
$window.ShowDialog() | Out-Null
