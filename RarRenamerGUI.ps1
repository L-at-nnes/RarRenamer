# RAR Renamer - GUI Version
# Dark mode WPF interface for renaming RAR files

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

# XAML for dark mode GUI
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="RAR Renamer" Height="600" Width="900"
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
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="SelectionMode" Value="Extended"/>
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
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Folder selection -->
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,15">
            <Label Content="Folder:" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <TextBox x:Name="txtFolder" Width="550" VerticalAlignment="Center" IsReadOnly="True"/>
            <Button x:Name="btnBrowse" Content="Browse" Margin="10,0,0,0" Width="100"/>
        </StackPanel>
        
        <!-- Scan button -->
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,15">
            <Button x:Name="btnScan" Content="Scan Archives" Width="150" IsEnabled="False"/>
            <Label x:Name="lblStatus" Content="" Margin="15,0,0,0" Foreground="#4EC9B0"/>
        </StackPanel>
        
        <!-- Results grid -->
        <DataGrid Grid.Row="2" x:Name="dgResults" Margin="0,0,0,15">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Current Name" Binding="{Binding CurrentName}" Width="2*"/>
                <DataGridTextColumn Header="New Name" Binding="{Binding NewName}" Width="2*"/>
                <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="*"/>
            </DataGrid.Columns>
        </DataGrid>
        
        <!-- Action buttons -->
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnRenameAll" Content="Rename All" Width="150" IsEnabled="False"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Load XAML
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Get controls
$txtFolder = $window.FindName("txtFolder")
$btnBrowse = $window.FindName("btnBrowse")
$btnScan = $window.FindName("btnScan")
$lblStatus = $window.FindName("lblStatus")
$dgResults = $window.FindName("dgResults")
$btnRenameAll = $window.FindName("btnRenameAll")

# Data collection
$script:results = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
$dgResults.ItemsSource = $script:results

# Set default folder on load
$scriptFolder = Split-Path -Parent $MyInvocation.MyCommand.Path
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
            })
            continue
        }
        
        $proposedName = "$firstFolder.rar"
        
        if ($rarFile.Name -eq $proposedName) {
            $script:results.Add([PSCustomObject]@{
                CurrentName = $rarFile.Name
                NewName = $proposedName
                Status = "Already correct"
                FullPath = $rarFile.FullName
                CanRename = $false
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
            })
            continue
        }
        
        $script:results.Add([PSCustomObject]@{
            CurrentName = $rarFile.Name
            NewName = $proposedName
            Status = "Ready to rename"
            FullPath = $rarFile.FullName
            CanRename = $true
        })
        $readyCount++
    }
    
    $lblStatus.Content = "Found $($rarFiles.Count) RAR file(s) - $readyCount ready to rename"
    $lblStatus.Foreground = "#4EC9B0"
    
    if ($readyCount -gt 0) {
        $btnRenameAll.IsEnabled = $true
        
        # Auto-select ready items
        $dgResults.SelectedItems.Clear()
        foreach ($item in $script:results) {
            if ($item.CanRename) {
                $dgResults.SelectedItems.Add($item)
            }
        }
    }
})

# Rename All button click
$btnRenameAll.Add_Click({
    $itemsToRename = $dgResults.SelectedItems | Where-Object { $_.CanRename }
    
    if ($itemsToRename.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No files selected for renaming", "Info", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    $renamed = 0
    $errors = 0
    
    foreach ($item in $itemsToRename) {
        try {
            Rename-Item -Path $item.FullPath -NewName $item.NewName -ErrorAction Stop
            $item.Status = "Renamed"
            $item.CurrentName = $item.NewName
            $item.CanRename = $false
            $renamed++
        }
        catch {
            $item.Status = "Error: $_"
            $errors++
        }
    }
    
    $dgResults.Items.Refresh()
    $lblStatus.Content = "Renamed: $renamed | Errors: $errors"
    $lblStatus.Foreground = if ($errors -eq 0) { "#4EC9B0" } else { "#F48771" }
    $btnRenameAll.IsEnabled = $false
})

# Add required assembly for folder browser
Add-Type -AssemblyName System.Windows.Forms

# Show window
$window.ShowDialog() | Out-Null
