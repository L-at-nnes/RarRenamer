# .NET Framework 4.5 Installer for Windows 7
# Auto-download and install if not present

Write-Host "Checking .NET Framework version..." -ForegroundColor Cyan

# Check if .NET 4.0+ is installed
$dotNetVersion = $null
try {
    $dotNetVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo([System.Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory() + "mscorlib.dll").ProductVersion
    Write-Host "Current .NET Framework version: $dotNetVersion" -ForegroundColor Yellow
    
    if ($dotNetVersion -ge "4.0") {
        Write-Host ".NET Framework 4.0+ is already installed!" -ForegroundColor Green
        Write-Host "Press any key to exit..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 0
    }
}
catch {
    Write-Host "Unable to detect .NET Framework version" -ForegroundColor Red
}

Write-Host "`n.NET Framework 4.5 is required but not installed." -ForegroundColor Yellow
Write-Host "This script will download and install it automatically.`n" -ForegroundColor Yellow

# Ask for confirmation
$response = Read-Host "Do you want to download and install .NET Framework 4.5? (Y/N)"
if ($response -ne "Y" -and $response -ne "y") {
    Write-Host "Installation cancelled." -ForegroundColor Red
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

Write-Host "`nOpening download page in your browser..." -ForegroundColor Cyan
Write-Host "Please download and install .NET Framework 4.5.2 manually." -ForegroundColor Yellow
Write-Host "`nSteps:" -ForegroundColor White
Write-Host "1. Click 'Download' on the page that will open" -ForegroundColor White
Write-Host "2. Run the downloaded file (NDP452-KB2901907-x86-x64-AllOS-ENU.exe)" -ForegroundColor White
Write-Host "3. Follow the installation wizard" -ForegroundColor White
Write-Host "4. Reboot your computer when prompted" -ForegroundColor White
Write-Host "5. Run RarRenamerGUI.ps1 again`n" -ForegroundColor White

# Open browser to download page
Start-Process "https://www.microsoft.com/en-us/download/details.aspx?id=42643"

Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
