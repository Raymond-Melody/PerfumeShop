# ============================================================
# V19 dotnet watch auto-start setup
# Replaces Windows Services with dotnet watch at login
# MUST run as Administrator
# ============================================================
param([switch]$Revert)

$ErrorActionPreference = 'Stop'

# Admin check
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host '[ERROR] Administrator required. Right-click PowerShell -> Run as Administrator.' -ForegroundColor Red
    pause; exit 1
}

$serviceApi   = 'PerfumeShopV19_API'
$serviceAdmin = 'PerfumeShopV19_Admin'
$scriptDir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$startupDir   = [Environment]::GetFolderPath('Startup')
$shortcutName = 'V19_Watch_Autostart.lnk'
$shortcutPath = Join-Path $startupDir $shortcutName
$batchPath    = Join-Path $scriptDir 'watch_v19_services.bat'

# Port check helper
function Test-PortFree([int]$Port) {
    try { $tcp = New-Object Net.Sockets.TcpClient; $tcp.Connect('127.0.0.1', $Port); $tcp.Close(); return $false }
    catch { return $true }
}

# ============================================================
# REVERT: Go back to Windows Services mode
# ============================================================
if ($Revert) {
    Write-Host '=== [REVERT] Restoring Windows Services auto-start ===' -ForegroundColor Cyan

    if (Test-Path $shortcutPath) {
        Remove-Item $shortcutPath -Force
        Write-Host '[OK] Removed startup shortcut' -ForegroundColor Green
    }

    Set-Service $serviceApi   -StartupType Automatic
    Set-Service $serviceAdmin -StartupType Automatic
    Start-Service $serviceApi, $serviceAdmin
    Write-Host '[OK] Services restored to Automatic + started' -ForegroundColor Green
    Write-Host 'Next reboot: Windows Services auto-start (NO hot reload).' -ForegroundColor Cyan
    pause; exit 0
}

# ============================================================
# SETUP: Switch to dotnet watch auto-start
# ============================================================
Write-Host '=== dotnet watch auto-start setup ===' -ForegroundColor Cyan
Write-Host 'Ports: API=5000  Admin=5207'
Write-Host ''

# Step 1: Stop services
Write-Host '[1/4] Stopping Windows Services...' -ForegroundColor Yellow
try { Stop-Service $serviceApi   -Force -ErrorAction SilentlyContinue } catch {}
try { Stop-Service $serviceAdmin -Force -ErrorAction SilentlyContinue } catch {}
Write-Host '[OK] Services stopped' -ForegroundColor Green

# Step 2: Set to Manual
Write-Host '[2/4] Setting services to Manual...' -ForegroundColor Yellow
Set-Service $serviceApi   -StartupType Manual
Set-Service $serviceAdmin -StartupType Manual
Write-Host '[OK] Services set to Manual' -ForegroundColor Green

# Step 3: Verify ports free
Write-Host '[3/4] Checking ports...' -ForegroundColor Yellow
foreach ($p in @(5000, 5207)) {
    if (Test-PortFree -Port $p) {
        Write-Host "  [OK] Port $p is free" -ForegroundColor Green
    } else {
        Write-Host "  [WARN] Port $p still in use!" -ForegroundColor Red
    }
}

# Step 4: Create startup shortcut
Write-Host '[4/4] Creating startup shortcut...' -ForegroundColor Yellow
$WshShell = New-Object -ComObject WScript.Shell
$sc = $WshShell.CreateShortcut($shortcutPath)
$sc.TargetPath       = 'cmd.exe'
$sc.Arguments        = "/c `"`"$batchPath`"`""
$sc.WorkingDirectory = $scriptDir
$sc.WindowStyle      = 7
$sc.Description      = 'V19 dotnet watch - API(5000) + Admin(5207)'
$sc.Save()
Write-Host '[OK] Shortcut created:' -ForegroundColor Green
Write-Host "     $shortcutPath" -ForegroundColor White

Write-Host ''
Write-Host '=== Done! ===' -ForegroundColor Cyan
Write-Host 'On next login: two watch windows auto-start (ports 5000/5207)'
Write-Host 'To revert to Windows Services: .\setup_watch_autostart.ps1 -Revert'
pause

