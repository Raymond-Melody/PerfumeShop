param(
    [ValidateSet('Api', 'Admin', 'Both')]
    [string]$Target = 'Both'
)

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function Test-PortInUse {
    param([int]$Port)
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $tcp.Connect('127.0.0.1', $Port)
        $tcp.Close()
        return $true
    } catch {
        return $false
    }
}

function Start-Watch {
    param(
        [string]$Name,
        [string]$ProjectDir,
        [int]$Port
    )

    if (Test-PortInUse -Port $Port) {
        Write-Host "[SKIP] $Name port $Port already in use"
        return
    }

    $url = "http://localhost:$Port"
    Write-Host "[START] $Name -> $url"

    $block = @"
Set-Location -LiteralPath '$ProjectDir'
`$env:ASPNETCORE_ENVIRONMENT = 'Development'
`$env:ASPNETCORE_URLS = '$url'
dotnet watch --non-interactive --no-launch-profile
"@

    Start-Process -FilePath 'powershell.exe' `
        -ArgumentList '-NoExit', '-NoProfile', '-Command', $block `
        -WindowStyle Normal
}

if ($Target -in @('Api', 'Both')) {
    Start-Watch -Name 'V19_API' -ProjectDir (Join-Path $scriptDir 'src\PerfumeShop.Api') -Port 5000
}
if ($Target -in @('Admin', 'Both')) {
    Start-Watch -Name 'V19_Admin' -ProjectDir (Join-Path $scriptDir 'src\PerfumeShop.Admin') -Port 5207
}
