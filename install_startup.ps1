$startupFolder = [Environment]::GetFolderPath("Startup")
$pythonw = (Get-Command pythonw.exe).Source
$scriptPath = Join-Path $PSScriptRoot "windows_helper.py"
$launcherPath = Join-Path $startupFolder "Windows Helper.cmd"

if (-not (Test-Path $scriptPath)) {
    throw "Could not find windows_helper.py at $scriptPath"
}

$content = @(
    '@echo off'
    'cd /d "' + $PSScriptRoot + '"'
    '"' + $pythonw + '" "' + $scriptPath + '"'
)

Set-Content -Path $launcherPath -Value $content -Encoding ASCII
Write-Host "Startup launcher installed at $launcherPath"
