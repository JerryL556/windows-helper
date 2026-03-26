$startupFolder = [Environment]::GetFolderPath("Startup")
$launcherPath = Join-Path $startupFolder "Windows Helper.cmd"

if (Test-Path $launcherPath) {
    Remove-Item $launcherPath -Force
    Write-Host "Removed $launcherPath"
} else {
    Write-Host "No startup launcher found."
}
