$AppPath = "D:\Exemple\Exemple\Exemple.exe"

if (-not (Test-Path $AppPath)) {
    Write-Host "Файл не найден: $AppPath" -ForegroundColor Red
    pause
    exit
}

$StartMenuPath = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"
$ShortcutPath = Join-Path $StartMenuPath "Just Color Picker.lnk"
$WorkingDir = Split-Path $AppPath -Parent

$Shell = New-Object -ComObject WScript.Shell
$Shortcut = $Shell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $AppPath
$Shortcut.WorkingDirectory = $WorkingDir
$Shortcut.IconLocation = "$AppPath,0"
$Shortcut.Description = "Just Color Picker"
$Shortcut.Save()

Write-Host "Ярлык создан:" -ForegroundColor Green
Write-Host $ShortcutPath
pause