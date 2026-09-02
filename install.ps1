# Lesson-to-Reviewer Converter PowerShell Installer
$ErrorActionPreference = 'Stop'

$exeUrl = 'https://raw.githubusercontent.com/rem-ctrl/LessonToReviewerConverter/main/dist/LessonToReviewerConverter.exe'
$installDir = "$env:LOCALAPPDATA\LessonToReviewerConverter"
$exePath = "$installDir\LessonToReviewerConverter.exe"
$desktopShortcut = "$env:USERPROFILE\Desktop\LessonToReviewerConverter.lnk"

Write-Host '[+] Installing Lesson-to-Reviewer Converter...' -ForegroundColor Cyan

if (-not (Test-Path $installDir)) {
    New-Item -ItemType Directory -Path $installDir -Force | Out-Null
}

Write-Host '[+] Downloading application executable...' -ForegroundColor Yellow
Invoke-WebRequest -Uri $exeUrl -OutFile $exePath -UseBasicParsing

# Create Desktop Shortcut
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($desktopShortcut)
$Shortcut.TargetPath = $exePath
$Shortcut.WorkingDirectory = $installDir
$Shortcut.Description = 'Lesson-to-Reviewer Converter Desktop App'
$Shortcut.Save()

Write-Host '[+] Installation complete! Shortcut created on Desktop.' -ForegroundColor Green
Write-Host '[+] Launching application...' -ForegroundColor Cyan
Start-Process -FilePath $exePath
