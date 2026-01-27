# Create shortcut pointing to PyInstaller-built iCharlotte.exe
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("C:\geminiterminal2\iCharlotte.lnk")
$Shortcut.TargetPath = "C:\geminiterminal2\dist\iCharlotte\iCharlotte.exe"
$Shortcut.WorkingDirectory = "C:\geminiterminal2"
$Shortcut.IconLocation = "C:\geminiterminal2\icharlotte.ico,0"
$Shortcut.Description = "iCharlotte Legal Document Management"
$Shortcut.Save()

Write-Host "Shortcut created: C:\geminiterminal2\iCharlotte.lnk"

# Copy to Desktop
$desktop = [Environment]::GetFolderPath('Desktop')
Copy-Item "C:\geminiterminal2\iCharlotte.lnk" $desktop -Force
Write-Host "Copied to Desktop: $desktop"
