Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Get all screens
$screens = [System.Windows.Forms.Screen]::AllScreens
Write-Host "Found $($screens.Count) monitors"

foreach ($i in 0..($screens.Count - 1)) {
    $screen = $screens[$i]
    Write-Host "Monitor $i : $($screen.Bounds)"

    $bounds = $screen.Bounds
    $bmp = New-Object System.Drawing.Bitmap($bounds.Width, $bounds.Height)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($bounds.X, $bounds.Y, 0, 0, $bounds.Size)
    $bmp.Save("C:\geminiterminal2\screenshot_monitor_$i.png")
    $g.Dispose()
    $bmp.Dispose()
    Write-Host "Saved monitor $i"
}

Write-Host "Done"
