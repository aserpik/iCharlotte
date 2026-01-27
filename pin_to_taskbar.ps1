$shell = New-Object -ComObject Shell.Application
$folder = $shell.Namespace("C:\geminiterminal2")
$item = $folder.ParseName("iCharlotte.lnk")
$verbs = $item.Verbs()
foreach ($verb in $verbs) {
    if ($verb.Name -match "pin.*taskbar|taskbar.*pin") {
        $verb.DoIt()
        Write-Host "Pinned to taskbar!"
        exit 0
    }
}
Write-Host "Taskbar pin verb not found. Available verbs:"
foreach ($verb in $verbs) {
    Write-Host "  - $($verb.Name)"
}
