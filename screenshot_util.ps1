# Screenshot Utility for Claude Code
# Usage:
#   .\screenshot_util.ps1                          # Capture all monitors
#   .\screenshot_util.ps1 -WindowTitle "iCharlotte" # Find window and capture its monitor
#   .\screenshot_util.ps1 -Monitor 1               # Capture specific monitor (0-indexed)

param(
    [string]$WindowTitle = "",
    [int]$Monitor = -1,
    [string]$OutputDir = $PSScriptRoot
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Add user32.dll functions for window detection
$signature = @"
[DllImport("user32.dll")]
public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

[DllImport("user32.dll", SetLastError = true)]
public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

[DllImport("user32.dll")]
public static extern bool IsWindowVisible(IntPtr hWnd);

[DllImport("user32.dll")]
public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

[DllImport("user32.dll")]
public static extern IntPtr GetDesktopWindow();

[DllImport("user32.dll")]
public static extern IntPtr GetTopWindow(IntPtr hWnd);

[StructLayout(LayoutKind.Sequential)]
public struct RECT {
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
}
"@

Add-Type -MemberDefinition $signature -Name Win32 -Namespace User32 -Using System.Text

function Get-AllWindows {
    $windows = @()
    $GW_HWNDNEXT = 2

    $hwnd = [User32.Win32]::GetTopWindow([IntPtr]::Zero)

    while ($hwnd -ne [IntPtr]::Zero) {
        $sb = New-Object System.Text.StringBuilder 256
        [User32.Win32]::GetWindowText($hwnd, $sb, 256) | Out-Null
        $title = $sb.ToString()

        if ($title -and [User32.Win32]::IsWindowVisible($hwnd)) {
            $windows += [PSCustomObject]@{
                Handle = $hwnd
                Title = $title
            }
        }

        $hwnd = [User32.Win32]::GetWindow($hwnd, $GW_HWNDNEXT)
    }

    return $windows
}

function Get-MonitorForWindow {
    param([IntPtr]$WindowHandle)

    $rect = New-Object User32.Win32+RECT
    [User32.Win32]::GetWindowRect($WindowHandle, [ref]$rect) | Out-Null

    # Get center point of window
    $centerX = [int](($rect.Left + $rect.Right) / 2)
    $centerY = [int](($rect.Top + $rect.Bottom) / 2)

    $screens = [System.Windows.Forms.Screen]::AllScreens
    $monitorIndex = 0

    foreach ($screen in $screens) {
        $bounds = $screen.Bounds
        if ($centerX -ge $bounds.X -and $centerX -lt ($bounds.X + $bounds.Width) -and
            $centerY -ge $bounds.Y -and $centerY -lt ($bounds.Y + $bounds.Height)) {
            return @{
                Index = $monitorIndex
                Screen = $screen
            }
        }
        $monitorIndex++
    }

    # Fallback to primary monitor
    return @{
        Index = 0
        Screen = [System.Windows.Forms.Screen]::PrimaryScreen
    }
}

function Capture-Monitor {
    param(
        [System.Windows.Forms.Screen]$Screen,
        [int]$Index,
        [string]$OutputPath
    )

    $bounds = $Screen.Bounds
    $bitmap = New-Object System.Drawing.Bitmap($bounds.Width, $bounds.Height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.CopyFromScreen($bounds.X, $bounds.Y, 0, 0, $bounds.Size)
    $bitmap.Save($OutputPath)
    $graphics.Dispose()
    $bitmap.Dispose()

    Write-Host "SCREENSHOT_SAVED: $OutputPath"
    Write-Host "MONITOR: $Index"
    Write-Host "BOUNDS: $($bounds.X),$($bounds.Y),$($bounds.Width),$($bounds.Height)"
}

# Main logic
$screens = [System.Windows.Forms.Screen]::AllScreens
Write-Host "MONITORS_FOUND: $($screens.Count)"

if ($WindowTitle -ne "") {
    Write-Host "SEARCHING_FOR: $WindowTitle"

    $allWindows = Get-AllWindows

    # Find window containing the search title (case-insensitive)
    $matchedWindow = $allWindows | Where-Object { $_.Title -like "*$WindowTitle*" } | Select-Object -First 1

    if ($matchedWindow) {
        Write-Host "WINDOW_FOUND: $($matchedWindow.Title)"

        $monitorInfo = Get-MonitorForWindow -WindowHandle $matchedWindow.Handle
        Write-Host "WINDOW_ON_MONITOR: $($monitorInfo.Index)"

        $outputPath = Join-Path $OutputDir "screenshot.png"
        Capture-Monitor -Screen $monitorInfo.Screen -Index $monitorInfo.Index -OutputPath $outputPath
    } else {
        Write-Host "WINDOW_NOT_FOUND: $WindowTitle"
        Write-Host "Available windows with similar names:"
        $allWindows | ForEach-Object { Write-Host "  - $($_.Title)" }

        # Fallback: capture all monitors
        Write-Host "FALLBACK: Capturing all monitors"
        for ($i = 0; $i -lt $screens.Count; $i++) {
            $outputPath = Join-Path $OutputDir "screenshot_monitor_$i.png"
            Capture-Monitor -Screen $screens[$i] -Index $i -OutputPath $outputPath
        }
    }
}
elseif ($Monitor -ge 0) {
    # Capture specific monitor
    if ($Monitor -lt $screens.Count) {
        $outputPath = Join-Path $OutputDir "screenshot.png"
        Capture-Monitor -Screen $screens[$Monitor] -Index $Monitor -OutputPath $outputPath
    } else {
        Write-Host "ERROR: Monitor $Monitor not found. Only $($screens.Count) monitors available."
        exit 1
    }
}
else {
    # Capture all monitors
    for ($i = 0; $i -lt $screens.Count; $i++) {
        $outputPath = Join-Path $OutputDir "screenshot_monitor_$i.png"
        Capture-Monitor -Screen $screens[$i] -Index $i -OutputPath $outputPath
    }
}

Write-Host "DONE"
