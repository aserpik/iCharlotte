# Work Activity Monitor - Task Scheduler Setup
# Right-click this file and select "Run with PowerShell" (as Administrator)

$ErrorActionPreference = "Stop"
$ProjectDir = "C:\geminiterminal2"
$PythonExe = (Get-Command python).Source

Write-Host "Setting up Work Activity Monitor scheduled tasks..." -ForegroundColor Cyan
Write-Host "Python: $PythonExe"
Write-Host "Working Dir: $ProjectDir"
Write-Host ""

# 1. Tracker - runs on logon + every 30 min as safety net, restarts on failure
Write-Host "Creating Tracker task (on logon + every 30 min)..." -ForegroundColor Yellow
$trackerAction = New-ScheduledTaskAction `
    -Execute $PythonExe `
    -Argument "-m sideprojects.work_monitor.service" `
    -WorkingDirectory $ProjectDir

$triggerLogon = New-ScheduledTaskTrigger -AtLogOn
$triggerRepeat = New-ScheduledTaskTrigger -Once -At "12:00 AM" `
    -RepetitionInterval (New-TimeSpan -Minutes 30) `
    -RepetitionDuration (New-TimeSpan -Days 365)
$trackerSettings = New-ScheduledTaskSettingsSet `
    -RestartCount 10 `
    -RestartInterval (New-TimeSpan -Minutes 2) `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Days 365)

Register-ScheduledTask `
    -TaskName "WorkMonitor\Tracker" `
    -Action $trackerAction `
    -Trigger @($triggerLogon, $triggerRepeat) `
    -Settings $trackerSettings `
    -Description "Work Activity Monitor - continuous tracking service" `
    -Force

Write-Host "  Tracker task created." -ForegroundColor Green

# 2. Daily Report - 7 PM every day
Write-Host "Creating Daily Report task (7 PM daily)..." -ForegroundColor Yellow
$dailyAction = New-ScheduledTaskAction `
    -Execute $PythonExe `
    -Argument "-m sideprojects.work_monitor.reports.daily" `
    -WorkingDirectory $ProjectDir

$dailyTrigger = New-ScheduledTaskTrigger -Daily -At "7:00 PM"
$dailySettings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

Register-ScheduledTask `
    -TaskName "WorkMonitor\DailyReport" `
    -Action $dailyAction `
    -Trigger $dailyTrigger `
    -Settings $dailySettings `
    -Description "Work Activity Monitor - daily summary report" `
    -Force

Write-Host "  Daily Report task created." -ForegroundColor Green

# 3. Weekly Report - Friday 7 PM
Write-Host "Creating Weekly Report task (Friday 7 PM)..." -ForegroundColor Yellow
$weeklyAction = New-ScheduledTaskAction `
    -Execute $PythonExe `
    -Argument "-m sideprojects.work_monitor.reports.weekly" `
    -WorkingDirectory $ProjectDir

$weeklyTrigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Friday -At "7:00 PM"
$weeklySettings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

Register-ScheduledTask `
    -TaskName "WorkMonitor\WeeklyReport" `
    -Action $weeklyAction `
    -Trigger $weeklyTrigger `
    -Settings $weeklySettings `
    -Description "Work Activity Monitor - weekly automation analysis" `
    -Force

Write-Host "  Weekly Report task created." -ForegroundColor Green

Write-Host ""
Write-Host "All 3 tasks created successfully!" -ForegroundColor Cyan
Write-Host ""
Write-Host "Tasks created under 'WorkMonitor' folder in Task Scheduler:"
Write-Host "  - Tracker        : runs on logon + every 30 min, skips if already running"
Write-Host "  - DailyReport    : 7:00 PM every day"
Write-Host "  - WeeklyReport   : 7:00 PM every Friday"
Write-Host ""
Write-Host "To start the tracker now, run:"
Write-Host "  schtasks /run /tn 'WorkMonitor\Tracker'"
Write-Host ""
Read-Host "Press Enter to close"
