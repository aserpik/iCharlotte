# Run as Administrator: Right-click > Run with PowerShell
$ErrorActionPreference = "Stop"

# 1. Enable file system auditing
Write-Host "Enabling file system auditing..." -ForegroundColor Yellow
auditpol /set /subcategory:"File System" /success:enable /failure:enable

# 2. Set audit rule on sideprojects folder
Write-Host "Setting deletion audit on sideprojects/..." -ForegroundColor Yellow
$acl = Get-Acl "C:\geminiterminal2\sideprojects"
$rule = New-Object System.Security.AccessControl.FileSystemAuditRule(
    "Everyone", "Delete,DeleteSubdirectoriesAndFiles",
    "ContainerInherit,ObjectInherit", "None", "Success,Failure"
)
$acl.AddAuditRule($rule)
Set-Acl "C:\geminiterminal2\sideprojects" $acl
Write-Host "  Audit rule applied." -ForegroundColor Green

# 3. Register scheduled tasks
Write-Host "Registering scheduled tasks..." -ForegroundColor Yellow
$python = (Get-Command python).Source
$dir = "C:\geminiterminal2"

Register-ScheduledTask -TaskName "WorkMonitor\Tracker" `
    -Action (New-ScheduledTaskAction -Execute $python -Argument "-m sideprojects.work_monitor.service" -WorkingDirectory $dir) `
    -Trigger @(
        (New-ScheduledTaskTrigger -AtLogOn),
        (New-ScheduledTaskTrigger -Once -At "12:00 AM" -RepetitionInterval (New-TimeSpan -Minutes 30) -RepetitionDuration (New-TimeSpan -Days 365))
    ) `
    -Settings (New-ScheduledTaskSettingsSet -RestartCount 10 -RestartInterval (New-TimeSpan -Minutes 2) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Days 365)) `
    -Force
Write-Host "  Tracker registered." -ForegroundColor Green

Register-ScheduledTask -TaskName "WorkMonitor\DailyReport" `
    -Action (New-ScheduledTaskAction -Execute $python -Argument "-m sideprojects.work_monitor.reports.daily" -WorkingDirectory $dir) `
    -Trigger (New-ScheduledTaskTrigger -Daily -At "7:00 PM") `
    -Settings (New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 30)) `
    -Force
Write-Host "  DailyReport registered." -ForegroundColor Green

Register-ScheduledTask -TaskName "WorkMonitor\WeeklyReport" `
    -Action (New-ScheduledTaskAction -Execute $python -Argument "-m sideprojects.work_monitor.reports.weekly" -WorkingDirectory $dir) `
    -Trigger (New-ScheduledTaskTrigger -Weekly -DaysOfWeek Friday -At "7:00 PM") `
    -Settings (New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 30)) `
    -Force
Write-Host "  WeeklyReport registered." -ForegroundColor Green

Write-Host ""
Write-Host "Done! All tasks registered + deletion auditing enabled." -ForegroundColor Cyan
Read-Host "Press Enter to close"
