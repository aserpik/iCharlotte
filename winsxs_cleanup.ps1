$ErrorActionPreference = 'Continue'
$sentinel = 'C:\geminiterminal2\logs\winsxs_done.txt'
Remove-Item $sentinel -Force -ErrorAction SilentlyContinue
$before = [math]::Round((Get-PSDrive C).Free / 1GB, 2)
# Analyze, then clean up superseded component-store (WinSxS) data. No /ResetBase (keeps update uninstall ability).
$analyze = Dism /Online /Cleanup-Image /AnalyzeComponentStore 2>&1 | Out-String
$clean   = Dism /Online /Cleanup-Image /StartComponentCleanup 2>&1 | Out-String
$after = [math]::Round((Get-PSDrive C).Free / 1GB, 2)
$out = @()
$out += "FREE_BEFORE_GB: $before"
$out += "FREE_AFTER_GB: $after"
$out += "FREED_GB: " + [math]::Round($after - $before, 2)
$out += "--- ANALYZE (tail) ---"
$out += ($analyze -split "`r?`n" | Where-Object { $_ -match 'Actual Size|Reclaimable|Recommended|Component Store' })
$out += "--- CLEANUP (tail) ---"
$out += ($clean -split "`r?`n" | Where-Object { $_ -match 'completed|error|Error|Version' } | Select-Object -Last 8)
$out += "DONE"
Set-Content -Path $sentinel -Value ($out -join "`r`n")
