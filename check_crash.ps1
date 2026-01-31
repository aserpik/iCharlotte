Get-WinEvent -LogName Application -MaxEvents 30 | Where-Object { $_.Level -eq 2 } | Select-Object TimeCreated, Id, Message | Format-List
