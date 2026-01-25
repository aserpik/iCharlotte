# start_dev.ps1

# 1. Check Node Dependencies
$noteTakerPath = Join-Path $PSScriptRoot "NoteTaker"
if (-not (Test-Path (Join-Path $noteTakerPath "node_modules"))) {
    Write-Host "First time setup: Installing NoteTaker dependencies..." -ForegroundColor Yellow
    Push-Location $noteTakerPath
    npm install
    Pop-Location
}

# 2. Start React Dev Server (New Window)
Write-Host "Starting React Dev Server (NoteTaker)..." -ForegroundColor Green
# We launch a new PowerShell window that stays open (-NoExit) so you can see logs/errors
Start-Process powershell -ArgumentList "-NoExit", "-Command", "cd '$noteTakerPath'; npm run dev"

# 3. Configure Environment
$env:NOTE_TAKER_DEV_URL = "http://localhost:5173"
Write-Host "Dev Mode Enabled: URL set to $env:NOTE_TAKER_DEV_URL" -ForegroundColor Gray

# 4. Start Python Process
Write-Host "Starting iCharlotte..." -ForegroundColor Cyan
Write-Host "Use the 'Restart' button in the app to reload after .py changes." -ForegroundColor Gray
Write-Host "Edit any .tsx file to hot-reload the UI." -ForegroundColor Gray

python dev_runner.py
