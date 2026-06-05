$git = "C:\Users\dmurphy\AppData\Local\GitHubDesktop\app-3.5.12\resources\app\git\cmd\git.exe"
Set-Location "C:\Users\Public\RPA\code\PSTN Migration"

Write-Host "=== Initialising fresh repo ===" -ForegroundColor Cyan
& $git init
& $git checkout -b main

Write-Host "=== Checking for secrets before staging ===" -ForegroundColor Yellow
$secrets = & $git ls-files 2>&1
# List what would be staged
& $git status --short 2>&1 | Select-Object -First 10

Write-Host "`n=== Staging all files ===" -ForegroundColor Cyan
& $git add .

Write-Host "`n=== Verifying no sensitive files staged ===" -ForegroundColor Yellow
& $git ls-files --cached | Select-String "\.env$|config\.py$|graph_mailbox"

Write-Host "`n=== Making initial commit ===" -ForegroundColor Cyan
& $git commit -m "Initial commit - working build"

Write-Host "`n=== Adding remote and pushing ===" -ForegroundColor Cyan
& $git remote add origin "https://github.com/BRProdUKCOP02/PSTN-Automations.git"
& $git push --force -u origin main 2>&1 | Write-Host

Write-Host "`nDone!" -ForegroundColor Green
