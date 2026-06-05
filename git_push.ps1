$git = "C:\Users\dmurphy\AppData\Local\GitHubDesktop\app-3.5.12\resources\app\git\cmd\git.exe"
Set-Location "C:\Users\Public\RPA\code\PSTN Migration"

Write-Host "=== Staging .gitignore ===" -ForegroundColor Cyan
& $git add .gitignore

Write-Host "=== Committing cleanup ===" -ForegroundColor Cyan
& $git commit -m "Add .gitignore, remove sensitive files and output data"

Write-Host "`n=== Adding remote ===" -ForegroundColor Cyan
& $git remote add origin "https://github.com/BRProdUKCOP02/PSTN-Automations.git" 2>&1 | Write-Host

Write-Host "`n=== Setting branch to main ===" -ForegroundColor Cyan
& $git branch -M main

Write-Host "`n=== Pushing to GitHub ===" -ForegroundColor Cyan
& $git push -u origin main 2>&1 | Write-Host

Write-Host "`nDone!" -ForegroundColor Green
