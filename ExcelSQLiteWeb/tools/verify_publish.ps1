param(
  [Parameter(Mandatory=$true)]
  [string]$PublishDir
)

$ErrorActionPreference = "Stop"

function Assert-File([string]$p) {
  if (!(Test-Path $p)) {
    throw "Missing: $p"
  }
}

Write-Host "Verify publish output: $PublishDir" -ForegroundColor Cyan

Assert-File (Join-Path $PublishDir "ExcelSQLiteWeb.exe")
Assert-File (Join-Path $PublishDir "index.html")
Assert-File (Join-Path $PublishDir "app.js")
Assert-File (Join-Path $PublishDir "index.normal.html")
Assert-File (Join-Path $PublishDir "index.expert.html")
Assert-File (Join-Path $PublishDir "assets\\codemirror\\codemirror.js")
Assert-File (Join-Path $PublishDir "assets\\codemirror\\codemirror.css")

Write-Host "OK: core files present." -ForegroundColor Green
Write-Host "" 
Write-Host "Manual checks:" -ForegroundColor Yellow
Write-Host "  1) Start ExcelSQLiteWeb.exe (should show WEB:disk in title)"
Write-Host "  2) Click Ribbon: 系统 -> 用户模式 (or press Ctrl+Shift+U)"
Write-Host "  3) Ctrl+Shift+M shows userMode=expert and visibleTabs includes quality/expert"

