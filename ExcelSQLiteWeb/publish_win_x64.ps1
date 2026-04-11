param(
  [string]$Configuration = "Release",
  [string]$Runtime = "win-x64"
)

$ErrorActionPreference = "Stop"

Write-Host "Publishing ExcelSQLiteWeb ($Configuration, $Runtime)..." -ForegroundColor Cyan

dotnet --version | Out-Host

dotnet publish .\ExcelSQLiteWeb.csproj `
  -c $Configuration `
  -r $Runtime `
  -p:SelfContained=true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -p:IncludeAllContentForSelfExtract=true `
  -p:EnableCompressionInSingleFile=true `
  -p:PublishReadyToRun=true

Write-Host "Done. Output in: .\bin\$Configuration\net8.0-windows\$Runtime\publish\" -ForegroundColor Green
