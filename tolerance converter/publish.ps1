# Tolerance Converter - ??????
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "Tolerance Converter - ?????" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

# Clean previous builds
Write-Host "[1/4] ?????..." -ForegroundColor Yellow
dotnet clean ToleranceConverter.csproj -c Release
Write-Host ""

# Publish self-contained version
Write-Host "[2/4] ????? (?????)..." -ForegroundColor Green
dotnet publish ToleranceConverter.csproj -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true /p:DebugType=None /p:DebugSymbols=false -o "bin\Publish\Release"
Write-Host ""

# Create compressed archive
Write-Host "[3/4] ?????..." -ForegroundColor Green
$version = "v1.0"
$zipPath = "ToleranceConverter_${version}_Win-x64.zip"
Compress-Archive -Path "bin\Publish\Release\ToleranceConverter.exe" -DestinationPath $zipPath -Force
Write-Host ""

# Display results
Write-Host "[4/4] ????:" -ForegroundColor Yellow
Write-Host ""
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "????!" -ForegroundColor Green
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "????:" -ForegroundColor Yellow
Write-Host "  ??? EXE: bin\Publish\Release\ToleranceConverter.exe" -ForegroundColor White
Write-Host "  ???:     $zipPath" -ForegroundColor White
Write-Host ""

# Show file sizes
Write-Host "????:" -ForegroundColor Yellow
if (Test-Path "bin\Publish\Release\ToleranceConverter.exe") {
    $size1 = (Get-Item "bin\Publish\Release\ToleranceConverter.exe").Length / 1MB
    Write-Host ("  EXE ??:  {0:N2} MB" -f $size1) -ForegroundColor White
}
if (Test-Path $zipPath) {
    $size2 = (Get-Item $zipPath).Length / 1MB
    Write-Host ("  ZIP ???: {0:N2} MB (???: {1:P0})" -f $size2, ($size2/$size1)) -ForegroundColor White
}
Write-Host ""

Write-Host "????:" -ForegroundColor Cyan
Write-Host "  - ????????? .NET 6??????" -ForegroundColor Gray
Write-Host "  - ?? Windows 7 SP1 ??? 64 ???" -ForegroundColor Gray
Write-Host ""

Write-Host "??????..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
