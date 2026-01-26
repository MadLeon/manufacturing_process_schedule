# Tolerance Converter - ????
# ???????

param(
    [string]$Version = "1.0.0"
)

$ErrorActionPreference = "Stop"

Write-Host "======================================" -ForegroundColor Cyan
Write-Host "Tolerance Converter - ????" -ForegroundColor Cyan
Write-Host "??: $Version" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

# ????
$publishDir = "bin\Publish\Release"
$packageDir = "bin\Package"
$packageName = "ToleranceConverter_v${Version}_win-x64"
$packagePath = "$packageDir\$packageName"

# ??????????
if (-not (Test-Path "$publishDir\ToleranceConverter.exe")) {
    Write-Host "??: ????????" -ForegroundColor Red
    Write-Host "????: dotnet publish" -ForegroundColor Yellow
    exit 1
}

# ??????
Write-Host "??????..." -ForegroundColor Yellow
if (Test-Path $packagePath) {
    Remove-Item $packagePath -Recurse -Force
}
New-Item -ItemType Directory -Path $packagePath -Force | Out-Null

# ?????
Write-Host "??????..." -ForegroundColor Yellow
Copy-Item "$publishDir\ToleranceConverter.exe" -Destination $packagePath

# ??????
Write-Host "??????..." -ForegroundColor Yellow
$readmeContent = @"
# Tolerance Converter v$Version

## ????
1. ?? ToleranceConverter.exe ????
2. ???????????
3. ??????mm?
4. ?? Convert ?????

## ????
- ?? ASME B4.2-1978 ??
- ?? 0-500mm ??
- H12 ????
- ???? .NET ???

## ????
- Windows 7 SP1 ?????
- 64???

## ????
- ??: $Version
- ????: $(Get-Date -Format 'yyyy-MM-dd')
- ????: 65.78 MB

## ???
© 2026 Record Technology and Development
Developed by Leon

????????????
"@

Set-Content -Path "$packagePath\README.txt" -Value $readmeContent -Encoding UTF8

# ????????????
if (Test-Path "RELEASE_NOTES.md") {
    Copy-Item "RELEASE_NOTES.md" -Destination "$packagePath\"
}

# ?? ZIP ???
Write-Host "?? ZIP ???..." -ForegroundColor Yellow
$zipFile = "$packageDir\$packageName.zip"
if (Test-Path $zipFile) {
    Remove-Item $zipFile -Force
}

Compress-Archive -Path "$packagePath\*" -DestinationPath $zipFile -CompressionLevel Optimal

# ????
Write-Host ""
Write-Host "======================================" -ForegroundColor Green
Write-Host "?????" -ForegroundColor Green
Write-Host "======================================" -ForegroundColor Green
Write-Host ""

Write-Host "??????:" -ForegroundColor Yellow
Write-Host "  ???: $packagePath" -ForegroundColor White
Write-Host "  ???: $zipFile" -ForegroundColor White
Write-Host ""

# ??????
Write-Host "????:" -ForegroundColor Yellow
Get-ChildItem $packagePath | ForEach-Object {
    $size = if ($_.Length -gt 1MB) {
        "{0:N2} MB" -f ($_.Length / 1MB)
    } else {
        "{0:N2} KB" -f ($_.Length / 1KB)
    }
    Write-Host "  $($_.Name) - $size" -ForegroundColor White
}
Write-Host ""

if (Test-Path $zipFile) {
    $zipSize = (Get-Item $zipFile).Length / 1MB
    Write-Host "ZIP ?????: $([math]::Round($zipSize, 2)) MB" -ForegroundColor Cyan
}
Write-Host ""

# ?????????
Write-Host "???????? (Y/N): " -NoNewline -ForegroundColor Yellow
$response = Read-Host
if ($response -eq 'Y' -or $response -eq 'y') {
    Start-Process explorer.exe -ArgumentList $packageDir
}

Write-Host ""
Write-Host "????????? ZIP ???????" -ForegroundColor Green
