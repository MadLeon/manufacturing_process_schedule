# Tolerance Converter - Release Notes

## Latest Release

**Version**: v1.0  
**Date**: 2024

### Published Files

1. **Single-file EXE**: `ToleranceConverter.exe` - 65.78 MB
   - Location: `bin\Release\net6.0-windows\win-x64\publish\`
   - Self-contained with .NET runtime included

2. **Compressed Archive**: `ToleranceConverter_v1.0_Win-x64.zip` - 61.26 MB
   - Easy distribution and download
   - Compression ratio: ~93%

## Quick Publish

### Method 1: Batch File (Recommended)
```cmd
publish.bat
```
Automatically: Clean ? Build ? Publish ? Compress

### Method 2: PowerShell Script
```powershell
.\publish.ps1
```
Colored output with detailed information

### Method 3: Manual Command
```cmd
dotnet publish ToleranceConverter.csproj -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true /p:DebugType=None /p:DebugSymbols=false
```

## Optimization Settings

Current publish configuration includes:

- PublishSingleFile: Single executable file
- SelfContained: .NET 6 runtime included
- EnableCompressionInSingleFile: Internal compression enabled
- PublishReadyToRun: AOT pre-compilation for fast startup
- DebugType=None: Remove debug symbols
- DebugSymbols=false: No PDB files
- ZIP compression: Further reduce distribution size

## File Size Comparison

| Type | Size | Notes |
|------|------|-------|
| Original EXE | 65.78 MB | Self-contained single file |
| ZIP Archive | 61.26 MB | Compressed (93%) |

**Note**: Windows Forms apps cannot use IL Trimming. This is the optimal size.

## System Requirements

- Windows 7 SP1 or higher
- 64-bit system
- No .NET installation required

## Distribution

- Direct: Share `ToleranceConverter.exe`
- Web: Upload `ToleranceConverter_v1.0_Win-x64.zip`
- Professional: Use installer (Inno Setup, WiX)

## Features

- Single file - no dependencies
- Embedded tolerance data (ASME B4.2-1978 H12)
- Enter key support for quick conversion
- Input validation (numbers and decimals only)
- Clean, modern UI
