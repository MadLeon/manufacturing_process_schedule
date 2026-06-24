# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A standalone C# WinForms application that looks up dimensional tolerances per **ANSI/ASME B4.2-1978** standard. Given a nominal dimension (mm) and fit type (hole/shaft), it returns upper/lower tolerance bounds.

## Build & Run

```powershell
# Run directly (debug)
dotnet run

# Build debug
dotnet build

# Build self-contained single-file release (win-x64)
dotnet publish -c Release
```

Two `.csproj` files exist:
- `ToleranceConverter.csproj` — self-contained, single-file, win-x64 (use this for releases)
- `ToleranceConverter-FrameworkDependent.csproj` — requires .NET 6 runtime installed

## Architecture

### Data Flow

1. `ToleranceDataService` (in `ToleranceData.cs`) loads `tolerance_table.json` as an **embedded resource** at startup via `Assembly.GetManifestResourceStream`.
2. `GetTolerance(dimension, isInternal)` finds the matching range entry using `dimension > minRange && dimension <= maxRange`.
3. `ToleranceConverterForm.cs` calls the service on button click and displays results.

### tolerance_table.json Schema

The embedded JSON has two top-level arrays `"internal"` (hole/H12) and `"external"` (shaft/h12), each entry:
```json
{ "minRange": 0, "maxRange": 3, "upper": 0.100, "lower": 0.000}
```
Values are in **millimeters**. Range is exclusive on min, inclusive on max (`(minRange, maxRange]`). Valid range: 0–500 mm.

When adding a new tolerance category (e.g., IT12/2), add a new top-level key to the JSON and extend `ToleranceTable` with a matching `List<ToleranceRange>` property.

### Key Classes

| File | Purpose |
|------|---------|
| `ToleranceData.cs` | `ToleranceRange` (data model), `ToleranceTable` (JSON root), `ToleranceDataService` (lookup logic) |
| `ToleranceConverterForm.cs` | UI event handlers; reads from `_dataService` |
| `ToleranceConverterForm.Designer.cs` | Auto-generated WinForms layout — edit via designer or carefully by hand |
| `tolerance_table.json` | Embedded tolerance data; must be declared as `<EmbeddedResource>` in `.csproj` |

### Embedded Resource Loading

The resource name follows the namespace + filename pattern: `ToleranceConverter.tolerance_table.json`. If the JSON file is renamed or moved, update the `resourceName` string in `ToleranceDataService.LoadDataFromEmbeddedResource()`.
