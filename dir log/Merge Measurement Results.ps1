# Merge CSVs in the order defined by first CSV's row numbers

# Ask user if they want to keep the "0." prefix (default: no)
$response = Read-Host "Keep 0. prefix? (y=yes, n=no, default: n)"
if ([string]::IsNullOrWhiteSpace($response)) {
    $response = "n"
}
$keepZeroPrefix = ($response -eq "y" -or $response -eq "Y")
Write-Host "Prefix setting: $(if ($keepZeroPrefix) { 'Keep 0.' } else { 'Remove 0.' })" -ForegroundColor Green

# Function to format decimal values to 4 decimal places
function Format-DecimalValue {
    param(
        [string]$value,
        [bool]$keepPrefix
    )
    
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $value
    }
    
    # Try to parse as decimal
    $doubleValue = 0
    if ([double]::TryParse($value, [ref]$doubleValue)) {
        # Format to 4 decimal places
        $formatted = $doubleValue.ToString("F4")
        
        # If not keeping prefix and starts with "0.", remove the "0"
        if (!$keepPrefix -and $formatted -match '^0\.') {
            $formatted = $formatted.Substring(1)
        }
        
        # Create a formula that forces text display with exact format
        # Use the formula =TEXT(...) to preserve the format
        # But wrap the result in quotes to force text display
        return "=""" + $formatted + """"
    }
    
    # If not a number, return as is
    return $value
}

$files = @(Get-ChildItem -Path "." -Filter "*.csv" |
    Where-Object { $_.Name -ne "merged.csv" -and $_.Name -ne "merged.xlsx" } |
    Sort-Object {
        $base = $_.BaseName
        # Try to extract number from parentheses first: (1), (2), or complex formats like (72373-1)
        if ($base -match '\((\d+)\)$') {
            # Simple format: (1), (2), etc.
            [int]$matches[1]
        }
        elseif ($base -match '\(([^)]*-\d+)\)') {
            # Complex format: (xxx-1), (xxx-2), etc.
            $numStr = $matches[1]
            if ($numStr -match '-(\d+)$') { [int]$matches[1] } else { [int]::MaxValue }
        }
        # Otherwise try to extract from end with dash: -1, -2
        elseif ($base -match "-(\d+)$") { 
            [int]$matches[1] 
        } 
        else { 
            [int]::MaxValue 
        }
    })

if ($files.Count -eq 0) {
    Write-Host "No CSV files found." -ForegroundColor Yellow
    exit 1
}

$output = "merged.csv"
"" | Out-File -FilePath $output -Encoding utf8

# Detect if file has explicit area headers
function HasExplicitAreas {
    param($file)
    $lines = @(Get-Content $file)
    foreach ($line in $lines) {
        if ($line -match "^<No\.\s+\d+\s+area>") {
            return $true
        }
    }
    return $false
}

# Detect format from first file
# This determines how ALL files in the batch will be parsed
# - If first file has explicit <No. # area> headers: multi-area format
# - If first file only has "Measurement results" sections: single-area format
$useExplicitAreas = HasExplicitAreas $files[0].FullName
Write-Host "Format detection: $(if ($useExplicitAreas) { 'Multi-area format (with area headers)' } else { 'Single area format (measurement results only)' })" -ForegroundColor Cyan

# Parse CSV rows - collect data from all areas, grouped by area
function Parse-DataRows {
    param($file, $expectExplicitAreas)
    $allAreas = @()
    $currentArea = @()
    $inDataSection = $false
    $areaName = ""
    $hasExplicitAreas = $false
    $skipNextLine = $false
    
    $lines = @(Get-Content $file)
    
    for ($lineNum = 0; $lineNum -lt $lines.Count; $lineNum++) {
        $line = $lines[$lineNum]
        
        # Skip empty lines
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        
        # Detect area header - only process if we expect explicit areas
        if ($expectExplicitAreas -and $line -match "^<No\.\s+\d+\s+area>") {
            $hasExplicitAreas = $true
            # Save previous area if exists
            if ($currentArea.Count -gt 0) {
                $allAreas += @{
                    Name = $areaName
                    Data = $currentArea
                }
            }
            $currentArea = @()
            $areaName = $line.Trim()
            $inDataSection = $false
            $skipNextLine = $true
            continue
        }
        
        # Check if we're entering a measurement results section
        if ($line -match "Measurement results") {
            $inDataSection = $true
            $skipNextLine = $true
            # If no explicit areas found, create default area name
            if (!$hasExplicitAreas -and $areaName -eq "") {
                $areaName = "Default Area"
            }
            continue
        }
        
        # Skip the header row after "Measurement results"
        if ($skipNextLine -and $inDataSection) {
            $skipNextLine = $false
            continue
        }
        
        # Parse data rows: must have a number (possibly decimal) followed by comma
        if ($inDataSection) {
            $trimmed = $line.Trim()
            # Check if line starts with a number followed by comma
            if ($trimmed -match '^(\d+(?:\.\d+)?)\s*,') {
                $cols = $line -split ","
                if ($cols.Count -lt 2) { continue }
                
                $colZero = $cols[0].Trim()
                $colOne = $cols[1].Trim()
                
                # Skip any remaining header indicators
                if ($colZero -eq "No." -or $colOne -eq "measurement item") {
                    continue
                }
                
                # Skip lines that aren't measurement data
                if ($colZero -eq "" -or ![double]::TryParse($colZero, [ref]$null)) {
                    continue
                }
                
                $obj = [PSCustomObject]@{
                    NoStr = $colZero
                    Name  = $colOne
                    Line  = $line
                    Cols  = $cols
                }
                $currentArea += $obj
            }
        }
    }
    
    # Save last area
    if ($currentArea.Count -gt 0) {
        $allAreas += @{
            Name = $areaName
            Data = $currentArea
        }
    }
    
    return $allAreas
}

# Reference order from first CSV - build map for each area by position
$refAreas = @(Parse-DataRows $files[0].FullName $useExplicitAreas)
if ($refAreas.Count -eq 0) {
    Write-Host "No data areas found." -ForegroundColor Red
    exit 1
}

Write-Host "Reference file: $($files[0].Name), Areas found: $($refAreas.Count)" -ForegroundColor Cyan
for ($i = 0; $i -lt $refAreas.Count; $i++) {
    Write-Host "  Area $i ($($refAreas[$i].Name)): $($refAreas[$i].Data.Count) rows" -ForegroundColor Cyan
}

# Create reference orders indexed by area position (1st area, 2nd area, etc)
$refOrdersByPosition = @{}
$fieldCount = 0

for ($i = 0; $i -lt $refAreas.Count; $i++) {
    $areaData = $refAreas[$i].Data
    
    if ($areaData -and $areaData.Count -gt 0) {
        $refData = $areaData | Sort-Object {[double]($_.NoStr)}
        $refOrdersByPosition[$i] = $refData | Select-Object -Property NoStr, Name
        if ($fieldCount -eq 0 -and $refData[0].Cols) {
            $fieldCount = ($refData[0].Cols).Count
        }
    }
}

# Process each file
$index = 1
foreach ($file in $files) {
    Write-Host "Processing: $($file.Name)"
    Add-Content -Path $output -Value "#$index"

    $allAreas = @(Parse-DataRows $file.FullName $useExplicitAreas)

    # Collect all rows from all areas, then fill global gaps
    $allAreaRows = @()
    
    # Process each area in this file
    for ($areaIdx = 0; $areaIdx -lt $allAreas.Count; $areaIdx++) {
        $areaGroup = $allAreas[$areaIdx]
        $data = $areaGroup.Data
        
        # Try to match reference order:
        # 1. First try to match by position
        # 2. If no data found, try to find first available reference order
        $refOrder = $refOrdersByPosition[$areaIdx]
        if (!$refOrder -or $refOrder.Count -eq 0) {
            # Try using the first reference order as default
            $refOrder = $refOrdersByPosition[0]
        }
        
        if (!$refOrder) {
            Write-Host "Warning: No reference order for area position $areaIdx" -ForegroundColor Yellow
            continue
        }
        
        # Collect all rows from reference order
        foreach ($refRow in $refOrder) {
            $row = $data | Where-Object { $_.Name -eq $refRow.Name }
            if ($row) {
                # Create a copy of the columns to avoid reference issues
                $cols = @($row.Cols)
                $cols[0] = $refRow.NoStr
                # Format C column (index 2) if it exists
                if ($cols.Count -gt 2) {
                    $formattedValue = Format-DecimalValue -value $cols[2] -keepPrefix $keepZeroPrefix
                    $cols[2] = $formattedValue
                }
                $allAreaRows += @{
                    NoStr = [double]$refRow.NoStr
                    Line  = ($cols -join ",")
                }
            } else {
                # Missing row: fill with blanks, but keep reference NoStr
                $empty = @($refRow.NoStr)
                for ($i = 1; $i -lt $fieldCount; $i++) { $empty += "" }
                $allAreaRows += @{
                    NoStr = [double]$refRow.NoStr
                    Line  = ($empty -join ",")
                }
            }
        }
    }
    
    # Find gaps in integer row numbers across all areas and add empty rows
    $integerParts = @()
    foreach ($row in $allAreaRows) {
        if ($row.NoStr -match '^(\d+)') {
            $intVal = [int]$matches[1]
            if ($integerParts -notcontains $intVal) {
                $integerParts += $intVal
            }
        }
    }
    
    $integerParts = $integerParts | Sort-Object -Unique
    
    if ($integerParts.Count -gt 0) {
        $minRow = $integerParts[0]
        $maxRow = $integerParts[-1]
        
        for ($rowNum = $minRow; $rowNum -le $maxRow; $rowNum++) {
            if ($integerParts -notcontains $rowNum) {
                # This row number is missing, add empty row
                $empty = @($rowNum.ToString())
                for ($i = 1; $i -lt $fieldCount; $i++) { $empty += "" }
                $allAreaRows += @{
                    NoStr = [double]$rowNum
                    Line  = ($empty -join ",")
                }
            }
        }
    }
    
    # Sort all rows by number and output
    $allAreaRows = $allAreaRows | Sort-Object { $_.NoStr }
    foreach ($row in $allAreaRows) {
        $row.Line | Add-Content -Path $output
    }

    Add-Content -Path $output -Value ""
    $index++
}

Write-Host "✅ Merge complete. Output file: $output" -ForegroundColor Green
