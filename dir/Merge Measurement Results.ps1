# Merge CSVs in the order defined by first CSV's row numbers
$files = Get-ChildItem -Path "." -Filter "*.csv" |
    Sort-Object {
        $base = $_.BaseName
        if ($base -match "-(\d+)$") { [int]$matches[1] } else { [int]::MaxValue }
    }

if ($files.Count -eq 0) {
    Write-Host "No CSV files found." -ForegroundColor Yellow
    exit 1
}

$output = "merged.csv"
"" | Out-File -FilePath $output -Encoding utf8

# Parse CSV rows
function Parse-DataRows {
    param($file)
    $data = @()
    foreach ($line in Get-Content $file) {
        if ($line -match '^\s*\d+(\.\d+)?\s*,') {
            $cols = $line -split ","
            if ($cols.Count -lt 2) { continue }
            $obj = [PSCustomObject]@{
                NoStr = $cols[0].Trim()
                Name  = $cols[1].Trim()
                Line  = $line
                Cols  = $cols
            }
            $data += $obj
        }
    }
    return $data
}

# Reference order from first CSV
$refData = Parse-DataRows $files[0]
$refData = $refData | Sort-Object {[double]($_.NoStr)}
$refOrder = $refData | Select-Object -Property NoStr, Name

$fieldCount = ($refData[0].Cols).Count

# Process each file
$index = 1
foreach ($file in $files) {
    Add-Content -Path $output -Value "#$index"

    $data = Parse-DataRows $file

    foreach ($refRow in $refOrder) {
        $row = $data | Where-Object { $_.Name -eq $refRow.Name }
        if ($row) {
            # Replace NoStr with reference NoStr, keep the rest
            $cols = $row.Cols
            $cols[0] = $refRow.NoStr
            ($cols -join ",") | Add-Content -Path $output
        } else {
            # Missing row: fill with blanks, but keep reference NoStr
            $empty = @($refRow.NoStr)
            for ($i = 1; $i -lt $fieldCount; $i++) { $empty += "" }
            ($empty -join ",") | Add-Content -Path $output
        }
    }

    Add-Content -Path $output -Value ""
    $index++
}

Write-Host "✅ Merge complete. Output file: $output"
