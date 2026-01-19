$sourceDir = "C:\Users\ee\manufacturing_process_schedule\dir"
$files = @("rt-87900-PN-R003(1-4).csv","rt-87900-PN-R003(5-8)NEW.csv","rt-87900-PN-R003(9-10).csv")
$outputFile = Join-Path -Path $sourceDir -ChildPath "rt-87900-PN-R003-merged.csv"
$newOrder = @("2","1","3","4","14","8","6","16","7","9","15","5","13","10","11","12")
$mergedOutput = @("Part report,,,,,,,,",",,,,,,,,",(Get-Content -Path (Join-Path $sourceDir "rt-87900-PN-R003(1-4).csv") -Raw).Split("`r`n")[0..10] + "")
$areaNumber = 1
foreach ($file in $files) {
    $filePath = Join-Path -Path $sourceDir -ChildPath $file
    if (Test-Path $filePath) {
        $content = Get-Content -Path $filePath -Raw
        $lines = $content -split "`r`n"
        $i = 0
        while ($i -lt $lines.Count) {
            if ($lines[$i] -like "*<No.*area>*") {
                $mergedOutput += "<No. $areaNumber area>"
                $i++
                $mergedOutput += $lines[$i]; $i++
                $mergedOutput += $lines[$i]; $i++
                $mergedOutput += $lines[$i]; $i++
                $mergedOutput += $lines[$i]; $i++
                $measurementLines = @()
                while ($i -lt $lines.Count -and $lines[$i] -notlike "*<No.*area>*") {
                    if ($lines[$i] -match "^[0-9]") { $measurementLines += $lines[$i] }
                    if ($lines[$i] -eq "") { $i++; break }
                    $i++
                }
                foreach ($num in $newOrder) {
                    $measurementLines | Where-Object { ($_ -split ",")[0].Trim() -eq $num } | ForEach-Object { $mergedOutput += $_ }
                }
                $mergedOutput += ""
                $areaNumber++
            } else { $i++ }
        }
    }
}
$mergedOutput -join "`r`n" | Set-Content -Path $outputFile -NoNewline
Write-Host "Done: $outputFile"
