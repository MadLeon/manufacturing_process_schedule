# Convert all Excel files in the current folder to PDF
# Supports .xlsx, .xls, .xlsm

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$folder = Get-Location

Get-ChildItem $folder -File |
Where-Object { $_.Extension -in ".xlsx", ".xls", ".xlsm" } |
ForEach-Object {

    Write-Host "Converting $($_.Name)..."

    $workbook = $excel.Workbooks.Open($_.FullName)

    $pdfPath = [System.IO.Path]::ChangeExtension($_.FullName, ".pdf")

    # 0 = xlTypePDF
    $workbook.ExportAsFixedFormat(0, $pdfPath)

    $workbook.Close($false)

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
}

$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Done. All Excel files converted to PDF."
