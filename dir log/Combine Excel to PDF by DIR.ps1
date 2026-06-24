# This script finds all "DIR 1", "DIR 2", etc. folders and merges all Excel files in each folder into a single PDF
# Requires: Adobe Acrobat Pro

param(
    [string]$ParentFolder = (Get-Location)
)

# Create Excel application object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Verify Adobe Acrobat is available (we'll need it for PDF merging)
try {
    $testAcrobat = New-Object -ComObject "AcroExch.PDDoc"
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($testAcrobat) | Out-Null
}
catch {
    Write-Host "Error: Adobe Acrobat Pro not detected. Please ensure it is installed."
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    exit 1
}

# Get all "DIR *" folders, sorted by Windows file name order
$dirFolders = Get-ChildItem $ParentFolder -Directory |
    Where-Object { $_.Name -match '^DIR\s+\d+$' } |
    Sort-Object { [int]($_.Name -replace 'DIR\s+', '') }

if ($dirFolders.Count -eq 0) {
    Write-Host "No 'DIR *' folders found"
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    exit
}

foreach ($dirFolder in $dirFolders) {
    Write-Host "Processing folder: $($dirFolder.Name)"
    
    # Get all Excel files in this folder, sorted by Windows file name order (ascending)
    $excelFiles = Get-ChildItem $dirFolder.FullName -File |
        Where-Object { $_.Extension -in ".xlsx", ".xls", ".xlsm" } |
        Sort-Object Name
    
    if ($excelFiles.Count -eq 0) {
        Write-Host "  - No Excel files in $($dirFolder.Name), skipping"
        continue
    }
    
    Write-Host "  - Found $($excelFiles.Count) Excel files"
    
    # Temporary folder for storing PDF files
    $tempPdfFolder = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "PDFMerge_$(Get-Random)")
    New-Item -ItemType Directory -Path $tempPdfFolder -Force | Out-Null
    
    $pdfFiles = @()
    
    # Convert each Excel file to PDF
    foreach ($excelFile in $excelFiles) {
        Write-Host "    Converting: $($excelFile.Name)"
        
        try {
            $workbook = $excel.Workbooks.Open($excelFile.FullName, $null, $false)
            
            $pdfFileName = [System.IO.Path]::ChangeExtension($excelFile.Name, ".pdf")
            $pdfPath = Join-Path $tempPdfFolder $pdfFileName
            
            # 0 = xlTypePDF
            $workbook.ExportAsFixedFormat(0, $pdfPath)
            
            # Verify PDF was created
            if (Test-Path $pdfPath) {
                $pdfFiles += $pdfPath
                Write-Host "      -> OK"
            }
            else {
                Write-Host "    Error: PDF file was not created"
            }
            
            $workbook.Close($false)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
            
            # Small delay to ensure file is written
            Start-Sleep -Milliseconds 100
        }
        catch {
            Write-Host "    Error: Failed to convert $($excelFile.Name): $_"
        }
    }
    
    
    if ($pdfFiles.Count -eq 0) {
        Write-Host "  - No successfully converted PDFs, skipping"
        Remove-Item $tempPdfFolder -Recurse -Force
        continue
    }
    
    # Create output path for merged PDF
    $outputPdfName = "$($dirFolder.Name)_Combined.pdf"
    $outputPdfPath = Join-Path $ParentFolder $outputPdfName
    
    Write-Host "  - Merging PDF files..."
    
    try {
        # Create a new empty PDF document
        $pdDoc = New-Object -ComObject "AcroExch.PDDoc"
        $pdDoc.Create() | Out-Null
        
        # Insert all PDFs in order
        for ($i = 0; $i -lt $pdfFiles.Count; $i++) {
            Write-Host "    Adding: $(Split-Path $pdfFiles[$i] -Leaf)"
            
            # Open the PDF to insert
            $insertDoc = New-Object -ComObject "AcroExch.PDDoc"
            $insertOpened = $insertDoc.Open($pdfFiles[$i])
            
            if ($insertOpened) {
                # Get number of pages to insert
                $numPagesToInsert = $insertDoc.GetNumPages()
                
                # Get current page count (to insert after last page)
                $currentPageCount = $pdDoc.GetNumPages()
                $insertAfterPage = $currentPageCount - 1
                
                # Insert all pages
                # Parameters: insertAfterThisPage, sourcePDDoc, startPage, numPages, addBookmarks
                $insertResult = $pdDoc.InsertPages($insertAfterPage, $insertDoc, 0, $numPagesToInsert, 0) | Out-Null
                
                $insertDoc.Close() | Out-Null
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($insertDoc) | Out-Null
            }
            else {
                Write-Host "    Error: Could not open $($pdfFiles[$i])"
            }
        }
        
        # Save the merged PDF
        $saveResult = $pdDoc.Save(1, $outputPdfPath) | Out-Null
        
        Write-Host "  - Complete: $outputPdfName (Total pages: $($pdDoc.GetNumPages()))"
        
        # Close the document
        try {
            $pdDoc.Close() | Out-Null
        }
        catch {
            # Ignore close errors - file is already saved
        }
        
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null
    }
    catch {
        Write-Host "  - Error: Failed to merge PDFs: $_"
    }
    finally {
        # Clean up temporary files
        try {
            Remove-Item $tempPdfFolder -Recurse -Force
        }
        catch {
            # Ignore cleanup errors
        }
    }
}

# Clean up Excel COM object
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host ""
Write-Host "All done"
