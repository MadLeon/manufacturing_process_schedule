# Convert all Excel files in the current folder to PDF, then merge them into one file
# Supports .xlsx, .xls, .xlsm
# Requires: Adobe Acrobat Pro

param(
    [string]$OutputFileName = "Combined.pdf"
)

# ==================== DEBUG LOGGING FUNCTION ====================
$debugLog = @()

function Add-DebugLog {
    param([string]$Message)
    Write-Host $Message
}

# ==================== SYSTEM INFO ====================
Add-DebugLog "========== EXCEL TO PDF COMBINED =========="
Add-DebugLog "PowerShell Version: $($PSVersionTable.PSVersion)"
Add-DebugLog "OS: $([System.Environment]::OSVersion)"
Add-DebugLog "Current User: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
Add-DebugLog "Working Directory: $(Get-Location)"
Add-DebugLog "Output File Name: $OutputFileName"

# ==================== EXCEL CHECK ====================
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
}
catch {
    Add-DebugLog "[ERROR] Failed to create Excel COM object"
    Add-DebugLog "  Exception: $_"
    exit 1
}

# Verify if Adobe Acrobat is available
try {
    $testAcrobat = New-Object -ComObject "AcroExch.PDDoc"
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($testAcrobat) | Out-Null
}
catch {
    Add-DebugLog "[ERROR] Failed to create Acrobat COM object"
    Add-DebugLog "  Exception: $_"
    Add-DebugLog ""
    Add-DebugLog "POSSIBLE REASONS:"
    Add-DebugLog "  1. Adobe Acrobat Pro is not installed"
    Add-DebugLog "  2. Only Adobe Reader is installed (Reader doesn't have PDDoc)"
    Add-DebugLog "  3. Acrobat installation is corrupted"
    Add-DebugLog "  4. Different Acrobat version than expected"
    
    # Check if any Adobe products exist in registry
    $adobeKeys = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" -ErrorAction SilentlyContinue |
        Where-Object { $_.GetValue('DisplayName') -like '*Adobe*' }
    
    if ($adobeKeys.Count -eq 0) {
      Add-DebugLog "  No Adobe products found in registry"
    }
    
    exit 1
}

# ==================== FILE DISCOVERY ====================
Add-DebugLog ""

$folder = Get-Location

$excelFiles = Get-ChildItem $folder -File -ErrorAction SilentlyContinue |
  Where-Object { $_.Extension -in ".xlsx", ".xls", ".xlsm" } |
  Sort-Object Name

if ($excelFiles.Count -eq 0) {
  Add-DebugLog "[ERROR] No Excel files found in $folder"
  Add-DebugLog ""
  Add-DebugLog "POSSIBLE REASONS:"
  Add-DebugLog "  1. No Excel files (.xlsx, .xls, .xlsm) in the directory"
  Add-DebugLog "  2. Files have different extensions"
  Add-DebugLog "  3. Permission issue reading directory"
  exit
}

# ==================== TEMP FOLDER SETUP ====================

# Create temporary folder for storing PDF files
$tempPdfFolder = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "PDFMerge_$(Get-Random)")

try {
    New-Item -ItemType Directory -Path $tempPdfFolder -Force -ErrorAction Stop | Out-Null
    
    # Verify folder exists
    if (-not (Test-Path $tempPdfFolder)) {
        Add-DebugLog "[ERROR] Temp folder was not created"
        throw "Temp folder creation failed"
    }
}
catch {
    Add-DebugLog "[ERROR] Failed to create temp folder"
    Add-DebugLog "  Exception: $_"
    Add-DebugLog ""
    Add-DebugLog "POSSIBLE REASONS:"
    Add-DebugLog "  1. No write permission to temp directory"
    Add-DebugLog "  2. Temp path doesn't exist or is corrupted"
    Add-DebugLog "  3. Antivirus blocking folder creation"
    exit 1
}

$pdfFiles = @()

# Convert each Excel file to PDF
foreach ($excelFile in $excelFiles) {
    Write-Host "Converting: $($excelFile.Name)"
    Add-DebugLog "Converting: $($excelFile.Name)"
    
    try {
        $workbook = $excel.Workbooks.Open($excelFile.FullName, $null, $false)
        
        $pdfFileName = [System.IO.Path]::ChangeExtension($excelFile.Name, ".pdf")
        $pdfPath = Join-Path $tempPdfFolder $pdfFileName
        
        # 0 = xlTypePDF
        $workbook.ExportAsFixedFormat(0, $pdfPath)
        
        # Verify PDF was created
        Start-Sleep -Milliseconds 200
        if (Test-Path $pdfPath) {
            $fileSize = (Get-Item $pdfPath).Length
            $pdfFiles += $pdfPath
        }
        else {
            Add-DebugLog "  [ERROR] PDF file was not created at $pdfPath"
            Add-DebugLog "  POSSIBLE REASONS:"
            Add-DebugLog "    1. Excel export failed silently"
            Add-DebugLog "    2. File permission issue in temp folder"
            Add-DebugLog "    3. Antivirus blocking file creation"
            Add-DebugLog "    4. Excel file is corrupted"
        }
        
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        
        # Small delay to ensure file is written
        Start-Sleep -Milliseconds 100
    }
    catch {
        Add-DebugLog "  [ERROR] Failed to convert $($excelFile.Name)"
        Add-DebugLog "  Exception: $_"
        Add-DebugLog "  POSSIBLE REASONS:"
        Add-DebugLog "    1. Excel file is locked by another process"
        Add-DebugLog "    2. Excel file is corrupted"
        Add-DebugLog "    3. Unsupported Excel format"
        Add-DebugLog "    4. Memory issue during conversion"
    }
}

if ($pdfFiles.Count -eq 0) {
    Add-DebugLog ""
    Add-DebugLog "[ERROR] No PDFs were created"
    Add-DebugLog "POSSIBLE REASONS:"
    Add-DebugLog "  1. All Excel files failed to convert"
    Add-DebugLog "  2. PDF temp folder has permission issues"
    Add-DebugLog "  3. Antivirus is blocking PDF creation"
    Add-DebugLog "  4. Disk space issue"
    Write-Host "No successfully converted PDFs, skipping"
    Add-DebugLog ""
    Remove-Item $tempPdfFolder -Recurse -Force -ErrorAction SilentlyContinue
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    exit
}

Add-DebugLog ""
Add-DebugLog "[OK] Successfully converted $($pdfFiles.Count) Excel files to PDF"

# Determine the output path dynamically
$outputFolder = (Get-Location).ProviderPath  # Use ProviderPath to get a clean path
$outputPdfPath = Join-Path -Path $outputFolder -ChildPath $OutputFileName

Add-DebugLog "Output file path: $outputPdfPath"

Add-DebugLog "Saving merged PDF..."
Add-DebugLog ""
Add-DebugLog "Output file path: $outputPdfPath"

try {
    # Create a new empty PDF document
    $pdDoc = New-Object -ComObject "AcroExch.PDDoc"
    
    $pdDoc.Create() | Out-Null
    
    Add-DebugLog "  Adding PDFs to merged document..."
    # Insert all PDFs in order
    for ($i = 0; $i -lt $pdfFiles.Count; $i++) {
        $pdfFileName = Split-Path $pdfFiles[$i] -Leaf
        
        try {
            # Open PDF to insert
            $insertDoc = New-Object -ComObject "AcroExch.PDDoc"
            $insertOpened = $insertDoc.Open($pdfFiles[$i])
            
            if ($insertOpened) {
                # Get number of pages to insert
                $numPagesToInsert = $insertDoc.GetNumPages()
                
                # Get current page count (insert after last page)
                $currentPageCount = $pdDoc.GetNumPages()
                $insertAfterPage = $currentPageCount - 1
                Add-DebugLog "    Inserting at position $insertAfterPage..."
                
                # Insert all pages
                # Parameters: insertAfterThisPage, sourcePDDoc, startPage, numPages, addBookmarks
                $insertResult = $pdDoc.InsertPages($insertAfterPage, $insertDoc, 0, $numPagesToInsert, 0) | Out-Null
                
                $insertDoc.Close() | Out-Null
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($insertDoc) | Out-Null
            }
            else {
                Add-DebugLog "    [ERROR] Could not open PDF: $($pdfFiles[$i])"
                Add-DebugLog "    POSSIBLE REASONS:"
                Add-DebugLog "      1. PDF file is corrupted"
                Add-DebugLog "      2. PDF is locked by another process"
                Add-DebugLog "      3. Insufficient permissions"
            }
        }
        catch {
            Add-DebugLog "    [ERROR] Failed to insert PDF"
            Add-DebugLog "    Exception: $_"
        }
    }
    
    Add-DebugLog ""
    Add-DebugLog "Saving merged PDF..."
    $totalPages = $pdDoc.GetNumPages()
    Add-DebugLog "  Total pages in merged document: $totalPages"
    
    # Save merged PDF
    $saveResult = $pdDoc.Save(1, $outputPdfPath) | Out-Null
    
    # Verify output file
    Start-Sleep -Milliseconds 500
    if (Test-Path $outputPdfPath) {
        $outputSize = (Get-Item $outputPdfPath).Length
        Write-Host "Complete: $OutputFileName (Total pages: $totalPages)"
    }
    else {
        Add-DebugLog "[ERROR] Output PDF was not created at $outputPdfPath"
        Add-DebugLog "POSSIBLE REASONS:"
        Add-DebugLog "  1. Acrobat Save failed silently"
        Add-DebugLog "  2. No write permission to output folder"
        Add-DebugLog "  3. Antivirus blocking file creation"
        Add-DebugLog "  4. Output path is invalid"
    }
    
    # Close document
    try {
        $pdDoc.Close() | Out-Null
    }
    catch {
        Add-DebugLog "[WARNING] Error closing PDF (file already saved)"
    }
    
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null
}
catch {
    Add-DebugLog "[ERROR] Failed to merge PDFs"
    Add-DebugLog "Exception: $_"
    Add-DebugLog ""
    Add-DebugLog "POSSIBLE REASONS:"
    Add-DebugLog "  1. Acrobat COM interface issue"
    Add-DebugLog "  2. PDFs are corrupted or incompatible"
    Add-DebugLog "  3. Memory issue"
    Add-DebugLog "  4. Acrobat version mismatch"
}
finally {
    # Clean up temporary files
    Add-DebugLog ""
    try {
        Add-DebugLog "Removing temporary folder: $tempPdfFolder"
        Remove-Item $tempPdfFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
    catch {
        Add-DebugLog "[WARNING] Could not delete temp folder (may delete manually)"
    }
}

$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Add-DebugLog ""
Add-DebugLog "========== SCRIPT FINISHED =========="

Write-Host ""
Write-Host "All done"
