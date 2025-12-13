<#
.SYNOPSIS
Synchronizes jobs database with Order Entry Log's DELIVERY SCHEDULE sheet
.DESCRIPTION
- Assumes jobs table and job_history table exist by default
- Uses unique_key instead of job_number as the key for dictionary operations
- Records not in delivery schedule are moved to job_history table instead of being deleted
- Ensures consistency between jobs.db and DELIVERY SCHEDULE sheet
#>

# ==================== Configuration ====================
# $dbPath = "C:\Users\ee\manufacturing_process_schedule\oe\jobs.db"
# $oeEntryPath = "C:\Users\ee\manufacturing_process_schedule\oe\Order Entry Log.xlsm"
$dbPath = "C:\Users\ee\jobs.db"
$oeEntryPath = "C:\Users\ee\Order Entry Log.xlsm"
$worksheetName = "DELIVERY SCHEDULE"
$firstDataRow = 4  # Data starts from row 4 (rows 1-3 contain header info)

# ==================== Main Sync Function ====================
function Sync-JobsDBAndOrderEntryLog {
    param()

    Write-Host "Starting database synchronization..." -ForegroundColor Cyan

    # Step 1: Check if database exists
    if (-not (Test-Path $dbPath)) {
        Write-Host "Error: Database file not found at $dbPath" -ForegroundColor Red
        exit
    }

    # Step 2: Open Excel workbook in background (read-only)
    Write-Host "Opening Order Entry Log..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    try {
        $workbook = $excel.Workbooks.Open($oeEntryPath, $false, $true)  # Read-only mode
        $worksheet = $workbook.Sheets($worksheetName)
        
        # Find the last row with actual data more carefully
        $lastRow = $worksheet.Cells($worksheet.Rows.Count, 2).End(-4162).row  # Find last non-empty cell in column B (Job_Number)
        if ($lastRow -lt $firstDataRow) {
            $lastRow = $worksheet.UsedRange.Rows.Count
        }
        Write-Host "Total rows to scan: $lastRow (starting from row $firstDataRow)" -ForegroundColor Cyan
        
        # Step 3: Build dictionary from Order Entry Log using unique_key generated at runtime
        # Read all data into memory at once (much faster than cell-by-cell access)
        Write-Host "Reading all data into memory..." -ForegroundColor Yellow
        
        # Step 3: Load all data into memory at once for performance
        Write-Host "Loading data from Excel..." -ForegroundColor Yellow
        
        # Build range string A4:Q548 format for proper COM access
        $rangeStr = "A$firstDataRow" + ":Q$lastRow"
        $dataArray = $worksheet.Range($rangeStr).Value2
        
        # Helper function to safely convert cell values to string
        function Safe-ToString {
            param($value)
            if ($null -eq $value) { return "" }
            return [string]$value
        }
        
        # Helper function to convert Excel serial date numbers to date string
        function SafeToDateString {
            param($value)
            if ($null -eq $value) { return "" }
            
            # Check if it's a number (Excel date serial)
            if ($value -is [double] -or $value -is [int]) {
                try {
                    # Excel date serial: days since 1/1/1900
                    $excelDate = [DateTime]::FromOADate($value)
                    return $excelDate.ToString("d-MMM-yy")  # Format: 7-Mar-24
                }
                catch {
                    return [string]$value
                }
            }
            # If already a string, return as-is
            return [string]$value
        }
        
        Write-Host "Building dictionary from Order Entry Log..." -ForegroundColor Yellow
        $entryDict = @{}
        
        # Determine if we have data
        if ($null -eq $dataArray) {
            Write-Host "No data found in range" -ForegroundColor Yellow
        }
        elseif ($dataArray -is [object[,]]) {
            # 2D array (1-based indexing from COM)
            $rowCount = $dataArray.GetLength(0)
            
            for ($i = 1; $i -le $rowCount; $i++) {
                if ($i % 100 -eq 0) {
                    Write-Host "  Processed $i rows..." -ForegroundColor Gray
                }
                
                try {
                    # COM arrays are 1-based: [row, col]
                    # Col 2 = Job_Number, Col 9 = Line_Number
                    $jobNumber = (Safe-ToString $dataArray[$i, 2]).Trim()
                    $lineNumber = (Safe-ToString $dataArray[$i, 9]).Trim()
                    
                    # If line number is empty, default to 1
                    if ([string]::IsNullOrWhiteSpace($lineNumber)) {
                        $lineNumber = "1"
                    }
                    
                    if (-not [string]::IsNullOrWhiteSpace($jobNumber)) {
                        $uniqueKey = "$jobNumber|$lineNumber"
                        $entryDict[$uniqueKey] = $firstDataRow + $i - 1
                    }
                }
                catch {
                    # Silent continue on errors
                    continue
                }
            }
        }
        else {
            # Single row - treat as array-like object
            $jobNumber = (Safe-ToString $dataArray[2]).Trim()
            $lineNumber = (Safe-ToString $dataArray[9]).Trim()
            
            # If line number is empty, default to 1
            if ([string]::IsNullOrWhiteSpace($lineNumber)) {
                $lineNumber = "1"
            }
            
            if (-not [string]::IsNullOrWhiteSpace($jobNumber)) {
                $uniqueKey = "$jobNumber|$lineNumber"
                $entryDict[$uniqueKey] = $firstDataRow
            }
        }
        
        Write-Host "Unique keys in Order Entry Log: $($entryDict.Count)" -ForegroundColor Green
        
        # Step 4: Build dictionary from jobs table using unique_key
        Write-Host "Loading jobs from database..." -ForegroundColor Yellow
        $dbDict = @{}
        
        try {
            $result = sqlite3 $dbPath "SELECT unique_key FROM jobs WHERE unique_key IS NOT NULL AND unique_key != '';" 2>&1
            if ($result) {
                foreach ($key in $result) {
                    if (-not [string]::IsNullOrWhiteSpace($key)) {
                        $dbDict[$key.Trim()] = 1
                    }
                }
            }
        }
        catch {
            Write-Host "Warning: Could not query jobs table. Continuing..." -ForegroundColor Yellow
        }
        
        Write-Host "Unique keys in jobs table: $($dbDict.Count)" -ForegroundColor Green
        
        # Step 5: Synchronization
        # A. Add records from entry to database (not in jobs table)
        Write-Host "Checking for new records to add..." -ForegroundColor Yellow
        
        $addedCount = 0
        $entryKeysCount = $entryDict.Keys.Count
        $entryIndex = 0
        
        foreach ($key in $entryDict.Keys) {
            $entryIndex++
            if ($entryIndex % 50 -eq 0) {
                Write-Host "  Processing add operations: $entryIndex / $entryKeysCount" -ForegroundColor Gray
            }
            
            if (-not $dbDict.ContainsKey($key)) {
                try {
                    $rowNum = $entryDict[$key]
                    $arrayIndex = $rowNum - $firstDataRow + 1  # Convert Excel row number to 1-based array index
                    
                    # Get values from the data array (already in memory)
                    $oeNumber = (Safe-ToString $dataArray[$arrayIndex, 1]).Trim()
                    $jobNumber = (Safe-ToString $dataArray[$arrayIndex, 2]).Trim()
                    $customerName = (Safe-ToString $dataArray[$arrayIndex, 3]).Trim()
                    $jobQuantity = (Safe-ToString $dataArray[$arrayIndex, 4]).Trim()
                    $partNumber = (Safe-ToString $dataArray[$arrayIndex, 5]).Trim()
                    $revision = (Safe-ToString $dataArray[$arrayIndex, 6]).Trim()
                    $customerContact = (Safe-ToString $dataArray[$arrayIndex, 7]).Trim()
                    $drawingRelease = (SafeToDateString $dataArray[$arrayIndex, 8]).Trim()
                    $lineNumber = (Safe-ToString $dataArray[$arrayIndex, 9]).Trim()
                    $partDescription = (Safe-ToString $dataArray[$arrayIndex, 10]).Trim()
                    $unitPrice = (Safe-ToString $dataArray[$arrayIndex, 11]).Trim()
                    $poNumber = (Safe-ToString $dataArray[$arrayIndex, 12]).Trim()
                    $packingSlip = (Safe-ToString $dataArray[$arrayIndex, 13]).Trim()
                    $packingQuantity = (Safe-ToString $dataArray[$arrayIndex, 14]).Trim()
                    $invoiceNumber = (Safe-ToString $dataArray[$arrayIndex, 15]).Trim()
                    $deliveryRequiredDate = (SafeToDateString $dataArray[$arrayIndex, 16]).Trim()
                    $deliveryShippedDate = (SafeToDateString $dataArray[$arrayIndex, 17]).Trim()
                    
                    # Escape single quotes in string values
                    $oeNumber = (Escape-SqlString $oeNumber)
                    $jobNumber = (Escape-SqlString $jobNumber)
                    $customerName = (Escape-SqlString $customerName)
                    $jobQuantity = (Escape-SqlString $jobQuantity)
                    $partNumber = (Escape-SqlString $partNumber)
                    $revision = (Escape-SqlString $revision)
                    $customerContact = (Escape-SqlString $customerContact)
                    $drawingRelease = (Escape-SqlString $drawingRelease)
                    $lineNumber = (Escape-SqlString $lineNumber)
                    $partDescription = (Escape-SqlString $partDescription)
                    $unitPrice = (Escape-SqlString $unitPrice)
                    $poNumber = (Escape-SqlString $poNumber)
                    $packingSlip = (Escape-SqlString $packingSlip)
                    $packingQuantity = (Escape-SqlString $packingQuantity)
                    $invoiceNumber = (Escape-SqlString $invoiceNumber)
                    $deliveryRequiredDate = (Escape-SqlString $deliveryRequiredDate)
                    $deliveryShippedDate = (Escape-SqlString $deliveryShippedDate)
                    
                    $insertSQL = @"
INSERT INTO jobs (oe_number, job_number, customer_name, job_quantity, part_number, revision, 
                  customer_contact, drawing_release, line_number, part_description, unit_price, 
                  po_number, packing_slip, packing_quantity, invoice_number, delivery_required_date, 
                  delivery_shipped_date, unique_key, last_modified)
VALUES (
    '$oeNumber',
    '$jobNumber',
    '$customerName',
    '$jobQuantity',
    '$partNumber',
    '$revision',
    '$customerContact',
    '$drawingRelease',
    '$lineNumber',
    '$partDescription',
    '$unitPrice',
    '$poNumber',
    '$packingSlip',
    '$packingQuantity',
    '$invoiceNumber',
    '$deliveryRequiredDate',
    '$deliveryShippedDate',
    '$key',
    datetime('now','localtime')
);
"@
                    
                    sqlite3 $dbPath $insertSQL 2>&1 | Out-Null
                    $addedCount++
                }
                catch {
                    Write-Host "Warning: Failed to insert record with unique_key=$key : $_" -ForegroundColor Yellow
                }
            }
        }
        Write-Host "Records added to jobs table: $addedCount" -ForegroundColor Green
        
        # B. Move records from jobs to job_history (in jobs but not in entry)
        Write-Host "Checking for records to move to history..." -ForegroundColor Yellow
        
        $movedCount = 0
        $dbKeysCount = $dbDict.Keys.Count
        $dbIndex = 0
        
        foreach ($key in $dbDict.Keys) {
            $dbIndex++
            if ($dbIndex % 50 -eq 0) {
                Write-Host "  Processing move operations: $dbIndex / $dbKeysCount" -ForegroundColor Gray
            }
            
            if (-not $entryDict.ContainsKey($key)) {
                try {
                    # Step 1: Get the record from jobs table
                    $selectSQL = "SELECT * FROM jobs WHERE unique_key = '$key' LIMIT 1;"
                    $jobRecord = sqlite3 $dbPath $selectSQL 2>&1
                    
                    if ($jobRecord) {
                        # Step 2: Insert into job_history with completed_timestamp
                        $insertHistorySQL = @"
INSERT INTO job_history (
    oe_number, job_number, customer_name, job_quantity, part_number, revision,
    customer_contact, drawing_release, line_number, part_description, unit_price,
    po_number, packing_slip, packing_quantity, invoice_number, delivery_required_date,
    delivery_shipped_date, unique_key, create_timestamp, last_modified, completed_timestamp
)
SELECT 
    oe_number, job_number, customer_name, job_quantity, part_number, revision,
    customer_contact, drawing_release, line_number, part_description, unit_price,
    po_number, packing_slip, packing_quantity, invoice_number, delivery_required_date,
    delivery_shipped_date, unique_key, create_timestamp, last_modified, datetime('now','localtime')
FROM jobs WHERE unique_key = '$key';
"@
                        sqlite3 $dbPath $insertHistorySQL 2>&1 | Out-Null
                        
                        # Step 3: Delete from jobs table
                        $deleteSQL = "DELETE FROM jobs WHERE unique_key = '$key';"
                        sqlite3 $dbPath $deleteSQL 2>&1 | Out-Null
                        
                        $movedCount++
                        Write-Host "Moved to job_history: unique_key=$key" -ForegroundColor Cyan
                    }
                }
                catch {
                    Write-Host "Warning: Failed to move record with unique_key=$key : $_" -ForegroundColor Yellow
                }
            }
        }
        Write-Host "Records moved to job_history: $movedCount" -ForegroundColor Green
        
        # Step 6: Summary
        Write-Host ""
        Write-Host "========== Synchronization Complete ==========" -ForegroundColor Cyan
        Write-Host "Records added to jobs table: $addedCount" -ForegroundColor Green
        Write-Host "Records moved to job_history: $movedCount" -ForegroundColor Green
        Write-Host "=============================================" -ForegroundColor Cyan
        
    }
    finally {
        # Clean up Excel objects
        if ($workbook) {
            $workbook.Close($false)
        }
        if ($excel) {
            $excel.Quit()
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
    }
}

# ==================== Utility Function: Escape SQL String ====================
function Escape-SqlString {
    param([string]$str)
    if ([string]::IsNullOrEmpty($str)) {
        return ""
    }
    return $str.Replace("'", "''")
}

# ==================== Execute Main Function ====================
Sync-JobsDBAndOrderEntryLog
