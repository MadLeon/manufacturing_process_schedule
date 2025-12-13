# One-time use script to create job_history table with unique_key column
# Run this script once to initialize the job_history table in the SQLite database

# $dbPath = "c:\Users\ee\manufacturing_process_schedule\oe\jobs.db"
$dbPath = "C:\Users\ee\jobs.db"

# Check if database exists
if (-not (Test-Path $dbPath)) {
    Write-Host "Error: Database file not found at $dbPath" -ForegroundColor Red
    exit
}

# SQLite3 command to create the job_history table with unique_key column
$createTableSQL = @"
CREATE TABLE IF NOT EXISTS job_history (
    job_id INTEGER PRIMARY KEY AUTOINCREMENT,
    oe_number TEXT,
    job_number TEXT,
    customer_name TEXT,
    job_quantity TEXT,
    part_number TEXT,
    revision TEXT,
    customer_contact TEXT,
    drawing_release TEXT,
    line_number TEXT,
    part_description TEXT,
    unit_price TEXT,
    po_number TEXT,
    packing_slip TEXT,
    packing_quantity TEXT,
    invoice_number TEXT,
    delivery_required_date TEXT,
    delivery_shipped_date TEXT,
    unique_key TEXT UNIQUE,
    create_timestamp TEXT,
    last_modified TEXT,
    completed_timestamp TEXT DEFAULT (datetime('now','localtime'))
);
"@

# Save SQL to temp file
$tempSqlFile = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.sql'
$createTableSQL | Out-File -FilePath $tempSqlFile -Encoding UTF8

try {
    Write-Host "Creating job_history table with unique_key column..." -ForegroundColor Cyan
    
    # Use sqlite3 command line tool with piped input
    $createTableSQL | sqlite3 $dbPath
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "SUCCESS: job_history table created successfully!" -ForegroundColor Green
        Write-Host "Table includes columns:" -ForegroundColor Green
        Write-Host "  - job_id (PRIMARY KEY)" -ForegroundColor Green
        Write-Host "  - Standard job fields (oe_number, job_number, customer_name, etc.)" -ForegroundColor Green
        Write-Host "  - unique_key (TEXT UNIQUE) - for uniqueness constraint" -ForegroundColor Green
        Write-Host "  - Timestamps (create_timestamp, last_modified, completed_timestamp)" -ForegroundColor Green
    } else {
        Write-Host "ERROR: Failed to create job_history table. Exit code: $LASTEXITCODE" -ForegroundColor Red
    }
} catch {
    Write-Host "ERROR: Exception occurred: $_" -ForegroundColor Red
} finally {
    # Clean up temp file
    if (Test-Path $tempSqlFile) {
        Remove-Item $tempSqlFile -Force
    }
}
