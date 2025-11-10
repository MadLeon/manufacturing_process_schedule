<#
.SYNOPSIS
Scans G:\CANDU for PDF files and builds/updates a SQLite database.
.DESCRIPTION
Creates a database with drawings table if not exists, then performs high-speed scan.
#>

# Configuration
$gDrivePath = "G:\Bombardier"
$dbPath = ".\jobs.db"  # <-- Change this to your DB path
$batchSize = 500  # Files per transaction

# Load SQLite assembly
Add-Type -Path ".\System.Data.SQLite.dll"

# Create database and table if not exists
function Initialize-Database {
    param($dbPath)
    
    $conn = New-Object System.Data.SQLite.SQLiteConnection
    $conn.ConnectionString = "Data Source=$dbPath"
    $conn.Open()

    $cmd = $conn.CreateCommand()
    $cmd.CommandText = @"
CREATE TABLE IF NOT EXISTS drawings (
    drawing_name TEXT PRIMARY KEY,
    drawing_number TEXT DEFAULT NULL,
    file_location TEXT UNIQUE
);
CREATE INDEX IF NOT EXISTS idx_drawing_name ON drawings(drawing_name);
"@
    $cmd.ExecuteNonQuery()
    $conn.Close()
}

# High-performance scanner
function Scan-PDFs {
    param($gDrivePath, $dbPath, $batchSize)
    
    $conn = New-Object System.Data.SQLite.SQLiteConnection
    $conn.ConnectionString = "Data Source=$dbPath"
    $conn.Open()

    # Prepare batch insert
    $insertCmd = $conn.CreateCommand()
    $insertCmd.CommandText = "INSERT OR IGNORE INTO drawings (drawing_name, file_location) VALUES (@name, @path)"
    [void]$insertCmd.Parameters.Add("@name", [System.Data.SQLite.SQLiteType]::Text)
    [void]$insertCmd.Parameters.Add("@path", [System.Data.SQLite.SQLiteType]::Text)

    $counter = 0
    $transaction = $conn.BeginTransaction()

    try {
        Get-ChildItem -Path $gDrivePath -Recurse -Filter *.pdf | ForEach-Object {
            $insertCmd.Parameters["@name"].Value = $_.Name
            $insertCmd.Parameters["@path"].Value = $_.FullName
            [void]$insertCmd.ExecuteNonQuery()

            $counter++
            if ($counter % $batchSize -eq 0) {
                $transaction.Commit()
                Write-Progress -Activity "Scanning" -Status "$counter files processed"
                $transaction = $conn.BeginTransaction()
            }
        }
        $transaction.Commit()
    }
    catch {
        $transaction.Rollback()
        Write-Error "Error: $_"
    }
    finally {
        $conn.Close()
    }
    return $counter
}

# Main execution
Initialize-Database $dbPath
$processedCount = Scan-PDFs $gDrivePath $dbPath $batchSize
Write-Host "Scan complete! Processed $processedCount PDF files."