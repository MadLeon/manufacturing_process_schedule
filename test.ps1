# -----------------------------
# Test Script for System.Data.SQLite.dll
# -----------------------------

# Path to your DLL
$sqliteDll = "C:\Users\ee\manufacturing_process_schedule\System.Data.SQLite.dll"

# Load the DLL
Add-Type -Path $sqliteDll

# Path for test SQLite database
$dbPath = "$PSScriptRoot\test_db.sqlite"

# Delete existing DB if exists
if (Test-Path $dbPath) { Remove-Item $dbPath }

# Create a new SQLite connection
$conn = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$dbPath;Version=3;")
$conn.Open()

# Create a test table
$cmd = $conn.CreateCommand()
$cmd.CommandText = @"
CREATE TABLE IF NOT EXISTS test_table (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT,
    age INTEGER
);
"@
$cmd.ExecuteNonQuery()

# Insert a sample row
$cmd.CommandText = "INSERT INTO test_table (name, age) VALUES ('Leon', 25);"
$cmd.ExecuteNonQuery()

# Query the table
$cmd.CommandText = "SELECT * FROM test_table;"
$reader = $cmd.ExecuteReader()

Write-Output "Contents of test_table:"
while ($reader.Read()) {
    Write-Output "ID: $($reader["id"]), Name: $($reader["name"]), Age: $($reader["age"])"
}

# Close the connection
$conn.Close()
