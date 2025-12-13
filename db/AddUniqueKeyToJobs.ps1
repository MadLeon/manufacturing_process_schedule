# 为 jobs 表添加唯一键列并创建唯一索引
# 执行操作：
# 1. 向 jobs 表添加 unique_key 列
# 2. 根据 Job_Number, Line_Number 生成唯一键
# 3. 创建唯一索引 idx_unique_key

[System.IO.File]::WriteAllText($PSScriptRoot + "\encoding_test.txt", "test", [System.Text.Encoding]::UTF8)

# Database configuration
# $dbPath = "C:\Users\ee\manufacturing_process_schedule\oe\jobs.db"
$dbPath = "C:\Users\ee\jobs.db"

# Verify database file exists
if (-not (Test-Path $dbPath)) {
    Write-Error "Database file not found: $dbPath"
    exit 1
}

Write-Host "Connecting to database: $dbPath" -ForegroundColor Cyan

# SQL commands
$sqlStatements = @(
    "ALTER TABLE jobs ADD COLUMN unique_key TEXT;",
    "UPDATE jobs SET unique_key = Job_Number || '|' || COALESCE(NULLIF(Line_Number, ''), '1');",
    "CREATE UNIQUE INDEX idx_unique_key ON jobs(unique_key);"
)

try {
    # Execute each SQL statement
    foreach ($stmt in $sqlStatements) {
        Write-Host "Executing: $stmt" -ForegroundColor Yellow
        sqlite3 $dbPath $stmt
    }
    
    Write-Host "`nOperation completed successfully!" -ForegroundColor Green
    
    # Verify results
    Write-Host "`nVerification results:" -ForegroundColor Yellow
    $verifyCmd = "SELECT COUNT(*) as total_rows, COUNT(unique_key) as unique_key_count, COUNT(DISTINCT unique_key) as distinct_keys FROM jobs;"
    sqlite3 $dbPath $verifyCmd
    
    Write-Host "`nIndex verification:" -ForegroundColor Yellow
    $indexCmd = "SELECT name, sql FROM sqlite_master WHERE type='index' AND name='idx_unique_key';"
    sqlite3 $dbPath $indexCmd
    
}
catch {
    Write-Error "Failed to execute SQL: $_"
    exit 1
}

Write-Host "`nAll operations completed successfully!" -ForegroundColor Green
