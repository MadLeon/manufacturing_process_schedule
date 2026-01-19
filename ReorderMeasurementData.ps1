# 重新排列测量数据的No.列顺序
# 按照指定顺序：2,1,3,4,13,8,10,7,11,15,16,14,6,12,9

$sourceDir = "C:\Users\ee\manufacturing_process_schedule\dir"
$files = @(
    "rt-87900-PN-R003(1-4).csv",
    "rt-87900-PN-R003(5-8)NEW.csv",
    "rt-87900-PN-R003(9-10).csv"
)

# 定义新的排列顺序
$newOrder = @(2, 1, 3, 4, 13, 8, 10, 7, 11, 15.1, 15.2, 16, 14, 6, 12, 9)

function Reorder-MeasurementData {
    param(
        [string]$filePath,
        [array]$newOrder
    )
    
    # 读取文件内容
    $content = Get-Content -Path $filePath -Raw
    $lines = $content -split "`r`n"
    
    $result = @()
    $i = 0
    
    while ($i -lt $lines.Count) {
        $line = $lines[$i]
        
        # 检查是否是"<No. X area>"标记
        if ($line -like "*<No.*area>*") {
            $result += $line
            $i++
            
            # 添加Serial Counter和Overall result行
            if ($i -lt $lines.Count) { $result += $lines[$i]; $i++ }
            if ($i -lt $lines.Count) { $result += $lines[$i]; $i++ }
            
            # 添加"Measurement results"标题
            if ($i -lt $lines.Count) { $result += $lines[$i]; $i++ }
            
            # 添加表头
            if ($i -lt $lines.Count) { $result += $lines[$i]; $i++ }
            
            # 现在处理测量数据行
            $measurementLines = @()
            while ($i -lt $lines.Count -and $lines[$i] -ne "" -and $lines[$i] -notlike "*area*") {
                if ($lines[$i] -match "^[0-9]") {
                    $measurementLines += $lines[$i]
                    $i++
                }
                elseif ($lines[$i] -match "^[0-9]+\.[0-9]+") {
                    $measurementLines += $lines[$i]
                    $i++
                }
                else {
                    # 空行或其他标记
                    if ($lines[$i] -ne "") {
                        break
                    }
                    $i++
                }
            }
            
            # 按新顺序重新排列测量数据
            $reorderedLines = @()
            foreach ($newNum in $newOrder) {
                $found = $false
                foreach ($mLine in $measurementLines) {
                    $parts = $mLine -split ","
                    if ($parts[0] -eq [string]$newNum) {
                        $reorderedLines += $mLine
                        $found = $true
                        break
                    }
                }
                if (-not $found) {
                    Write-Host "警告：未找到No. $newNum 的数据行"
                }
            }
            
            # 添加重新排列后的行
            $result += $reorderedLines
            
            # 添加空行（如果有）
            if ($i -lt $lines.Count -and $lines[$i] -eq "") {
                $result += $lines[$i]
                $i++
            }
        }
        else {
            $result += $line
            $i++
        }
    }
    
    # 写回文件
    $output = $result -join "`r`n"
    Set-Content -Path $filePath -Value $output -NoNewline
    Write-Host "已处理：$filePath"
}

# 处理所有文件
foreach ($file in $files) {
    $filePath = Join-Path -Path $sourceDir -ChildPath $file
    if (Test-Path $filePath) {
        Write-Host "处理文件：$file"
        Reorder-MeasurementData -filePath $filePath -newOrder $newOrder
    }
    else {
        Write-Host "文件不存在：$filePath"
    }
}

Write-Host "所有文件处理完成！"
