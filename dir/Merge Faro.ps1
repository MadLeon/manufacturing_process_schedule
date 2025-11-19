# PowerShell script to sort the first CSV file based on column A, copy the col a to the rest of files, and merge the files

# Get the directory of the script
$scriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Path -Parent

# Get all CSV files in the directory
$csvFiles = Get-ChildItem -Path $scriptDirectory -Filter "*.csv" | Sort-Object {$_.Name -replace '^.*-(\d+)\.csv$', '$1' -as [int]}

# Get the first CSV file (assuming it ends with -1.csv)
$firstFile = $csvFiles | Where-Object {$_.Name -like "*-1.csv"} | Select-Object -First 1

# Check if the first file exists
if ($firstFile) {
    $firstFilePath = $firstFile.FullName

    # Import the CSV data from the first file
    $firstFileData = Import-Csv -Path $firstFilePath

    # Extract the 'Part' column from the first file
    $partValues = $firstFileData | Select-Object -ExpandProperty Part

    # Loop through all CSV files to copy 'Part' values
    foreach ($csvFile in $csvFiles) {
        $csvFilePath = $csvFile.FullName

        # Import the CSV data
        $csvData = Import-Csv -Path $csvFilePath

        # Add the 'Part' values to the current CSV data
        for ($i = 0; $i -lt $csvData.Count; $i++) {
            if ($i -lt $partValues.Count) {
                $csvData[$i].Part = $partValues[$i]
            }
        }

        # Export the modified data back to the same CSV file, without type information
        $csvData | Export-Csv -Path $csvFilePath -NoTypeInformation
    }

    # Define the column order
    $columnOrder = "Part","Feature","Property","Actual","Nominal","Deviation","Low Tol","Up Tol","Out of Tol"

    # Create the merged data array
    $mergedData = @()

    # Loop through the CSV files to merge
    foreach ($csvFile in $csvFiles) {
        $csvFilePath = $csvFile.FullName
        $csvFileName = $csvFile.Name

        # Import CSV data
        $csvData = Import-Csv -Path $csvFilePath

        # Sort the CSV data based on the 'Part' column
        $sortedData = $csvData | Sort-Object -Property @{Expression={
            if($_.Part -match "^([\d\.]+).*"){
                [decimal]$matches[1]
            }
            elseif($_.Part -match "^([\d]+).*"){
                [decimal]$matches[1]
            }
            else{
                999999999
            }
        }}

        # Extract the number from the filename
        if ($csvFileName -match "-(\d+)\.csv") {
            $number = $Matches[1]
        } else {
            $number = "Unknown"
        }

        # Add the separator row
        $separator = [PSCustomObject]@{}
        foreach ($col in $columnOrder) {
            if ($col -eq "Part") {
                $separator | Add-Member -MemberType NoteProperty -Name $col -Value "# $number"
            } else {
                $separator | Add-Member -MemberType NoteProperty -Name $col -Value ""
            }
        }
        $mergedData += $separator

        # Add the CSV data to the merged data
        foreach ($row in $sortedData) {
            $newRow = [PSCustomObject]@{}
            foreach ($col in $columnOrder) {
                $newRow | Add-Member -MemberType NoteProperty -Name $col -Value $($row."$col")
            }
            $mergedData += $newRow
        }
    }

    # Export the merged data to a new CSV file
    $mergedData | Export-Csv -Path "$scriptDirectory\merged.csv" -NoTypeInformation

    Write-Host "Successfully copied 'Part' column, sorted each file based on column A, and merged the files."
} else {
    Write-Host "Error: Could not find the first CSV file (ending with '-1.csv')."
}