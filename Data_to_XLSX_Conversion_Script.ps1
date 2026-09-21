# 1. Define the folder where the .txt files are stored and the output files
$InputFolder = "Input_Folder_Path" 
$CsvFile = "CSV_File_Destination_Path"
$ExcelFile = "XLSX_File_Destination_Path"

# 2. Check for existing data to avoid duplicates
$existingSerials = @()
if (Test-Path $CsvFile) {
    $existingData = Import-Csv $CsvFile
    # Store existing serial numbers to compare against
    $existingSerials = $existingData | Select-Object -ExpandProperty "SerialNumber"
}

$newEntries = @()

# 3. Grab all text files in the folder
$files = Get-ChildItem -Path $InputFolder -Filter *.txt

foreach ($file in $files) {
    $content = Get-Content $file.FullName
    $record = @{}
    
    # 4. Parse the text file line by line
    foreach ($line in $content) {
        # Look for the pattern "Key: Value" with or without spaces before the colon
        if ($line -match "^(.*?)\s*:\s*(.*)$") {
            # Remove spaces from the key name
            $key = $matches[1].Trim() -replace ' ', ''
            $value = $matches[2].Trim()
            $record[$key] = $value
        }
    }

    $serial = $record["SerialNumber"]
    
    # 5. Check if the Serial Number is already in the Master Sheet
    if ($serial -and $existingSerials -notcontains $serial) {
        $newRow = [PSCustomObject]@{
            Hostname     = $record["Hostname"]
            Username     = $record["Username"]
            SerialNumber = $serial
            ModelNumber  = $record["ModelNumber"]
            OSVersion    = $record["OSVersion"]
            CPU          = $record["CPU"]
            MemoryRAM    = $record["Memory(RAM)"]
            Storage      = $record["Storage"]
        }
        $newEntries += $newRow
    }
}

# 6. Append new entries to the master CSV file
if ($newEntries.Count -gt 0) {
    $newEntries | Export-Csv -Path $CsvFile -NoTypeInformation -Append
    Write-Host "Added $($newEntries.Count) new records to the CSV." -ForegroundColor Green
} else {
    Write-Host "No new records found. The CSV is already up to date." -ForegroundColor Yellow
}

# 7. Convert the CSV to an Excel (.xlsx) file and apply all formatting
if (Test-Path $CsvFile) {
    Write-Host "Generating the updated Excel file..." -ForegroundColor Cyan

    # Delete the old Excel file if it exists to ensure a clean overwrite
    if (Test-Path $ExcelFile) {
        Remove-Item $ExcelFile -Force
    }

    # Open Excel in the background and save as .xlsx
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false # Suppresses any overwrite prompts

        $workbook = $excel.Workbooks.Open($CsvFile)
        $worksheet = $workbook.Worksheets.Item(1)
        
        # Bold the first row (headers)
        $worksheet.UsedRange.Rows.Item(1).Font.Bold = $true

        # Center align all cells vertically and horizontally
        $xlCenter = -4108
        $worksheet.UsedRange.HorizontalAlignment = $xlCenter
        $worksheet.UsedRange.VerticalAlignment = $xlCenter

        # --- NEW: Apply "All Borders" to filled cells ---
        $xlContinuous = 1
        $xlThin = 2
        $worksheet.UsedRange.Borders.LineStyle = $xlContinuous
        $worksheet.UsedRange.Borders.Weight = $xlThin

        # AutoFit all columns containing data
        $worksheet.UsedRange.Columns.AutoFit() | Out-Null
        
        # Save as file format 51, which is xlOpenXMLWorkbook (.xlsx)
        $workbook.SaveAs($ExcelFile, 51)
        
        # Close and cleanly release the COM objects from system memory
        $workbook.Close()
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        Write-Host "Success! Your updated Excel sheet is ready at: $ExcelFile" -ForegroundColor Green
    } catch {
        Write-Host "An error occurred while communicating with Excel:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Yellow
        
        # Emergency cleanup just in case it failed mid-process
        if ($excel) { 
            $excel.Quit() 
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
    }
} else {
    Write-Host "Waiting for data. Please ensure there are employee .txt files in $InputFolder" -ForegroundColor Yellow
}