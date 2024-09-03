# Define the path for the output file
$outputFile = "C:\Temp\ExcelCOMTest.xlsx"
$outputDir = [System.IO.Path]::GetDirectoryName($outputFile)

# Ensure the directory exists
if (-not (Test-Path $outputDir)) {
    Write-Host "Directory does not exist. Creating directory: $outputDir" -ForegroundColor Yellow
    try {
        New-Item -Path $outputDir -ItemType Directory | Out-Null
        Write-Host "Directory created successfully." -ForegroundColor Green
    } catch {
        Write-Error "Failed to create directory: $_"
        exit
    }
}

# Test Excel COM Object
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false  # Set to false for non-interactive mode

    # Add a new workbook
    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Sheets.Item(1)

    # Write a value to a cell
    $worksheet.Cells.Item(1, 1) = "Hello, Excel!"

    # Save the workbook to the specified location
    try {
        $workbook.SaveAs($outputFile, 51)  # 51 is the code for .xlsx
        Write-Host "Test file saved to $outputFile." -ForegroundColor Green
    } catch {
        Write-Error "Failed to save the test Excel file: $_"
    } finally {
        # Close the workbook and quit Excel
        $workbook.Close($false)
        $excel.Quit()

        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
} catch {
    Write-Error "Failed to initialize or operate with Excel COM object: $_"
}

# Check if the file was created
if (Test-Path $outputFile) {
    Write-Host "Test file created successfully!" -ForegroundColor Green
} else {
    Write-Host "Test file was not created. Something went wrong." -ForegroundColor Red
}
