# Test Excel COM Object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true  # Set to true to observe Excel in action

# Add a new workbook
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Sheets.Item(1)

# Write a value to a cell
$worksheet.Cells.Item(1, 1) = "Hello, Excel!"

# Save the workbook to a known location
$outputFile = "C:\Temp\ExcelCOMTest.xlsx"
$outputFilePath = Resolve-Path $outputFile

try {
    $workbook.SaveAs($outputFilePath.Path, 51)  # 51 is the code for .xlsx
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

# Check if the file was created
if (Test-Path $outputFile) {
    Write-Host "Test file created successfully!" -ForegroundColor Green
} else {
    Write-Host "Test file was not created. Something went wrong." -ForegroundColor Red
}
