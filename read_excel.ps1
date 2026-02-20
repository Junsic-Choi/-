[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$excelFile = Get-ChildItem "C:\Users\i0215099\Desktop\ANTI_TEST\*.xlsx" | Select-Object -First 1
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelFile.FullName)
foreach ($worksheet in $workbook.Worksheets) {
    Write-Host "=== Sheet: $($worksheet.Name) ==="
    $range = $worksheet.UsedRange
    $rows = $range.Rows.Count
    $cols = $range.Columns.Count
    Write-Host "Rows: $rows, Cols: $cols"
    for ($r=1; $r -le [math]::Min($rows, 20); $r++) {
        $rowData = ""
        for ($c=1; $c -le [math]::Min($cols, 10); $c++) {
            $val = $range.Cells.Item($r, $c).Text
            if ($null -eq $val) { $val = "" }
            # remove newlines within cells to avoid messy output
            $val = $val -replace "`n"," " -replace "`r",""
            $rowData += $val + "|TAB|"
        }
        Write-Host $rowData
    }
}
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
