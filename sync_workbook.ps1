param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$JsonPath,
    [int]$MaxRows = 0
)

$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null
$oldCalculation = $null
try {
    $decoded = Get-Content -LiteralPath $JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $rows = if ($null -eq $decoded) { @() } else { @($decoded) }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $false)
    if ($workbook.ReadOnly) { throw 'Workbook is read-only. Close the workbook and try again.' }
    $oldCalculation = $excel.Calculation
    $excel.Calculation = -4135
    $sheet = $workbook.Worksheets.Item(2)
    $quoteSheet = $workbook.Worksheets.Item(1)
    $lastUsedRow = $sheet.Cells.Item($sheet.Rows.Count, 1).End(-4162).Row
    $clearEndRow = [Math]::Max(6, $lastUsedRow)
    $sheet.Range("A6:K$clearEndRow").ClearContents()
    $count = if ($MaxRows -gt 0) { [Math]::Min($rows.Count, $MaxRows) } else { $rows.Count }
    if ($count -gt 1048570) { throw '已选型号超过 Excel 工作表容量。' }
    if ($count -gt 0) {
        $currencyFormat = ([string][char]0x00A5) + '#,##0.00'
        $sheet.Range("E6:F$($count + 5)").NumberFormat = $currencyFormat
        $sheet.Range("I6:I$($count + 5)").NumberFormat = 'yyyy-mm-dd hh:mm'
        $values = New-Object 'object[,]' $count, 11
        for ($r = 0; $r -lt $count; $r++) {
            for ($c = 0; $c -lt 11; $c++) {
                $value = $rows[$r].values[$c]
                if ($value -is [decimal] -or $value -is [int64] -or $value -is [int32]) {
                    $value = [double]$value
                }
                if ($null -ne $value) { $values[$r,$c] = $value }
            }
        }
        $target = $sheet.Range('A6').Resize($count, 11)
        $target.Value2 = $values
    }

    # 下拉来源和报价公式随本次实际数据行数扩展，既不限制同步数量，
    # 也避免整列检索拖慢 Excel 计算。
    $dataEndRow = [Math]::Max(6, $count + 5)
    try { $workbook.Names.Item('LimoProductList').Delete() } catch {}
    $escapedSheetName = $sheet.Name.Replace("'", "''")
    $refersTo = "='" + $escapedSheetName + "'!`$A`$6:`$A`$$dataEndRow"
    [void]$workbook.Names.Add('LimoProductList', $refersTo)
    $validationRange = $quoteSheet.Range('C11:C25')
    $validationRange.Validation.Delete()
    $validationRange.Validation.Add(3, 1, 1, '=LimoProductList')
    $validationRange.Validation.IgnoreBlank = $true
    $validationRange.Validation.InCellDropdown = $true
    # Keep the dropdown suggestions while allowing models not yet in the database.
    $validationRange.Validation.ShowError = $false

    # 旧模板前四行曾遗留 1、2、3、4 测试值；报价单单价列统一沿用第 15 行公式和字体。
    $priceSeed = $quoteSheet.Range('E15')
    if ($priceSeed.HasFormula) {
        foreach ($cell in $quoteSheet.Range('E11:E14').Cells) {
            if (-not $cell.HasFormula -and $cell.Value2 -in @(1, 2, 3, 4)) {
                $cell.FormulaR1C1 = $priceSeed.FormulaR1C1
            }
        }
    }
    $priceRange = $quoteSheet.Range('E11:E25')
    $priceRange.Font.Name = 'Microsoft YaHei'
    $priceRange.Font.Size = 16
    $priceRange.Font.Bold = $true
    $priceRange.NumberFormat = ([string][char]0x00A5) + '#,##0'

    foreach ($cell in $quoteSheet.Range('E11:J25').Cells) {
        if ($cell.HasFormula) {
            $formula = [string]$cell.Formula
            $formula = [regex]::Replace($formula, '\$A(?:\$6)?:\$K(?:\$\d+)?', ('$A$6:$K$' + $dataEndRow))
            $formula = [regex]::Replace($formula, '\$A(?:\$6)?:\$A(?:\$\d+)?', ('$A$6:$A$' + $dataEndRow))
            $formula = [regex]::Replace($formula, '\$E(?:\$6)?:\$E(?:\$\d+)?', ('$E$6:$E$' + $dataEndRow))
            $formula = [regex]::Replace($formula, '\$F(?:\$6)?:\$F(?:\$\d+)?', ('$F$6:$F$' + $dataEndRow))
            $cell.Formula = $formula
        }
    }
    # VBA uses this hidden text copy to restore automatic unit-price formulas
    # after a user manually overrides a price or starts a new quote.
    $detailSheet = $workbook.Worksheets.Item('历史明细')
    $detailSheet.Range('K1').NumberFormat = '@'
    $detailSheet.Range('K1').Value2 = [string]$quoteSheet.Range('E15').FormulaR1C1
    [void]$detailSheet.Range('K2').ClearContents()
    $excel.Calculation = -4105
    $excel.CalculateFullRebuild()
    $workbook.Save()
    Write-Output ("SYNCED_ROWS=" + $count)
}
catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    [Console]::Error.WriteLine($_.InvocationInfo.PositionMessage)
    exit 1
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $excel) {
        if ($null -ne $oldCalculation) { try { $excel.Calculation = $oldCalculation } catch {} }
        $excel.Quit()
    }
    if ($null -ne $validationRange) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($validationRange) }
    if ($null -ne $quoteSheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($quoteSheet) }
    if ($null -ne $sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
    if ($null -ne $workbook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
    if ($null -ne $excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
