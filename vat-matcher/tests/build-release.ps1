param(
    [string]$Version = '1.41',
    [switch]$ForceRebuild
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$workspaceRoot = Split-Path -Parent $projectRoot
$source = Join-Path $projectRoot 'VAT_Matcher_v1.40.xlsm'
$outputDir = Join-Path $workspaceRoot ('outputs\vat-matcher-v' + $Version)
$target = Join-Path $outputDir ('VAT_Matcher_v' + $Version + '.xlsm')
$manifest = Join-Path $outputDir ('VAT_Matcher_v' + $Version + '.sha256')
$buildReport = Join-Path $outputDir 'BUILD_REPORT.md'
$buildStart = Get-Date

if (-not (Test-Path -LiteralPath $source)) { throw "Release source does not exist: $source" }
if (Test-Path -LiteralPath $target) {
    if (-not $ForceRebuild) { throw "Refusing to overwrite existing release: $target" }
    Remove-Item -LiteralPath $target -Force
}
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
Copy-Item -LiteralPath $source -Destination $target
Write-Output "BUILD: copied v1.40 source to disposable v$Version target."

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$excel.ScreenUpdating = $false
$excel.AskToUpdateLinks = $false
$wb = $null
try {
    $wb = $excel.Workbooks.Open($target, $false, $false)
    $excel.Calculation = -4135 # xlCalculationManual during the disposable acceptance run

    $moduleFiles = @(Get-ChildItem -LiteralPath (Join-Path $projectRoot 'src\modules') -Filter '*.bas' -File | Sort-Object Name)
    if ($moduleFiles.Count -lt 18) { throw "Expected at least 18 standard modules, found $($moduleFiles.Count)." }
    foreach ($moduleFile in $moduleFiles) {
        $moduleName = $moduleFile.BaseName
        $component = $null
        foreach ($item in @($wb.VBProject.VBComponents)) {
            if ($item.Name -eq $moduleName) { $component = $item; break }
        }
        if ($null -ne $component) { $wb.VBProject.VBComponents.Remove($component) }
        $moduleText = Get-Content -LiteralPath $moduleFile.FullName -Raw -Encoding utf8
        $moduleText = $moduleText -replace '(?m)^Attribute VB_Name = ".*"\r?\n', ''
        $component = $wb.VBProject.VBComponents.Add(1)
        $component.Name = $moduleName
        $component.CodeModule.AddFromString($moduleText)
    }
    Write-Output "BUILD: imported $($moduleFiles.Count) VBA modules."

    # Compile every module before any acceptance macro is allowed to run.
    $compileControl = $excel.VBE.CommandBars.FindControl(1, 578)
    if ($null -eq $compileControl) { throw 'VBA compile command is unavailable.' }
    if ($compileControl.Enabled) { $compileControl.Execute() }
    Write-Output 'BUILD: VBA compile gate completed.'

    $macroPrefix = "'" + $wb.Name + "'!"
    $excel.Run($macroPrefix + 'SetupWorkbook')
    $excel.Run($macroPrefix + 'SetupDashboard')

    # Seed the seven known IB hints on the disposable build copy so T13 is a real assertion, not a skip.
    $hints = $wb.Worksheets.Item('EMAIL_HINTS').ListObjects.Item('tblEmailHints')
    if ($hints.ListRows.Count -gt 0) { $hints.DataBodyRange.Delete() }
    $gr = $wb.Worksheets.Item('GR_DATA').ListObjects.Item('tblGrData')
    $grData = $gr.DataBodyRange.Value2
    $ibRows = @{}
    for ($i = 1; $i -le $gr.ListRows.Count; $i++) {
        $ibValue = [string]$grData.GetValue($i, 8)
        if (-not [string]::IsNullOrWhiteSpace($ibValue) -and -not $ibRows.ContainsKey($ibValue)) { $ibRows[$ibValue] = $i }
    }
    for ($ib = 185068322; $ib -le 185068328; $ib++) {
        if (-not $ibRows.ContainsKey([string]$ib)) { throw "T13 fixture missing GR IB $ib." }
        $grRow = $ibRows[[string]$ib]
        $row = $hints.ListRows.Add()
        $row.Range.Cells.Item(1, 1).Resize(1, 8).Value2 = @(
            'BUILD_FIXTURE.msg', [string]$ib,
            [string]$grData.GetValue($grRow, 3),
            [double]$grData.GetValue($grRow, 7),
            $grData.GetValue($grRow, 9),
            [string]$grData.GetValue($grRow, 4),
            'HIGH', 'Acceptance fixture for IB-to-receipt 5007'
        )
        $row.Range.Cells.Item(1, 2).NumberFormat = '@'
        $row.Range.Cells.Item(1, 2).Value2 = [string]$ib
    }

    $excel.Run($macroPrefix + 'SetMatcherAutomationMode', $true)
    $excel.Run($macroPrefix + 'EnsureMaterialVendorMap')
    $excel.Run($macroPrefix + 'ParsePdfRawToInvoices')
    $excel.Run($macroPrefix + 'ParseVatLinesFromPdfRaw')
    Write-Output 'BUILD: sample PDF parsing completed.'
    $excel.Run($macroPrefix + 'RunMatch')
    Write-Output 'BUILD: matching and allocation completed.'
    $excel.Run($macroPrefix + 'VietHoaLyDoTrongBaoCao')
    $excel.Run($macroPrefix + 'CreateRenamePreviews')
    $excel.Run($macroPrefix + 'BuildReceiptReport')
    $excel.Run($macroPrefix + 'RunWorkbookSelfTestSilent')
    $excel.Run($macroPrefix + 'SetMatcherAutomationMode', $false)

    $testTable = $wb.Worksheets.Item('TEST_RESULTS').ListObjects.Item('tblTestResults')
    if ($testTable.ListRows.Count -eq 0) { throw 'Acceptance test table is empty.' }
    $failures = @()
    foreach ($testRow in @($testTable.ListRows)) {
        $result = [string]$testRow.Range.Cells.Item(1, 4).Value2
        if ($result -ne 'PASS') {
            $failures += ([string]$testRow.Range.Cells.Item(1, 1).Value2 + ': expected [' + [string]$testRow.Range.Cells.Item(1, 2).Value2 + '], actual [' + [string]$testRow.Range.Cells.Item(1, 3).Value2 + ']. ' + [string]$testRow.Range.Cells.Item(1, 5).Value2)
        }
    }
    if ($failures.Count -gt 0) { throw ("Acceptance failures:`n" + ($failures -join "`n")) }
    Write-Output "BUILD: $($testTable.ListRows.Count) workbook assertions passed."

    $invoiceReport = $wb.Worksheets.Item('BC_HOA_DON').ListObjects.Item('tblInvoiceReport')
    $partialRow = $null
    for ($i = 1; $i -le $invoiceReport.ListRows.Count; $i++) {
        if ([string]$invoiceReport.DataBodyRange.Cells.Item($i, 3).Value2 -eq '00002961') { $partialRow = $i; break }
    }
    if ($null -eq $partialRow) { throw 'T10 invoice 00002961 is missing.' }
    if ([string]$invoiceReport.DataBodyRange.Cells.Item($partialRow, 8).Value2 -ne 'PARTIAL_MATCHED') { throw 'T10 must be PARTIAL_MATCHED.' }
    $partialScore = [double]$invoiceReport.DataBodyRange.Cells.Item($partialRow, 7).Value2
    if ($partialScore -le 0 -or $partialScore -ge 100) { throw "T10 partial score must be between 0 and 100; got $partialScore." }
    if ([string]::IsNullOrWhiteSpace([string]$invoiceReport.DataBodyRange.Cells.Item($partialRow, 9).Value2)) { throw 'T10 rename preview is missing.' }
    if (-not [string]::IsNullOrWhiteSpace([string]$invoiceReport.DataBodyRange.Cells.Item($partialRow, 10).Value2)) { throw 'T10 Decision must default to blank.' }

    $allocations = $wb.Worksheets.Item('ALLOCATIONS').ListObjects.Item('tblAllocations')
    $capacity = @{}
    for ($i = 1; $i -le $gr.ListRows.Count; $i++) { $capacity[[string]$grData.GetValue($i, 1)] = [double]$grData.GetValue($i, 7) }
    $used = @{}
    $sourceOwner = @{}
    $allocationData = $allocations.DataBodyRange.Value2
    for ($i = 1; $i -le $allocations.ListRows.Count; $i++) {
        $key = [string]$allocationData.GetValue($i, 5)
        $invoiceId = [string]$allocationData.GetValue($i, 2)
        if (-not $used.ContainsKey($key)) { $used[$key] = 0.0 }
        $used[$key] += [double]$allocationData.GetValue($i, 6)
        if ($sourceOwner.ContainsKey($key) -and $sourceOwner[$key] -ne $invoiceId) { throw "T22 SourceRow ${key} is assigned to both $($sourceOwner[$key]) and $invoiceId." }
        $sourceOwner[$key] = $invoiceId
    }
    foreach ($key in $used.Keys) {
        if (-not $capacity.ContainsKey($key) -or $used[$key] -gt $capacity[$key] + 0.000001) { throw "T21 capacity conflict at SourceRow ${key}: $($used[$key]) > $($capacity[$key])." }
    }

    $shapeCount = $wb.Worksheets.Item('DASHBOARD').Shapes.Count
    if ($shapeCount -ne 11) { throw "Dashboard validation failed: expected 11 buttons, got $shapeCount." }

    # Remove all batch/test data and paths after the acceptance evidence has been written.
    $excel.Run($macroPrefix + 'PrepareReleaseWorkbook')
    $excel.Calculation = -4105 # xlCalculationAutomatic in the saved release
    $wb.Save()
    $wb.Close($true)
    $wb = $null

    $wb = $excel.Workbooks.Open($target, $false, $true)
    if ($wb.Worksheets.Item('DASHBOARD').Shapes.Count -ne 11) { throw 'Reopen validation failed: dashboard buttons missing.' }
    $runtimeTables = @(
        @('PDF_FILES','tblPdfFiles'), @('INVOICES','tblInvoices'), @('VAT_LINES','tblVatLines'),
        @('GR_DATA','tblGrData'), @('MATCH_CANDIDATES','tblCandidates'), @('ALLOCATIONS','tblAllocations'),
        @('BC_HOA_DON','tblInvoiceReport'), @('BC_PHIEU','tblReceiptReport'), @('LOG','tblLog'),
        @('EMAIL_ATTACHMENTS','tblEmailAttachments'), @('EMAIL_HINTS','tblEmailHints'), @('PQ_PDF_RAW','tblPdfRaw')
    )
    foreach ($tableRef in $runtimeTables) {
        $count = $wb.Worksheets.Item($tableRef[0]).ListObjects.Item($tableRef[1]).ListRows.Count
        if ($count -ne 0) { throw "Release sanitization failed: $($tableRef[1]) contains $count row(s)." }
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$wb.Worksheets.Item('HOME').Range('B4').Value2)) { throw 'Release PDF path is not blank.' }
    if (-not [string]::IsNullOrWhiteSpace([string]$wb.Worksheets.Item('HOME').Range('D4').Value2)) { throw 'Release Tracking P path is not blank.' }
    $reopenTests = $wb.Worksheets.Item('TEST_RESULTS').ListObjects.Item('tblTestResults')
    if ($reopenTests.ListRows.Count -eq 0) { throw 'Release lost its acceptance evidence.' }
    foreach ($testRow in @($reopenTests.ListRows)) {
        if ([string]$testRow.Range.Cells.Item(1, 4).Value2 -ne 'PASS') { throw 'Release contains a non-PASS acceptance result.' }
    }
    $testCount = $reopenTests.ListRows.Count
    $wb.Close($false)
    $wb = $null
    Write-Output 'BUILD: clean release reopened and verified.'
}
finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    $excel.EnableEvents = $true
    $excel.ScreenUpdating = $true
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$hash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
Set-Content -LiteralPath $manifest -Value ("$hash  " + [System.IO.Path]::GetFileName($target)) -Encoding ascii
$reportLines = @(
    '# VAT Matcher release build', '',
    ('- Version: ' + $Version),
    ('- Workbook: ' + [System.IO.Path]::GetFileName($target)),
    ('- SHA-256: `' + $hash + '`'),
    ('- Imported VBA modules: ' + $moduleFiles.Count),
    ('- Acceptance assertions: ' + $testCount + ' PASS; 0 FAIL'),
    '- Workbook reopened successfully with 11 dashboard buttons.',
    '- Runtime GR/PDF/email/report/log tables and source paths are empty.'
)
Set-Content -LiteralPath $buildReport -Value $reportLines -Encoding utf8
Write-Output "PASS: $target"
Write-Output "SHA256: $hash"
Write-Output "Acceptance assertions: $testCount PASS"
Write-Output ("Elapsed: " + [math]::Round(((Get-Date) - $buildStart).TotalSeconds, 1) + " seconds")
