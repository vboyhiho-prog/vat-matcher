param(
    [string]$Version = '1.41'
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$workspaceRoot = Split-Path -Parent $projectRoot
$outputDir = Join-Path $workspaceRoot ('outputs\vat-matcher-v' + $Version)
$target = Join-Path $outputDir ('VAT_Matcher_v' + $Version + '.xlsm')
$manifest = Join-Path $outputDir ('VAT_Matcher_v' + $Version + '.sha256')
$buildReport = Join-Path $outputDir 'BUILD_REPORT.md'
if (-not (Test-Path -LiteralPath $target)) { throw "Release does not exist: $target" }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$excel.AutomationSecurity = 3
$wb = $null
try {
    $wb = $excel.Workbooks.Open($target, $false, $true)
    if ($excel.Calculation -ne -4105) { throw "Release calculation mode is not Automatic: $($excel.Calculation)." }
    if ($wb.Worksheets.Item('DASHBOARD').Shapes.Count -ne 11) { throw 'Dashboard does not contain 11 buttons.' }

    $runtimeTables = @(
        @('PDF_FILES','tblPdfFiles'), @('INVOICES','tblInvoices'), @('VAT_LINES','tblVatLines'),
        @('GR_DATA','tblGrData'), @('MATCH_CANDIDATES','tblCandidates'), @('ALLOCATIONS','tblAllocations'),
        @('BC_HOA_DON','tblInvoiceReport'), @('BC_PHIEU','tblReceiptReport'), @('LOG','tblLog'),
        @('EMAIL_ATTACHMENTS','tblEmailAttachments'), @('EMAIL_HINTS','tblEmailHints'), @('PQ_PDF_RAW','tblPdfRaw')
    )
    foreach ($tableRef in $runtimeTables) {
        $count = $wb.Worksheets.Item($tableRef[0]).ListObjects.Item($tableRef[1]).ListRows.Count
        if ($count -ne 0) { throw "$($tableRef[1]) contains $count runtime row(s)." }
    }

    $homeSheet = $wb.Worksheets.Item('HOME')
    if (-not [string]::IsNullOrWhiteSpace([string]$homeSheet.Range('B4').Value2)) { throw 'Release PDF path is not blank.' }
    if (-not [string]::IsNullOrWhiteSpace([string]$homeSheet.Range('D4').Value2)) { throw 'Release Tracking P path is not blank.' }
    $homeTable = $homeSheet.ListObjects.Item('tblHome')
    $expectedHomeHeaders = @('Key','Value','Status')
    if ($homeTable.ListColumns.Count -ne $expectedHomeHeaders.Count) { throw 'tblHome has an unexpected column count.' }
    for ($i = 1; $i -le $expectedHomeHeaders.Count; $i++) {
        if ([string]$homeTable.HeaderRowRange.Cells.Item(1, $i).Value2 -ne $expectedHomeHeaders[$i - 1]) { throw 'tblHome headers are not normalized.' }
    }

    $tests = $wb.Worksheets.Item('TEST_RESULTS').ListObjects.Item('tblTestResults')
    if ($tests.ListRows.Count -eq 0) { throw 'TEST_RESULTS is empty.' }
    foreach ($testRow in @($tests.ListRows)) {
        if ([string]$testRow.Range.Cells.Item(1, 4).Value2 -ne 'PASS') {
            throw ('Non-PASS assertion: ' + [string]$testRow.Range.Cells.Item(1, 1).Value2)
        }
    }
    $testCount = $tests.ListRows.Count

    $expectedModules = @(Get-ChildItem -LiteralPath (Join-Path $projectRoot 'src\modules') -Filter '*.bas' -File | ForEach-Object BaseName)
    $actualModules = @($wb.VBProject.VBComponents | ForEach-Object Name)
    foreach ($moduleName in $expectedModules) {
        if ($actualModules -notcontains $moduleName) { throw "Release is missing VBA module $moduleName." }
    }
    $moduleCount = $expectedModules.Count
    $wb.Close($false)
    $wb = $null
}
finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
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
    ('- Imported VBA modules: ' + $moduleCount),
    ('- Acceptance assertions preserved in workbook: ' + $testCount + ' PASS; 0 FAIL'),
    '- Workbook reopened read-only with Calculation=Automatic and 11 dashboard buttons.',
    '- Runtime GR/PDF/email/report/log tables and source paths are empty.',
    '- HOME path storage is normalized outside tblHome headers.'
)
Set-Content -LiteralPath $buildReport -Value $reportLines -Encoding utf8
Write-Output "PASS: $target"
Write-Output "SHA256: $hash"
Write-Output "Acceptance assertions: $testCount PASS"
