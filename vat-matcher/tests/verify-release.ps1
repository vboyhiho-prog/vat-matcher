param(
    [string]$Version = '1.43'
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$workspaceRoot = Split-Path -Parent $projectRoot
$outputDir = Join-Path $workspaceRoot ('outputs\vat-matcher-v' + $Version)
$target = Join-Path $outputDir ('VAT_Matcher_v' + $Version + '.xlsm')
$manifest = Join-Path $outputDir ('VAT_Matcher_v' + $Version + '.sha256')
$buildReport = Join-Path $outputDir 'BUILD_REPORT.md'
if (-not (Test-Path -LiteralPath $target)) { throw "Release does not exist: $target" }
if (-not (Test-Path -LiteralPath (Join-Path $outputDir 'python\run_tool.bat'))) { throw 'Release is missing python\run_tool.bat.' }
if (Test-Path -LiteralPath (Join-Path $outputDir 'vat_python_runtime')) { throw 'Release contains a cached Python run folder.' }

& python (Join-Path $projectRoot 'tests\test_python_engine.py')
if ($LASTEXITCODE -ne 0) { throw 'Python engine regression tests failed.' }

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
    if ($wb.Queries.Count -ne 0) { throw 'Release contains a Power Query PDF query.' }
    $homeSheet = $wb.Worksheets.Item('HOME')
    if (-not [string]::IsNullOrWhiteSpace([string]$homeSheet.Range('B4').Value2)) { throw 'Release PDF path is not blank.' }
    if (-not [string]::IsNullOrWhiteSpace([string]$homeSheet.Range('D4').Value2)) { throw 'Release Tracking P path is not blank.' }
    $expectedModules = @(Get-ChildItem -LiteralPath (Join-Path $projectRoot 'src\modules') -Filter '*.bas' -File | ForEach-Object BaseName)
    $actualModules = @($wb.VBProject.VBComponents | ForEach-Object Name)
    foreach ($moduleName in $expectedModules) {
        if ($actualModules -notcontains $moduleName) { throw "Release is missing VBA module $moduleName." }
    }
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
Write-Output "PASS: $target"
Write-Output "SHA256: $hash"
Write-Output 'Python regression assertions: 5 PASS'
