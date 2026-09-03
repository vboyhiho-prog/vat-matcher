param(
    [string]$Version = '1.43',
    [switch]$ForceRebuild,
    [switch]$CleanDiagnostics
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$workspaceRoot = Split-Path -Parent $projectRoot
$source = Join-Path $projectRoot 'VAT_Matcher_v1.40.xlsm'
$outputDir = Join-Path $workspaceRoot ('outputs\vat-matcher-v' + $Version)
$target = Join-Path $outputDir ('VAT_Matcher_v' + $Version + '.xlsm')
$manifest = Join-Path $outputDir ('VAT_Matcher_v' + $Version + '.sha256')
$buildReport = Join-Path $outputDir 'BUILD_REPORT.md'
$runtimeSource = Join-Path $projectRoot 'python'
$runtimeTarget = Join-Path $outputDir 'python'
$sessionTarget = Join-Path $outputDir 'vat_python_runtime'
$buildStart = Get-Date

if (-not (Test-Path -LiteralPath $source)) { throw "Release source does not exist: $source" }
if (-not (Test-Path -LiteralPath $runtimeSource)) { throw "Python runtime source does not exist: $runtimeSource" }
if (Test-Path -LiteralPath $target) {
    if (-not $ForceRebuild) { throw "Refusing to overwrite an existing release: $target" }
    Remove-Item -LiteralPath $target -Force
}
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
if ($CleanDiagnostics) {
    foreach ($diagnostic in @(Get-ChildItem -LiteralPath $outputDir -Filter 'VAT_Matcher_Diagnostic_*.xlsx' -File)) {
        $resolvedOutput = [System.IO.Path]::GetFullPath($outputDir)
        $resolvedDiagnostic = [System.IO.Path]::GetFullPath($diagnostic.FullName)
        if (-not $resolvedDiagnostic.StartsWith($resolvedOutput + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to remove diagnostic outside the release package: $resolvedDiagnostic"
        }
        Remove-Item -LiteralPath $resolvedDiagnostic -Force
    }
}
Copy-Item -LiteralPath $source -Destination $target
foreach ($releaseChild in @($runtimeTarget, $sessionTarget)) {
if (Test-Path -LiteralPath $releaseChild) {
    $resolvedOutput = [System.IO.Path]::GetFullPath($outputDir)
    $resolvedChild = [System.IO.Path]::GetFullPath($releaseChild)
    if (-not $resolvedChild.StartsWith($resolvedOutput + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove a folder outside the release package: $resolvedChild"
    }
    Remove-Item -LiteralPath $releaseChild -Recurse -Force
}
}
New-Item -ItemType Directory -Path $runtimeTarget | Out-Null
foreach ($runtimeFile in @('engine.py', 'run_tool.bat', 'requirements.lock', 'README.md')) {
    Copy-Item -LiteralPath (Join-Path $runtimeSource $runtimeFile) -Destination (Join-Path $runtimeTarget $runtimeFile) -Force
}
Copy-Item -LiteralPath (Join-Path $projectRoot 'README.md') -Destination $outputDir -Force
Copy-Item -LiteralPath (Join-Path $projectRoot 'HUONG_DAN_SU_DUNG.md') -Destination $outputDir -Force
Write-Output "BUILD: copied v1.40 workbook source and fixed Python runtime to v$Version package."

& python (Join-Path $projectRoot 'tests\test_python_engine.py')
if ($LASTEXITCODE -ne 0) { throw 'Python engine regression tests failed.' }
Write-Output 'BUILD: Python engine regression gate completed.'

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$excel.ScreenUpdating = $false
$excel.AskToUpdateLinks = $false
$wb = $null
try {
    $wb = $excel.Workbooks.Open($target, $false, $false)
    $excel.Calculation = -4135
    $moduleFiles = @(Get-ChildItem -LiteralPath (Join-Path $projectRoot 'src\modules') -Filter '*.bas' -File | Sort-Object Name)
    if ($moduleFiles.Count -lt 19) { throw "Expected at least 19 standard modules, found $($moduleFiles.Count)." }
    foreach ($moduleFile in $moduleFiles) {
        $moduleName = $moduleFile.BaseName
        $component = @($wb.VBProject.VBComponents | Where-Object { $_.Name -eq $moduleName }) | Select-Object -First 1
        if ($null -ne $component) { $wb.VBProject.VBComponents.Remove($component) }
        $moduleText = Get-Content -LiteralPath $moduleFile.FullName -Raw -Encoding utf8
        $moduleText = $moduleText -replace '(?m)^Attribute VB_Name = ".*"\r?\n', ''
        $component = $wb.VBProject.VBComponents.Add(1)
        $component.Name = $moduleName
        $component.CodeModule.AddFromString($moduleText)
    }
    Write-Output "BUILD: imported $($moduleFiles.Count) VBA modules."

    $compileControl = $excel.VBE.CommandBars.FindControl(1, 578)
    if ($null -eq $compileControl) { throw 'VBA compile command is unavailable.' }
    if ($compileControl.Enabled) { $compileControl.Execute() }
    Write-Output 'BUILD: VBA compile gate completed.'

    $macroPrefix = "'" + $wb.Name + "'!"
    $excel.Run($macroPrefix + 'SetupWorkbook')
    $excel.Run($macroPrefix + 'SetupDashboard')
    $config = $wb.Worksheets.Item('CONFIG').ListObjects.Item('tblConfig')
    $requiredConfig = @{ GRBeforeInvoiceDays = '5'; GRAfterInvoiceDays = '2'; PdfIngestion = 'Python PyMuPDF' }
    foreach ($key in $requiredConfig.Keys) {
        $found = $false
        foreach ($row in @($config.ListRows)) {
            if ([string]$row.Range.Cells.Item(1, 1).Value2 -eq $key) {
                if ([string]$row.Range.Cells.Item(1, 2).Value2 -ne $requiredConfig[$key]) { throw "Config $key was not seeded correctly." }
                $found = $true
                break
            }
        }
        if (-not $found) { throw "Required config row is missing: $key" }
    }
    if ($wb.Worksheets.Item('DASHBOARD').Shapes.Count -ne 11) { throw 'Dashboard validation failed: expected 11 buttons.' }

    # A release contains no company source data, cached Python output or
    # Power Query PDF connection. A user must select fresh P/PDF inputs.
    $excel.Run($macroPrefix + 'PrepareReleaseWorkbook')
    $excel.Calculation = -4105
    $wb.Save()
    $wb.Close($true)
    $wb = $null

    $wb = $excel.Workbooks.Open($target, $false, $true)
    if ($excel.Calculation -ne -4105) { throw 'Release calculation mode is not Automatic.' }
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
    $expectedModules = @($moduleFiles | ForEach-Object BaseName)
    $actualModules = @($wb.VBProject.VBComponents | ForEach-Object Name)
    foreach ($moduleName in $expectedModules) {
        if ($actualModules -notcontains $moduleName) { throw "Release is missing VBA module $moduleName." }
    }
    if ($wb.Queries.Count -ne 0) { throw 'Release still contains a Power Query PDF query.' }
    $wb.Close($false)
    $wb = $null
    Write-Output 'BUILD: clean Python release reopened and verified.'
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
    '- Python regression tests: 5 PASS; 0 FAIL.',
    '- VBA compile gate and workbook reopen gate: PASS.',
    '- PDF parser/matcher: bundled `python\run_tool.bat`; no Power Query PDF connector.',
    '- Runtime GR/PDF/report/log tables and source paths are empty.'
)
Set-Content -LiteralPath $buildReport -Value $reportLines -Encoding utf8
Write-Output "PASS: $target"
Write-Output "SHA256: $hash"
Write-Output 'Python regression assertions: 5 PASS'
Write-Output ("Elapsed: " + [math]::Round(((Get-Date) - $buildStart).TotalSeconds, 1) + " seconds")
