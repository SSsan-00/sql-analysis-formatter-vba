$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$workbookPath = Join-Path $repoRoot 'SqlAnalysisFormatter.xlsm'
$mainModulePath = Join-Path $repoRoot 'src\vba\SqlAnalysisFormatter.bas'
$testModulePath = Join-Path $repoRoot 'src\vba\SqlAnalysisFormatterTests.bas'
$tempWorkbookPath = Join-Path $env:TEMP ('SqlAnalysisFormatter_Tests_' + [guid]::NewGuid().ToString('N') + '.xlsm')

function Release-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null
    }
}

Copy-Item -LiteralPath $workbookPath -Destination $tempWorkbookPath -Force

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

try {
    $workbook = $excel.Workbooks.Open($tempWorkbookPath)
    $components = $workbook.VBProject.VBComponents

    foreach ($moduleName in @('SqlAnalysisFormatter', 'SqlAnalysisFormatterTests')) {
        try {
            $components.Remove($components.Item($moduleName))
        } catch {
        }
    }

    $components.Import($mainModulePath) | Out-Null
    $components.Import($testModulePath) | Out-Null
    $excel.Run("'$tempWorkbookPath'!RunAllSqlAnalysisFormatterTests", $false) | Out-Null

    Write-Output 'VBA tests passed.'
} finally {
    if ($null -ne $workbook) {
        $workbook.Close($false) | Out-Null
    }
    $excel.Quit()
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Remove-Item -LiteralPath $tempWorkbookPath -Force -ErrorAction SilentlyContinue
}
