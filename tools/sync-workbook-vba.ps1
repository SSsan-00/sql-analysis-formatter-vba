param(
    [string]$WorkbookPath = '',
    [string]$MainModulePath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $repoRoot 'SqlAnalysisFormatter.xlsm'
}
if ([string]::IsNullOrWhiteSpace($MainModulePath)) {
    $MainModulePath = Join-Path $repoRoot 'src\vba\SqlAnalysisFormatter.bas'
}

$WorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$MainModulePath = (Resolve-Path -LiteralPath $MainModulePath).Path
$vbaSourceDirectory = Split-Path -Parent $MainModulePath
$productionComponents = @(
    @{ Name = 'SqlAnalysisToastEvents'; Path = Join-Path $vbaSourceDirectory 'SqlAnalysisToastEvents.cls' },
    @{ Name = 'SqlAnalysisToastManager'; Path = Join-Path $vbaSourceDirectory 'SqlAnalysisToastManager.bas' },
    @{ Name = 'SqlAnalysisToast'; Path = Join-Path $vbaSourceDirectory 'SqlAnalysisToast.frm' },
    @{ Name = 'SqlAnalysisFormatter'; Path = $MainModulePath }
)
$tempWorkbookPath = Join-Path $env:TEMP ('SqlAnalysisFormatter_Sync_' + [guid]::NewGuid().ToString('N') + '.xlsm')

function Release-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

$excel = $null
$workbooks = $null
$workbook = $null
$vbProject = $null
$components = $null
$sqlSheet = $null
$a1Cell = $null
$activeWindow = $null
$prepared = $false

Copy-Item -LiteralPath $WorkbookPath -Destination $tempWorkbookPath -Force

try {
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
        $excel.AutomationSecurity = 1
        $workbooks = $excel.Workbooks
        $workbook = $workbooks.Open($tempWorkbookPath, 0, $false)
        $vbProject = $workbook.VBProject
        $components = $vbProject.VBComponents

        foreach ($moduleName in @(
            'SqlAnalysisFormatter',
            'SqlAnalysisToastManager',
            'SqlAnalysisToastEvents',
            'SqlAnalysisToast',
            'SqlAnalysisFormatterTests',
            'SqlAnalysisFormatterGoldenTests')) {
            $existingComponent = $null
            try {
                $existingComponent = $components.Item($moduleName)
                $components.Remove($existingComponent)
            } catch {
                if ($moduleName -eq 'SqlAnalysisFormatter') {
                    throw
                }
            } finally {
                Release-ComObject $existingComponent
            }
        }

        foreach ($productionComponent in $productionComponents) {
            $componentPath = (Resolve-Path -LiteralPath $productionComponent.Path).Path
            $importedComponent = $components.Import($componentPath)
            Release-ComObject $importedComponent
        }

        $excel.Run("'$tempWorkbookPath'!SetupWorkbook") | Out-Null
        $excel.Run("'$tempWorkbookPath'!ClearData", $false) | Out-Null

        $sqlSheet = $workbook.Worksheets.Item(2)
        $sqlSheet.Activate() | Out-Null
        $a1Cell = $sqlSheet.Range('A1')
        $a1Cell.Select() | Out-Null
        $activeWindow = $excel.ActiveWindow
        $activeWindow.ScrollRow = 1
        $activeWindow.ScrollColumn = 1

        $workbook.Save()
        $prepared = $true
    } finally {
        Release-ComObject $activeWindow
        Release-ComObject $a1Cell
        Release-ComObject $sqlSheet
        Release-ComObject $components
        Release-ComObject $vbProject
        if ($null -ne $workbook) {
            try {
                $workbook.Close($false) | Out-Null
            } catch {
                Write-Warning "Could not close the workbook used for synchronization: $($_.Exception.Message)"
            }
        }
        Release-ComObject $workbook
        Release-ComObject $workbooks
        if ($null -ne $excel) {
            try {
                $excel.Quit()
            } catch {
                Write-Warning "Could not quit Excel: $($_.Exception.Message)"
            }
        }
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    if (-not $prepared) {
        throw 'Workbook synchronization did not complete.'
    }
    Copy-Item -LiteralPath $tempWorkbookPath -Destination $WorkbookPath -Force
    Write-Output "VBA synchronization and workbook initialization completed: $WorkbookPath"
} finally {
    Remove-Item -LiteralPath $tempWorkbookPath -Force -ErrorAction SilentlyContinue
}
