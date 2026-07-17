param(
    [string]$ParserExePath = '',
    [string[]]$TestName = @()
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$workbookPath = Join-Path $repoRoot 'SqlAnalysisFormatter.xlsm'
$mainModulePath = Join-Path $repoRoot 'src\vba\SqlAnalysisFormatter.bas'
$testModulePath = Join-Path $repoRoot 'src\vba\SqlAnalysisFormatterTests.bas'
$tempWorkbookPath = Join-Path $env:TEMP ('SqlAnalysisFormatter_Tests_' + [guid]::NewGuid().ToString('N') + '.xlsm')
$previousParserExePath = $env:SQL_ANALYSIS_FORMATTER_PARSER_EXE

function Release-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject) | Out-Null
    }
}

Copy-Item -LiteralPath $workbookPath -Destination $tempWorkbookPath -Force

if (-not [string]::IsNullOrWhiteSpace($ParserExePath)) {
    $env:SQL_ANALYSIS_FORMATTER_PARSER_EXE = (Resolve-Path $ParserExePath)
}

$excel = $null
$workbooks = $null
$workbook = $null
$vbProject = $null
$components = $null
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

try {
    $workbooks = $excel.Workbooks
    $workbook = $workbooks.Open($tempWorkbookPath)
    Release-ComObject $workbooks
    $workbooks = $null
    $vbProject = $workbook.VBProject
    $components = $vbProject.VBComponents

    foreach ($moduleName in @('SqlAnalysisFormatter', 'SqlAnalysisFormatterTests')) {
        $existingComponent = $null
        try {
            $existingComponent = $components.Item($moduleName)
            $components.Remove($existingComponent)
        } catch {
        } finally {
            Release-ComObject $existingComponent
        }
    }

    $importedComponent = $components.Import($mainModulePath)
    Release-ComObject $importedComponent
    $importedComponent = $components.Import($testModulePath)
    Release-ComObject $importedComponent
    $testMacros = @(
        'SetupWorkbook_CreatesOutputSheet',
        'CopyOutput_CopiesRenderedRange',
        'AnalyzeQueries_ConvertsCrudFixtures',
        'AnalyzeQueries_ConvertsTsqlFunctionFixtures',
        'AnalyzeQueries_WritesWithSubqueriesInsideOut',
        'AnalyzeQueries_PreservesLeadingApostropheInOutput',
        'AnalyzeQueries_DisablesWrappingAfterWritingLongText',
        'AnalyzeQueries_RendersDeeplyNestedCaseConditions',
        'AnalyzeQueries_NormalizesInvisibleOutputWhitespace',
        'AnalyzeQueries_SeparatesTransferExpressionsFromColumns',
        'AnalyzeQueries_RendersWrappedUpdateCaseAsTransferMethod',
        'AnalyzeQueries_HandlesSyntaxCharactersInFieldNames',
        'AnalyzeQueries_UsesStandaloneTableNameForSingleTable',
        'AnalyzeQueries_WritesUnsupportedQueryAsIs',
        'AnalyzeQueries_FramesOnlyTableBody',
        'ClearData_ClearsOutputSheet'
    )
    if ($TestName.Count -gt 0) {
        $testMacros = @($testMacros | Where-Object { $_ -in $TestName })
        if ($testMacros.Count -eq 0) {
            throw "VBA test not found: $($TestName -join ', ')"
        }
    }
    for ($testIndex = 0; $testIndex -lt $testMacros.Count; $testIndex++) {
        $macroName = $testMacros[$testIndex]
        $excel.Run("'$tempWorkbookPath'!$macroName") | Out-Null
        Write-Output ("VBA test progress: {0}/{1} {2}" -f ($testIndex + 1), $testMacros.Count, $macroName)
    }

    Write-Output 'VBA tests passed.'
} finally {
    Release-ComObject $components
    Release-ComObject $vbProject
    if ($null -ne $workbook) {
        try {
            $workbook.Close($false) | Out-Null
        } catch {
            Write-Warning "テスト用ブックを閉じられませんでした: $($_.Exception.Message)"
        }
    }
    Release-ComObject $workbook
    Release-ComObject $workbooks
    if ($null -ne $excel) {
        try {
            $excel.Quit()
        } catch {
            Write-Warning "Excelを終了できませんでした: $($_.Exception.Message)"
        }
    }
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Remove-Item -LiteralPath $tempWorkbookPath -Force -ErrorAction SilentlyContinue
    $env:SQL_ANALYSIS_FORMATTER_PARSER_EXE = $previousParserExePath
}
