param(
    [string]$ParserExePath = '',
    [string[]]$TestName = @(),
    [switch]$UseEmbeddedMainModule
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$workbookPath = Join-Path $repoRoot 'SqlAnalysisFormatter.xlsm'
$mainModulePath = Join-Path $repoRoot 'src\vba\SqlAnalysisFormatter.bas'
$productionComponents = @(
    @{ Name = 'SqlAnalysisToastEvents'; Path = Join-Path $repoRoot 'src\vba\SqlAnalysisToastEvents.cls' },
    @{ Name = 'SqlAnalysisToastManager'; Path = Join-Path $repoRoot 'src\vba\SqlAnalysisToastManager.bas' },
    @{ Name = 'SqlAnalysisToast'; Path = Join-Path $repoRoot 'src\vba\SqlAnalysisToast.frm' },
    @{ Name = 'SqlAnalysisFormatter'; Path = $mainModulePath }
)
$testModulePath = Join-Path $repoRoot 'src\vba\SqlAnalysisFormatterTests.bas'
$tempWorkbookPath = Join-Path $env:TEMP ('SqlAnalysisFormatter_Tests_' + [guid]::NewGuid().ToString('N') + '.xlsm')
$previousParserExePath = $env:SQL_ANALYSIS_FORMATTER_PARSER_EXE

function Release-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject) | Out-Null
    }
}

function Assert-ToastNotificationSourceContract {
    $source = Get-Content -LiteralPath $mainModulePath -Encoding UTF8 -Raw
    foreach ($requiredPattern in @(
        'SqlAnalysisToastManager\.ShowToast\s+AnalyzeDoneMessage\(\)',
        'SqlAnalysisToastManager\.ShowToast\s+ClearDoneMessage\(\)')) {
        if ($source -notmatch $requiredPattern) {
            throw "Completion toast call is missing from SqlAnalysisFormatter.bas: $requiredPattern"
        }
    }
    foreach ($forbiddenPattern in @(
        'MsgBox\s+AnalyzeDoneMessage\(\)',
        'MsgBox\s+ClearDoneMessage\(\)')) {
        if ($source -match $forbiddenPattern) {
            throw "Blocking completion dialog remains in SqlAnalysisFormatter.bas: $forbiddenPattern"
        }
    }
    foreach ($preservedPattern in @(
        'MsgBox\s+DuplicateTableWarningMessage\(duplicateTableIds\),\s*vbExclamation',
        'MsgBox\s+AnalyzeFallbackMessage\(fallbackReason\),\s*vbExclamation',
        'MsgBox\s*\(ClearConfirmMessage\(\),\s*vbQuestion')) {
        if ($source -notmatch $preservedPattern) {
            throw "Required warning or confirmation dialog was removed: $preservedPattern"
        }
    }
    $errorRaisePattern = 'Err\.Raise\s+errorNumber,\s*errorSource,\s*errorDescription'
    if ([regex]::Matches($source, $errorRaisePattern).Count -lt 2) {
        throw 'AnalyzeQueries or ClearData no longer propagates execution errors.'
    }
}

Assert-ToastNotificationSourceContract

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

    $modulesToRemove = @('SqlAnalysisFormatterTests')
    if (-not $UseEmbeddedMainModule) {
        $modulesToRemove += @($productionComponents.Name)
    }
    foreach ($moduleName in $modulesToRemove) {
        $existingComponent = $null
        try {
            $existingComponent = $components.Item($moduleName)
            $components.Remove($existingComponent)
        } catch {
        } finally {
            Release-ComObject $existingComponent
        }
    }

    if (-not $UseEmbeddedMainModule) {
        foreach ($productionComponent in $productionComponents) {
            $importedComponent = $components.Import($productionComponent.Path)
            Release-ComObject $importedComponent
        }
    }
    $importedComponent = $components.Import($testModulePath)
    Release-ComObject $importedComponent
    $testMacros = @(
        'SetupWorkbook_CreatesOutputSheet',
        'SetupWorkbook_CreatesOutputTwoLayout',
        'SetupWorkbook_CreatesStableSqlActionButtons',
        'SetupWorkbook_ProvidesToastForm',
        'SetupWorkbook_TracksMissingNameFillColor',
        'CompletionToast_UsesTwoSecondDuration',
        'CompletionToast_ShowsWithoutChangingSelection',
        'CompletionToast_ReplacesExistingNotification',
        'CompletionToast_DismissesImmediately',
        'CompletionToast_AutoDismissesAfterTwoSeconds',
        'AnalyzeQueries_ShowsCompletionToastOnSuccess',
        'AnalyzeQueries_ShowMessageFalseDoesNotShowToast',
        'CopyOutput_CopiesRenderedRange',
        'CopyOutputTwo_CopiesRenderedRange',
        'AnalyzeQueries_ConvertsCrudFixtures',
        'AnalyzeQueries_ConvertsTsqlFunctionFixtures',
        'AnalyzeQueries_ProcessesQueriesWithoutMappings',
        'AnalyzeQueries_RendersMatchedInputAndOutputTables',
        'AnalyzeQueries_ClassifiesModificationTargetsByRole',
        'AnalyzeQueries_UsesOutputTwoNameForExactMissingReference',
        'AnalyzeQueries_RendersSupportedTablesDuringPartialFallback',
        'AnalyzeQueries_SortsOutputTwoByCompositeTableNumber',
        'AnalyzeQueries_LeavesOutputTwoHeaderOnlyOnFallback',
        'AnalyzeQueries_WritesWithSubqueriesInsideOut',
        'AnalyzeQueries_PreservesLeadingApostropheInOutput',
        'AnalyzeQueries_DisablesWrappingAfterWritingLongText',
        'AnalyzeQueries_RendersDeeplyNestedCaseConditions',
        'AnalyzeQueries_NormalizesInvisibleOutputWhitespace',
        'AnalyzeQueries_ResolvesQualifiedStarAndMatchingAlias',
        'AnalyzeQueries_QualifiesUnqualifiedSelectColumns',
        'AnalyzeQueries_QualifiesStandaloneColumnThroughTableName',
        'AnalyzeQueries_WritesQualifiedReplacementValuesOnce',
        'AnalyzeQueries_MapsQualifiedReplacementsToMultilineRows',
        'AnalyzeQueries_PreservesReplacementValuesOnParserFallback',
        'AnalyzeQueries_FocusesSyntaxFallbackSqlRow',
        'AnalyzeQueries_RendersInsertSelectWithoutColumnList',
        'AnalyzeQueries_ResolvesMatchingTemporaryTableDefinition',
        'AnalyzeQueries_PreservesUnmatchedTemporaryTableDefinition',
        'AnalyzeQueries_SeparatesTransferExpressionsFromColumns',
        'AnalyzeQueries_RendersWrappedUpdateCaseAsTransferMethod',
        'AnalyzeQueries_HandlesSyntaxCharactersInFieldNames',
        'AnalyzeQueries_UsesStandaloneTableNameForSingleTable',
        'AnalyzeQueries_WritesUnsupportedQueryAsIs',
        'AnalyzeQueries_FramesOnlyTableBody',
        'AnalyzeQueries_RestoresSqlActionButtonsAfterLargeInput',
        'AnalyzeQueries_RestoresApplicationStateAfterOutputError',
        'ClearConfirmMessage_UsesAnalysisResultWording',
        'ClearData_ShowMessageFalseDismissesExistingToast',
        'ClearData_ClearsOutputSheet',
        'ClearData_InitializesOutputTwoAndPreservesTableList',
        'ClearData_RestoresSqlActionButtons'
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
