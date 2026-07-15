param(
    [string]$ParserExePath = 'dist\parser\SqlAnalysisFormatter.Parser.exe',
    [string]$CaseId = '',
    [switch]$MeasurePerformance
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$workbookPath = Join-Path $repoRoot 'SqlAnalysisFormatter.xlsm'
$expectationPath = Join-Path $repoRoot 'tests\SqlAnalysisFormatter.OutputExpectations.xlsx'
$fixturePath = Join-Path $repoRoot 'tests\OutputReportCases.json'
$mainModulePath = Join-Path $repoRoot 'src\vba\SqlAnalysisFormatter.bas'
$goldenTestModulePath = Join-Path $repoRoot 'src\vba\SqlAnalysisFormatterGoldenTests.bas'
$tempWorkbookPath = Join-Path $env:TEMP ('SqlAnalysisFormatter_Golden_' + [guid]::NewGuid().ToString('N') + '.xlsm')
$previousParserExePath = $env:SQL_ANALYSIS_FORMATTER_PARSER_EXE
$env:SQL_ANALYSIS_FORMATTER_PARSER_EXE = (Resolve-Path (Join-Path $repoRoot $ParserExePath))
$totalTimer = [System.Diagnostics.Stopwatch]::StartNew()
$phaseMilliseconds = [ordered]@{
    Reset = 0.0
    Input = 0.0
    Analyze = 0.0
    Values = 0.0
    Formats = 0.0
}

# COM参照を明示的に解放する
function Release-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null
    }
}

# 値が異なる場合はケースとセル番地を付けて停止する
function Assert-Equal {
    param(
        [string]$CaseId,
        [string]$Location,
        [object]$Expected,
        [object]$Actual
    )

    if ([string]$Expected -cne [string]$Actual) {
        throw "$CaseId $Location expected=[$Expected] actual=[$Actual]"
    }
}

Copy-Item -LiteralPath $workbookPath -Destination $tempWorkbookPath -Force
$fixture = Get-Content -LiteralPath $fixturePath -Encoding UTF8 -Raw | ConvertFrom-Json
$testCases = @($fixture.cases)
if (-not [string]::IsNullOrWhiteSpace($CaseId)) {
    $testCases = @($testCases | Where-Object id -eq $CaseId)
    if ($testCases.Count -eq 0) {
        throw "Output case not found: $CaseId"
    }
}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

try {
    $workbook = $excel.Workbooks.Open($tempWorkbookPath)
    $expectationBook = $excel.Workbooks.Open($expectationPath, 0, $true)
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $excel.Calculation = -4135
    $components = $workbook.VBProject.VBComponents
    foreach ($moduleName in @('SqlAnalysisFormatter', 'SqlAnalysisFormatterGoldenTests')) {
        try {
            $components.Remove($components.Item($moduleName))
        } catch {
        }
    }
    $components.Import($mainModulePath) | Out-Null
    $components.Import($goldenTestModulePath) | Out-Null

    $definitionSheet = $workbook.Worksheets.Item('変換定義')
    $sqlSheet = $workbook.Worksheets.Item('SQL解析')
    $outputSheet = $workbook.Worksheets.Item('アウトプット')
    $expectationBookName = [string]$expectationBook.Name
    $outputSheetName = [string]$outputSheet.Name
    $definitionClearLastRow = [Math]::Max(2, [int]$definitionSheet.UsedRange.Rows.Count)
    $sqlClearLastRow = [Math]::Max(2, [int]$sqlSheet.UsedRange.Rows.Count)
    $caseIndex = 0

    foreach ($testCase in $testCases) {
        $caseIndex++
        $phaseTimer = [System.Diagnostics.Stopwatch]::StartNew()
        $definitionSheet.Range("A2:D$definitionClearLastRow").ClearContents() | Out-Null
        $sqlSheet.Range("A2:Z$sqlClearLastRow").ClearContents() | Out-Null
        $phaseMilliseconds.Reset += $phaseTimer.Elapsed.TotalMilliseconds

        $phaseTimer.Restart()
        $definitionRow = 2
        $tableProperties = @($testCase.tables.PSObject.Properties)
        if ($tableProperties.Count -eq 0) {
            $tableProperties = @([pscustomobject]@{ Name = '__dummy__'; Value = '未使用' })
        }
        foreach ($table in $tableProperties) {
            $definitionSheet.Cells.Item($definitionRow, 1).Value2 = [string]$table.Name
            $definitionSheet.Cells.Item($definitionRow, 2).Value2 = [string]$table.Value
            $definitionSheet.Cells.Item($definitionRow, 3).Value2 = '__unused__'
            $definitionSheet.Cells.Item($definitionRow, 4).Value2 = '未使用'
            $definitionRow++
        }

        $sqlRow = 2
        foreach ($line in $testCase.sql_lines) {
            $sqlSheet.Cells.Item($sqlRow, 1).Value2 = [string]$line
            $sqlRow++
        }
        $definitionClearLastRow = [Math]::Max(2, $definitionRow - 1)
        $sqlClearLastRow = [Math]::Max(2, $sqlRow - 1)
        $phaseMilliseconds.Input += $phaseTimer.Elapsed.TotalMilliseconds

        $phaseTimer.Restart()
        $excel.Run("'$tempWorkbookPath'!AnalyzeQueries", $false) | Out-Null
        $phaseMilliseconds.Analyze += $phaseTimer.Elapsed.TotalMilliseconds

        $phaseTimer.Restart()
        $expectedSheet = $expectationBook.Worksheets.Item([string]$testCase.id)
        $expectedRowCount = $expectedSheet.UsedRange.Rows.Count
        $expectedRange = $expectedSheet.Range("A1:CL$expectedRowCount")
        $actualRange = $outputSheet.Range("A1:CL$expectedRowCount")
        $expectedValues = $expectedRange.Value2
        $actualValues = $actualRange.Value2

        for ($row = 1; $row -le $expectedRowCount; $row++) {
            for ($column = 1; $column -le 90; $column++) {
                Assert-Equal ([string]$testCase.id) ("R{0}C{1}" -f $row, $column) $expectedValues[$row, $column] $actualValues[$row, $column]
            }
        }
        $phaseMilliseconds.Values += $phaseTimer.Elapsed.TotalMilliseconds

        $phaseTimer.Restart()
        $formatFailure = [string]$excel.Run(
            "'$tempWorkbookPath'!CompareOutputGoldenFormat",
            [string]$testCase.id,
            $expectationBookName,
            [string]$testCase.id,
            $outputSheetName,
            [int]$expectedRowCount,
            [bool]($caseIndex -eq 1))
        if (-not [string]::IsNullOrEmpty($formatFailure)) {
            throw $formatFailure
        }
        if ($caseIndex -eq 1) {
            $originalWrapText = $outputSheet.Cells.Item(1, 1).WrapText
            $outputSheet.Cells.Item(1, 1).WrapText = -not [bool]$originalWrapText
            $comparatorSelfTest = [string]$excel.Run(
                "'$tempWorkbookPath'!CompareOutputGoldenFormat",
                [string]$testCase.id,
                $expectationBookName,
                [string]$testCase.id,
                $outputSheetName,
                [int]$expectedRowCount,
                $false)
            $outputSheet.Cells.Item(1, 1).WrapText = $originalWrapText
            if ($comparatorSelfTest -notlike "*$($testCase.id) A1 wrap*") {
                throw "Output format comparator did not detect the intentional mismatch."
            }
        }
        $phaseMilliseconds.Formats += $phaseTimer.Elapsed.TotalMilliseconds

        Release-ComObject $actualRange
        Release-ComObject $expectedRange
        Release-ComObject $expectedSheet
        if ($caseIndex % 10 -eq 0 -or $caseIndex -eq $testCases.Count) {
            Write-Output ("Output golden progress: {0}/{1}" -f $caseIndex, $testCases.Count)
        }
    }

    Write-Output ("Output golden tests passed: {0} cases." -f $testCases.Count)
    if ($MeasurePerformance) {
        $totalTimer.Stop()
        foreach ($phase in $phaseMilliseconds.GetEnumerator()) {
            Write-Output ("Output golden timing: {0}={1:N0} ms" -f $phase.Key, $phase.Value)
        }
        Write-Output ("Output golden timing: Total={0:N0} ms" -f $totalTimer.Elapsed.TotalMilliseconds)
    }
} finally {
    if ($null -ne $expectationBook) {
        $expectationBook.Close($false) | Out-Null
    }
    if ($null -ne $workbook) {
        $workbook.Close($false) | Out-Null
    }
    $excel.Quit()
    foreach ($comObject in @($outputSheet, $sqlSheet, $definitionSheet, $components, $expectationBook, $workbook, $excel)) {
        Release-ComObject $comObject
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Remove-Item -LiteralPath $tempWorkbookPath -Force -ErrorAction SilentlyContinue
    $env:SQL_ANALYSIS_FORMATTER_PARSER_EXE = $previousParserExePath
}
