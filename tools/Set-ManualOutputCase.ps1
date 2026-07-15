param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^SEL-\d{3}$')]
    [string]$CaseId,

    [string]$WorkbookPath,

    [string]$CaseDataPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $repositoryRoot 'SqlAnalysisFormatter.xlsm'
}
if ([string]::IsNullOrWhiteSpace($CaseDataPath)) {
    $CaseDataPath = Join-Path $repositoryRoot 'tests\ManualOutputCases.json'
}

$WorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$CaseDataPath = (Resolve-Path -LiteralPath $CaseDataPath).Path
$caseData = Get-Content -LiteralPath $CaseDataPath -Raw -Encoding UTF8 | ConvertFrom-Json
$matchingCases = @($caseData.cases | Where-Object { $_.id -eq $CaseId })
if ($matchingCases.Count -ne 1) {
    throw "ケースIDが一意に見つからない: $CaseId"
}

$targetCase = $matchingCases[0]
$excel = $null
$workbook = $null
$sqlSheet = $null
$definitionSheet = $null
$outputSheet = $null
$outputRange = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3

    $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $false)
    if ($workbook.ReadOnly) {
        throw 'マクロブックが読み取り専用で開かれた'
    }
    if (-not $workbook.HasVBProject) {
        throw 'VBAプロジェクトが見つからない'
    }

    $sqlSheet = $workbook.Worksheets.Item('SQL解析')
    $definitionSheet = $workbook.Worksheets.Item('変換定義')
    $outputSheet = $workbook.Worksheets.Item('アウトプット')

    $sqlSheet.Range('A2:CL1000').ClearContents() | Out-Null
    # 数値だけのSQL行もA5M2の空白を含む文字列として保持する。
    $sqlSheet.Range('A2:A1000').NumberFormat = '@'
    for ($index = 0; $index -lt $targetCase.sql_lines.Count; $index++) {
        # A5M2が付与した行末空白も入力データとして保持する。
        $sqlSheet.Cells.Item($index + 2, 1).Value2 = [string]$targetCase.sql_lines[$index]
    }

    $definitionSheet.Range('A2:D1000').ClearContents() | Out-Null
    for ($row = 0; $row -lt $targetCase.definitions.Count; $row++) {
        $definition = $targetCase.definitions[$row]
        $definitionSheet.Cells.Item($row + 2, 1).Value2 = [string]$definition.table_id
        $definitionSheet.Cells.Item($row + 2, 2).Value2 = [string]$definition.table_name_ja
        $definitionSheet.Cells.Item($row + 2, 3).Value2 = [string]$definition.field_id
        $definitionSheet.Cells.Item($row + 2, 4).Value2 = [string]$definition.field_name_ja
    }

    $outputSheet.Range('A1:CL200').Clear() | Out-Null
    $outputSheet.Columns.Item('A:CL').ColumnWidth = 1.14
    $outputSheet.Rows.Item('1:200').RowHeight = 13.5
    $outputRange = $outputSheet.Range('A1:CL200')
    $outputRange.Font.Name = 'ＭＳ ゴシック'
    $outputRange.Font.Size = 9
    $outputRange.WrapText = $false

    $workbook.Activate() | Out-Null
    $outputSheet.Activate() | Out-Null
    $outputSheet.Range('A1').Select() | Out-Null
    $excel.ActiveWindow.DisplayGridlines = $false

    $workbook.Save()
    Write-Output "設定完了: $($targetCase.id) $($targetCase.title)"
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($true)
    }
    if ($excel -ne $null) {
        $excel.Quit()
    }

    foreach ($comObject in @($outputRange, $outputSheet, $definitionSheet, $sqlSheet, $workbook, $excel)) {
        if ($comObject -ne $null -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
        }
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
