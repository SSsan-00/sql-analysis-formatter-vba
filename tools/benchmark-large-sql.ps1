param(
    [ValidateRange(1, 10000)]
    [int]$ItemCount = 2000,

    [ValidateRange(0, 2000)]
    [int]$MappingCount = 200,

    [string]$ParserExePath = 'dist\parser\SqlAnalysisFormatter.Parser.exe',

    [switch]$UseEmbeddedMainModule,

    [switch]$SkipAssertions
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
$parserPath = (Resolve-Path (Join-Path $repoRoot $ParserExePath)).Path
$tempWorkbookPath = Join-Path $env:TEMP (
    'SqlAnalysisFormatter_LargeSql_' + [guid]::NewGuid().ToString('N') + '.xlsm')
$tempSqlPath = Join-Path $env:TEMP (
    'SqlAnalysisFormatter_LargeSql_' + [guid]::NewGuid().ToString('N') + '.sql')
$tempMappingPath = Join-Path $env:TEMP (
    'SqlAnalysisFormatter_LargeSql_' + [guid]::NewGuid().ToString('N') + '.txt')
$tempPlanPath = Join-Path $env:TEMP (
    'SqlAnalysisFormatter_LargeSql_' + [guid]::NewGuid().ToString('N') + '.plan')
$previousParserPath = $env:SQL_ANALYSIS_FORMATTER_PARSER_EXE

function Release-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject -and
        [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject) | Out-Null
    }
}

function Read-ButtonState {
    param([object]$Worksheet)

    $states = @()
    foreach ($shape in $Worksheet.Shapes) {
        if ($shape.Name -in @('btnAnalyzeQueries', 'btnClearData')) {
            $states += [pscustomobject]@{
                Name = [string]$shape.Name
                Visible = [int]$shape.Visible
                Placement = [int]$shape.Placement
                Top = [double]$shape.Top
                Left = [double]$shape.Left
                Width = [double]$shape.Width
                Height = [double]$shape.Height
            }
        }
        Release-ComObject $shape
    }
    return $states
}

Copy-Item -LiteralPath $workbookPath -Destination $tempWorkbookPath -Force
$env:SQL_ANALYSIS_FORMATTER_PARSER_EXE = $parserPath

$excel = $null
$workbooks = $null
$workbook = $null
$definitionSheet = $null
$sqlSheet = $null
$outputSheet = $null
$vbProject = $null
$components = $null
$existingComponent = $null
$importedComponent = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1

    $workbooks = $excel.Workbooks
    $workbook = $workbooks.Open($tempWorkbookPath)
    Release-ComObject $workbooks
    $workbooks = $null

    if (-not $UseEmbeddedMainModule) {
        $vbProject = $workbook.VBProject
        $components = $vbProject.VBComponents
        foreach ($componentName in @($productionComponents.Name)) {
            try {
                $existingComponent = $components.Item($componentName)
                $components.Remove($existingComponent)
            } catch {
            } finally {
                Release-ComObject $existingComponent
                $existingComponent = $null
            }
        }
        foreach ($productionComponent in $productionComponents) {
            $importedComponent = $components.Import($productionComponent.Path)
            Release-ComObject $importedComponent
            $importedComponent = $null
        }
    }

    $excel.Run("'$tempWorkbookPath'!SetupWorkbook") | Out-Null
    $definitionSheet = $workbook.Worksheets.Item(1)
    $sqlSheet = $workbook.Worksheets.Item(2)
    $outputSheet = $workbook.Worksheets.Item(3)

    $definitionSheet.Range('A2:D5000').ClearContents() | Out-Null
    $sqlSheet.Range('A2:Z10000').ClearContents() | Out-Null

    $mappingLines = [string[]]::new($MappingCount + 1)
    $mappingLines[0] = "SAF_MAPPINGS`t2"
    if ($MappingCount -gt 0) {
        $mappingValues = [object[,]]::new($MappingCount, 4)
        for ($index = 0; $index -lt $MappingCount; $index++) {
            $number = $index + 1
            $mappingRowNumber = $number + 1
            $mappingValues[$index, 0] = 'tb1'
            $mappingValues[$index, 1] = 'User'
            $mappingValues[$index, 2] = 'col' + $number
            $mappingValues[$index, 3] = 'Field' + $number
            $mappingLines[$index + 1] = "M`ttb1`tUser`tcol$number`tField$number`t__SAF_FIELD_R$($mappingRowNumber.ToString('000000'))__"
        }
        $definitionSheet.Range('A2').Resize($MappingCount, 4).Value2 = $mappingValues
    }

    $sqlValues = [object[,]]::new($ItemCount + 2, 1)
    $sqlLines = [string[]]::new($ItemCount + 2)
    $parserSqlLines = [string[]]::new($ItemCount + 2)
    $sqlValues[0, 0] = 'SELECT'
    $sqlLines[0] = 'SELECT'
    $parserSqlLines[0] = 'SELECT'
    for ($index = 1; $index -le $ItemCount; $index++) {
        $mappingNumber = if ($MappingCount -gt 0) {
            (($index - 1) % $MappingCount) + 1
        }
        else {
            $index
        }
        $prefix = if ($index -eq 1) { '    ' } else { '    , ' }
        $sqlValues[$index, 0] = $prefix + 'tb1.col' + $mappingNumber + ' AS item' + $index
        $sqlLines[$index] = [string]$sqlValues[$index, 0]
        if ($MappingCount -gt 0) {
            $mappingRowNumber = $mappingNumber + 1
            $parserFieldId = '__SAF_FIELD_R' + $mappingRowNumber.ToString('000000') + '__'
            $parserSqlLines[$index] = $prefix + 'tb1.' + $parserFieldId + ' AS item' + $index
        }
        else {
            $parserSqlLines[$index] = $sqlLines[$index]
        }
    }
    $sqlValues[($ItemCount + 1), 0] = 'FROM users AS tb1;'
    $sqlLines[$ItemCount + 1] = 'FROM users AS tb1;'
    $parserSqlLines[$ItemCount + 1] = 'FROM users AS tb1;'
    $sqlSheet.Range('A2').Resize($ItemCount + 2, 1).Value2 = $sqlValues

    $utf8WithoutBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines($tempSqlPath, $parserSqlLines, $utf8WithoutBom)
    [System.IO.File]::WriteAllLines($tempMappingPath, $mappingLines, $utf8WithoutBom)
    $parserTimer = [System.Diagnostics.Stopwatch]::StartNew()
    & $parserPath --input $tempSqlPath --mappings $tempMappingPath --output $tempPlanPath --format vba-plan
    $parserExitCode = $LASTEXITCODE
    $parserTimer.Stop()
    if ($parserExitCode -ne 0) {
        throw "Parser benchmark failed with exit code $parserExitCode."
    }
    $planLines = [System.IO.File]::ReadAllLines($tempPlanPath)
    $planSectionCount = @($planLines | Where-Object { $_.StartsWith("S`t") }).Count
    $planCellCount = @($planLines | Where-Object { $_.StartsWith("C`t") }).Count
    $planColumns = @(
        $planLines |
            Where-Object { $_.StartsWith("C`t") } |
            ForEach-Object { [int]($_ -split "`t")[2] } |
            Sort-Object -Unique)

    $beforeButtons = @(Read-ButtonState $sqlSheet)
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $excel.Run("'$tempWorkbookPath'!AnalyzeQueries", $false) | Out-Null
    $timer.Stop()
    $afterButtons = @(Read-ButtonState $sqlSheet)
    $displayDrawingObjects = [int]$workbook.DisplayDrawingObjects
    $outputRows = [int]$outputSheet.UsedRange.Rows.Count

    if (-not $SkipAssertions) {
        if ($displayDrawingObjects -ne -4104) {
            throw "Workbook drawing objects are not displayed after AnalyzeQueries."
        }
        if ($afterButtons.Count -ne 2) {
            throw "Expected two SQL action buttons after AnalyzeQueries; found $($afterButtons.Count)."
        }
        $expectedButtonLeft = @{
            btnAnalyzeQueries = [double]$sqlSheet.Columns.Item(5).Left
            btnClearData = [double]$sqlSheet.Columns.Item(5).Left + 82
        }
        foreach ($buttonName in @('btnAnalyzeQueries', 'btnClearData')) {
            $buttonState = @($afterButtons | Where-Object Name -eq $buttonName)
            if ($buttonState.Count -ne 1) {
                throw "SQL action button is missing or duplicated: $buttonName"
            }
            $button = $buttonState[0]
            if ($button.Visible -ne -1 -or $button.Placement -ne 3) {
                throw "SQL action button is hidden or not free-floating: $buttonName"
            }
            if ([math]::Abs($button.Top - 2) -gt 0.1 -or
                [math]::Abs($button.Left - $expectedButtonLeft[$buttonName]) -gt 0.1 -or
                [math]::Abs($button.Width - 72) -gt 0.1 -or
                [math]::Abs($button.Height - 24) -gt 0.1) {
                throw "SQL action button position or size is invalid: $buttonName"
            }
        }
        $expectedOutputRows = [int](($planLines[0] -split "`t")[2])
        if ($outputRows -ne $expectedOutputRows) {
            throw "Output row count expected $expectedOutputRows; found $outputRows."
        }
    }

    [pscustomobject]@{
        ItemCount = $ItemCount
        MappingCount = $MappingCount
        ParserMilliseconds = [math]::Round($parserTimer.Elapsed.TotalMilliseconds, 1)
        ElapsedMilliseconds = [math]::Round($timer.Elapsed.TotalMilliseconds, 1)
        PlanSectionCount = $planSectionCount
        PlanCellCount = $planCellCount
        PlanColumnCount = $planColumns.Count
        PlanColumns = $planColumns -join ','
        BeforeButtonCount = $beforeButtons.Count
        AfterButtonCount = $afterButtons.Count
        DisplayDrawingObjects = $displayDrawingObjects
        OutputRows = $outputRows
        SqlUsedColumns = [int]$sqlSheet.UsedRange.Columns.Count
        SqlHeaderRowHeight = [double]$sqlSheet.Rows.Item(1).RowHeight
    } | Format-List

    'Before buttons:'
    $beforeButtons | Format-Table -AutoSize
    'After buttons:'
    $afterButtons | Format-Table -AutoSize
}
finally {
    try {
        if ($null -ne $workbook) {
            $workbook.Close($false)
        }
    }
    catch {
        Write-Warning "Could not close benchmark workbook: $($_.Exception.Message)"
    }
    try {
        if ($null -ne $excel) {
            $excel.Quit()
        }
    }
    catch {
        Write-Warning "Could not quit benchmark Excel instance: $($_.Exception.Message)"
    }
    $env:SQL_ANALYSIS_FORMATTER_PARSER_EXE = $previousParserPath

    foreach ($comObject in @(
        $definitionSheet,
        $sqlSheet,
        $outputSheet,
        $importedComponent,
        $existingComponent,
        $components,
        $vbProject,
        $workbook,
        $excel)) {
        try {
            Release-ComObject $comObject
        }
        catch {
            Write-Warning "Could not release benchmark COM object: $($_.Exception.Message)"
        }
    }
    try {
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    catch {
        Write-Warning "Could not complete benchmark COM cleanup: $($_.Exception.Message)"
    }

    foreach ($tempPath in @(
        $tempWorkbookPath,
        $tempSqlPath,
        $tempMappingPath,
        $tempPlanPath)) {
        try {
            if (Test-Path -LiteralPath $tempPath) {
                Remove-Item -LiteralPath $tempPath -Force
            }
        }
        catch {
            Write-Warning "Could not delete benchmark temp file '$tempPath': $($_.Exception.Message)"
        }
    }
}
