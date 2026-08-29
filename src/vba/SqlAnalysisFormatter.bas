Attribute VB_Name = "SqlAnalysisFormatter"
Option Explicit

Private Const REF_LEGACY_SHEET As String = "Sheet1"
Private Const SQL_LEGACY_SHEET As String = "Sheet2"

Private Const COL_TABLE_ID As Long = 1
Private Const COL_TABLE_NAME As Long = 2
Private Const COL_FIELD_ID As Long = 3
Private Const COL_FIELD_NAME As Long = 4

Private Const COL_SQL As Long = 1
Private Const COL_RESULT As Long = 2
Private Const COL_REPLACEMENT As Long = 3

Private Const OUTPUT_LAST_COLUMN As Long = 90
Private Const OUTPUT_TWO_LAST_COLUMN As Long = 100
Private Const OUTPUT_COLUMN_WIDTH As Double = 1.14
Private Const OUTPUT_ROW_HEIGHT As Double = 13.5
Private Const OUTPUT_FILL_COLOR As Long = &HEFCEF2
Private Const MISSING_NAME_FILL_COLOR As Long = &HD6E4FC

' ブックのシート名、見出し、操作ボタンを初期化
Public Sub SetupWorkbook()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim wsOutputTwo As Worksheet
    Dim wsTableList As Worksheet

    Set wsRef = ResolveOrCreateSheet(ReferenceSheetName(), REF_LEGACY_SHEET, 1)
    Set wsSql = ResolveOrCreateSheet(SqlSheetName(), SQL_LEGACY_SHEET, 2)
    Set wsOutput = ResolveOrCreateOutputSheet()
    Set wsOutputTwo = ResolveOrCreateNamedSheet(OutputSheetTwoName())
    Set wsTableList = ResolveOrCreateNamedSheet(TableListSheetName())

    MoveWorksheetToIndex wsRef, 1
    MoveWorksheetToIndex wsSql, 2
    MoveWorksheetToIndex wsOutput, 3
    MoveWorksheetToIndex wsOutputTwo, 4
    MoveWorksheetToIndex wsTableList, 5

    RestoreHeaders wsRef, wsSql
    ApplyMissingNameConditionalFormatting wsRef
    ApplyOutputSheetLayout wsOutput
    EnsureOutputTwoStructure wsOutputTwo
    ApplyTableListLayout wsTableList
    EnsureSqlActionButtons wsSql
    InstallOutputButton wsOutput, "btnCopyOutput", "CopyOutput", OUTPUT_LAST_COLUMN
    InstallOutputButton wsOutputTwo, "btnCopyOutputTwo", "CopyOutputTwo", OUTPUT_TWO_LAST_COLUMN
End Sub

' SQL解析シートのA列を変換し、B列以降へ結果を出力
Public Sub AnalyzeQueries(Optional ByVal showMessage As Boolean = True)
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim wsOutputTwo As Worksheet
    Dim wsTableList As Worksheet
    Dim qualifiedMap As Object
    Dim standaloneMap As Object
    Dim qualifiedParserMap As Object
    Dim standaloneParserMap As Object
    Dim qualifiedRegexes As Object
    Dim standaloneRegexes As Object
    Dim qualifiedKeys As Variant
    Dim standaloneKeys As Variant
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim sourceText As String
    Dim convertedText As String
    Dim convertedQueryText As String
    Dim parserText As String
    Dim parserQueryText As String
    Dim fallbackReason As String
    Dim fallbackStartLine As Long
    Dim fallbackEndLine As Long
    Dim replacementValues As Object
    Dim replacementValuesByRow As Object
    Dim queryLineRows As Collection
    Dim qualifications As Collection
    Dim transformedQueryLines As Collection
    Dim finalReplacementValues As Collection
    Dim tableNameReferences As Collection
    Dim inputTableIds As Collection
    Dim outputTableIds As Collection
    Dim hasTransformedQueryData As Boolean
    Dim hasTableNameReferenceData As Boolean
    Dim tableMaster As Object
    Dim duplicateTableIds As Collection
    Dim sqlValues As Variant
    Dim convertedValues() As Variant
    Dim convertedLines() As String
    Dim parserLines() As String
    Dim inputRowCount As Long
    Dim inputRowIndex As Long
    Dim nonEmptyLineCount As Long
    Dim maxReplacementCount As Long
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim previousStatusBar As Variant
    Dim applicationStateCaptured As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim analysisStage As String
    Dim analysisOutputReset As Boolean

    On Error GoTo AnalyzeFail
    SqlAnalysisToastManager.DismissToast

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    previousStatusBar = Application.StatusBar
    applicationStateCaptured = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "SQL analysis is running..."

    analysisStage = "resolve workbook sheets"
    Set wsRef = GetReferenceSheet()
    Set wsSql = GetSqlSheet()
    analysisStage = "restore SQL action buttons"
    EnsureSqlActionButtons wsSql
    Set wsOutput = GetOutputSheet(False)
    Set wsOutputTwo = GetOutputTwoSheet(False)
    Set wsTableList = GetTableListSheet(False)
    Set qualifiedMap = CreateTextDictionary()
    Set standaloneMap = CreateTextDictionary()
    Set qualifiedParserMap = CreateTextDictionary()
    Set standaloneParserMap = CreateTextDictionary()
    Set qualifiedRegexes = CreateTextDictionary()
    Set standaloneRegexes = CreateTextDictionary()
    Set replacementValuesByRow = CreateTextDictionary()
    Set queryLineRows = New Collection
    Set qualifications = New Collection
    Set tableNameReferences = New Collection
    Set inputTableIds = New Collection
    Set outputTableIds = New Collection
    analysisStage = "load table definitions"
    LoadTableMaster wsTableList, tableMaster, duplicateTableIds

    analysisStage = "load conversion definitions"
    LoadMappings wsRef, qualifiedMap, standaloneMap, qualifiedParserMap, standaloneParserMap
    qualifiedKeys = SortedKeysByLengthDesc(qualifiedMap)
    standaloneKeys = SortedKeysByLengthDesc(standaloneMap)

    analysisStage = "prepare output sheets"
    lastRow = LastUsedRowInColumn(wsSql, COL_SQL)
    ClearAnalyzeOutput wsSql, lastRow
    ClearOutputSheet wsOutput, False
    analysisOutputReset = True

    analysisStage = "convert SQL input"
    If lastRow >= 2 Then
        inputRowCount = lastRow - 1
        sqlValues = wsSql.Range( _
            wsSql.Cells(2, COL_SQL), _
            wsSql.Cells(lastRow, COL_SQL)).Value2
        ReDim convertedValues(1 To inputRowCount, 1 To 1)
        ReDim convertedLines(0 To inputRowCount - 1)
        ReDim parserLines(0 To inputRowCount - 1)

        For inputRowIndex = 1 To inputRowCount
            rowNumber = inputRowIndex + 1
            If inputRowCount = 1 Then
                sourceText = CStr(sqlValues)
            Else
                sourceText = CStr(sqlValues(inputRowIndex, 1))
            End If
            If Len(sourceText) > 0 Then
                Set replacementValues = CreateTextDictionary()
                convertedText = ApplyMappings( _
                    sourceText, qualifiedMap, qualifiedKeys, standaloneMap, standaloneKeys, _
                    replacementValues, True, qualifiedRegexes, standaloneRegexes)
                parserText = ApplyMappings( _
                    sourceText, qualifiedParserMap, qualifiedKeys, _
                    standaloneParserMap, standaloneKeys, Nothing, False, _
                    qualifiedRegexes, standaloneRegexes)
                convertedValues(inputRowIndex, 1) = OutputTextValue(convertedText)
                Set replacementValuesByRow(CStr(rowNumber)) = replacementValues
                convertedLines(nonEmptyLineCount) = convertedText
                parserLines(nonEmptyLineCount) = parserText
                nonEmptyLineCount = nonEmptyLineCount + 1
                AddQueryLineRows queryLineRows, rowNumber, parserText
            End If
        Next inputRowIndex

        With wsSql.Range(wsSql.Cells(2, COL_RESULT), wsSql.Cells(lastRow, COL_RESULT))
            .NumberFormat = "@"
            .Value = convertedValues
        End With

        If nonEmptyLineCount > 0 Then
            ReDim Preserve convertedLines(0 To nonEmptyLineCount - 1)
            ReDim Preserve parserLines(0 To nonEmptyLineCount - 1)
            convertedQueryText = Join(convertedLines, vbCrLf)
            parserQueryText = Join(parserLines, vbCrLf)
        End If
    End If

    analysisStage = "parse SQL and render output"
    If Len(convertedQueryText) > 0 Then
        If Not TryWriteExternalOutputPlan( _
            wsOutput, wsRef, parserQueryText, qualifications, inputTableIds, outputTableIds, _
            fallbackReason, fallbackStartLine, fallbackEndLine, queryLineRows.Count, _
            transformedQueryLines, finalReplacementValues, hasTransformedQueryData, _
            tableNameReferences, hasTableNameReferenceData) Then
            ClearOutputTwoSheet wsOutputTwo
            WriteFallbackOutput wsOutput, convertedQueryText, fallbackReason
        Else
            If hasTransformedQueryData Then
                ApplyTransformedQueryLines _
                    wsSql, lastRow, queryLineRows, transformedQueryLines
                ApplyFinalReplacementValues _
                    replacementValuesByRow, queryLineRows, finalReplacementValues
            Else
                ApplyReplacementQualifications _
                    replacementValuesByRow, queryLineRows, qualifications
            End If
            RenderOutputTwo wsOutputTwo, inputTableIds, outputTableIds, tableMaster
            ApplyOutputTwoNamesToMissingTableDisplays _
                wsOutput, inputTableIds, outputTableIds, tableMaster, _
                tableNameReferences, hasTableNameReferenceData
        End If
    Else
        ClearOutputTwoSheet wsOutputTwo
        ApplyOutputSheetLayout wsOutput
    End If

    analysisStage = "write SQL analysis results"
    WriteReplacementValuesBatch _
        wsSql, lastRow, replacementValuesByRow, maxReplacementCount

    wsSql.Range(wsSql.Cells(1, COL_RESULT), wsSql.Cells(MaxLong(lastRow, 1), COL_RESULT)).WrapText = False
    SetReplacementColumnsWrapText _
        wsSql, False, _
        MaxLong(26, COL_REPLACEMENT + maxReplacementCount - 1), _
        MaxLong(lastRow, 1)
    analysisStage = "finalize SQL analysis sheet"
    RestoreFindSearchOrderByRows wsSql
    EnsureSqlActionButtons wsSql

AnalyzeCleanUp:
    On Error Resume Next
    If errorNumber <> 0 And analysisOutputReset Then
        If Not wsOutputTwo Is Nothing Then ClearOutputTwoSheet wsOutputTwo
    End If
    If Not wsSql Is Nothing Then EnsureSqlActionButtons wsSql
    If applicationStateCaptured Then
        Application.StatusBar = previousStatusBar
        Application.EnableEvents = previousEnableEvents
        Application.ScreenUpdating = previousScreenUpdating
    End If
    On Error GoTo 0

    If errorNumber <> 0 Then
        Err.Raise errorNumber, errorSource, errorDescription
    End If
    If Len(fallbackReason) > 0 And fallbackStartLine > 0 Then
        FocusFallbackQueryLine wsSql, queryLineRows, fallbackStartLine
    End If
    If showMessage And duplicateTableIds.Count > 0 Then
        MsgBox DuplicateTableWarningMessage(duplicateTableIds), vbExclamation
    End If
    If showMessage Then
        If Len(fallbackReason) > 0 Then
            MsgBox AnalyzeFallbackMessage(fallbackReason), vbExclamation
        Else
            SqlAnalysisToastManager.ShowToast AnalyzeDoneMessage()
        End If
    End If
    Exit Sub

AnalyzeFail:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = analysisStage & ": " & Err.Description
    Resume AnalyzeCleanUp
End Sub

' 確認後、入力シートの2行目以降とアウトプットシートをクリア
Public Sub ClearData(Optional ByVal showMessage As Boolean = True)
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim wsOutputTwo As Worksheet
    Dim wsTableList As Worksheet
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    SqlAnalysisToastManager.DismissToast

    If showMessage Then
        If MsgBox(ClearConfirmMessage(), vbQuestion + vbYesNo + vbDefaultButton2, ConfirmTitle()) <> vbYes Then
            Exit Sub
        End If
    End If

    On Error GoTo ClearFail

    Set wsRef = GetReferenceSheet()
    Set wsSql = GetSqlSheet()
    Set wsOutput = GetOutputSheet(False)
    Set wsOutputTwo = GetOutputTwoSheet(False)
    Set wsTableList = GetTableListSheet()

    ClearRowsBelowHeader wsRef, COL_FIELD_NAME
    ClearRowsBelowHeader wsSql, COL_REPLACEMENT
    ClearOutputSheet wsOutput
    ClearOutputTwoSheet wsOutputTwo
    RestoreHeaders wsRef, wsSql
    ResetSheetViewPosition wsRef
    ResetSheetViewPosition wsSql
    ResetSheetViewPosition wsOutput
    ResetSheetViewPosition wsOutputTwo
    ResetSheetViewPosition wsTableList
    EnsureSqlActionButtons wsSql
    RestoreFindSearchOrderByRows wsSql
    If showMessage Then
        SqlAnalysisToastManager.ShowToast ClearDoneMessage()
    End If
    Exit Sub

ClearFail:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    On Error Resume Next
    If Not wsSql Is Nothing Then EnsureSqlActionButtons wsSql
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

' 指定シートの選択セルとスクロールをA1へ戻し、元のシート表示を維持
Private Sub ResetSheetViewPosition(ByVal ws As Worksheet)
    Dim previousSheet As Object
    Dim previousScreenUpdating As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo RestoreApplicationState
    previousScreenUpdating = Application.ScreenUpdating
    Set previousSheet = ActiveSheet
    Application.ScreenUpdating = False

    ' 選択セルとスクロール位置の変更には対象シートの一時表示が必要
    ws.Activate
    ws.Range("A1").Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1

RestoreApplicationState:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description

    On Error Resume Next
    If Not previousSheet Is Nothing Then previousSheet.Activate
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0

    If errorNumber <> 0 Then
        Err.Raise errorNumber, errorSource, errorDescription
    End If
End Sub

' 変換定義シートから表示用とparser用の変換表を作成
Private Sub LoadMappings( _
    ByVal wsRef As Worksheet, _
    ByVal qualifiedMap As Object, _
    ByVal standaloneMap As Object, _
    ByVal qualifiedParserMap As Object, _
    ByVal standaloneParserMap As Object)

    Dim lastRow As Long
    Dim rowNumber As Long
    Dim tableId As String
    Dim fieldId As String
    Dim fieldName As String
    Dim parserFieldId As String
    Dim mappingValues As Variant

    lastRow = LastUsedRow(wsRef)
    If lastRow < 2 Then Exit Sub

    mappingValues = wsRef.Range( _
        wsRef.Cells(2, COL_TABLE_ID), _
        wsRef.Cells(lastRow, COL_FIELD_NAME)).Value2
    For rowNumber = 2 To lastRow
        tableId = NormalizeKey(mappingValues(rowNumber - 1, COL_TABLE_ID))
        fieldId = NormalizeKey(mappingValues(rowNumber - 1, COL_FIELD_ID))
        fieldName = NormalizeName(mappingValues(rowNumber - 1, COL_FIELD_NAME))

        If Len(fieldId) > 0 And IsUsableJapaneseName(fieldName) Then
            parserFieldId = ParserFieldIdentifier(rowNumber)
            If tableId = "-" Then
                standaloneMap(fieldId) = fieldName
                standaloneParserMap(fieldId) = parserFieldId
            ElseIf Len(tableId) > 0 Then
                qualifiedMap(tableId & "." & fieldId) = tableId & "." & fieldName
                qualifiedParserMap(tableId & "." & fieldId) = tableId & "." & parserFieldId
            End If
        End If
    Next rowNumber
End Sub

' 1行分のSQLへ変換表を適用し、必要な場合だけ変換後値を記録
Private Function ApplyMappings( _
    ByVal sourceText As String, _
    ByVal qualifiedMap As Object, _
    ByVal qualifiedKeys As Variant, _
    ByVal standaloneMap As Object, _
    ByVal standaloneKeys As Variant, _
    ByVal replacementValues As Object, _
    Optional ByVal trackReplacements As Boolean = True, _
    Optional ByVal qualifiedRegexes As Object = Nothing, _
    Optional ByVal standaloneRegexes As Object = Nothing) As String
    Dim resultText As String

    resultText = sourceText

    resultText = ApplyMappingSet( _
        resultText, qualifiedMap, qualifiedKeys, replacementValues, False, _
        trackReplacements, qualifiedRegexes)
    resultText = ApplyMappingSet( _
        resultText, standaloneMap, standaloneKeys, replacementValues, True, _
        trackReplacements, standaloneRegexes)

    ApplyMappings = resultText
End Function

' 指定された1組の変換表を1行分のSQLへ適用
Private Function ApplyMappingSet( _
    ByVal sourceText As String, _
    ByVal mapping As Object, _
    ByVal sortedKeys As Variant, _
    ByVal replacementValues As Object, _
    ByVal standaloneMode As Boolean, _
    ByVal trackReplacements As Boolean, _
    ByVal regexes As Object) As String
    Dim resultText As String
    Dim key As Variant
    Dim searchText As String
    Dim changeCount As Long
    Dim replacementText As String
    Dim firstMatchIndex As Long

    If mapping.Count = 0 Then
        ApplyMappingSet = sourceText
        Exit Function
    End If

    resultText = sourceText
    For Each key In sortedKeys
        searchText = CStr(key)
        If InStr(1, resultText, searchText, vbBinaryCompare) > 0 Then
            replacementText = CStr(mapping(searchText))
            changeCount = 0
            resultText = ReplaceIdentifier( _
                resultText, searchText, replacementText, changeCount, firstMatchIndex, _
                standaloneMode, regexes)
            If changeCount > 0 And trackReplacements Then
                AddReplacementValue replacementValues, replacementText, firstMatchIndex
            End If
        End If
    Next key

    ApplyMappingSet = resultText
End Function

' 識別子単位で文字列を置換し、置換数と初回位置を返却
Private Function ReplaceIdentifier( _
    ByVal sourceText As String, _
    ByVal searchText As String, _
    ByVal replacementText As String, _
    ByRef changeCount As Long, _
    ByRef firstMatchIndex As Long, _
    ByVal standaloneMode As Boolean, _
    ByVal regexes As Object) As String
    Dim re As Object
    Dim matches As Object
    Dim matchItem As Object
    Dim resultText As String
    Dim index As Long
    Dim prefix As String
    Dim suffix As String
    Dim identifierStart As Long

    If Not regexes Is Nothing Then
        If regexes.Exists(searchText) Then Set re = regexes(searchText)
    End If
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.IgnoreCase = False
    ' 識別子の一部一致を避けるため、前後を英数字とアンダースコア以外に限定
    If standaloneMode Then
        re.Pattern = "(^|[^A-Za-z0-9_.])" & EscapeRegexLiteral(searchText) & "([^A-Za-z0-9_.]|$)"
    Else
        re.Pattern = "(^|[^A-Za-z0-9_])" & EscapeRegexLiteral(searchText) & "([^A-Za-z0-9_]|$)"
        End If
        If Not regexes Is Nothing Then Set regexes(searchText) = re
    End If

    Set matches = re.Execute(sourceText)
    changeCount = 0
    firstMatchIndex = -1
    resultText = sourceText

    ' 後方から置換し、FirstIndexのずれを防ぐ
    For index = matches.Count - 1 To 0 Step -1
        Set matchItem = matches.Item(index)
        prefix = CStr(matchItem.SubMatches(0))
        suffix = CStr(matchItem.SubMatches(1))
        identifierStart = matchItem.FirstIndex + Len(prefix) + 1
        If Not (standaloneMode And IsAliasAfterAs(sourceText, identifierStart)) Then
            If firstMatchIndex = -1 Or matchItem.FirstIndex < firstMatchIndex Then
                firstMatchIndex = matchItem.FirstIndex
            End If
            changeCount = changeCount + 1
            resultText = Left$(resultText, matchItem.FirstIndex) _
                & prefix & replacementText & suffix _
                & Mid$(resultText, matchItem.FirstIndex + matchItem.Length + 1)
        End If
    Next index

    ReplaceIdentifier = resultText
End Function

' AS直後の単独IDはエイリアスとして扱う
Private Function IsAliasAfterAs(ByVal sourceText As String, ByVal identifierStart As Long) As Boolean
    Dim index As Long
    Dim tokenEnd As Long
    Dim tokenStart As Long

    index = identifierStart - 1
    Do While index > 0
        If Not IsWhitespace(Mid$(sourceText, index, 1)) Then Exit Do
        index = index - 1
    Loop

    tokenEnd = index
    Do While index > 0
        If Not IsIdentifierCharacter(Mid$(sourceText, index, 1)) Then Exit Do
        index = index - 1
    Loop
    tokenStart = index + 1

    If tokenEnd - tokenStart + 1 = 2 Then
        IsAliasAfterAs = (UCase$(Mid$(sourceText, tokenStart, 2)) = "AS")
    End If
End Function

' SQL上の空白文字か判定
Private Function IsWhitespace(ByVal value As String) As Boolean
    IsWhitespace = (value = " " Or value = vbTab Or value = vbCr Or value = vbLf)
End Function

' ASCII識別子として使う文字か判定
Private Function IsIdentifierCharacter(ByVal value As String) As Boolean
    IsIdentifierCharacter = (value Like "[A-Za-z0-9_]")
End Function

' 正規表現の特殊文字をリテラル扱いへエスケープ
Private Function EscapeRegexLiteral(ByVal value As String) As String
    Dim index As Long
    Dim character As String
    Dim resultText As String

    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        Select Case character
            Case "\", ".", "^", "$", "|", "?", "*", "+", "(", ")", "[", "]", "{", "}"
                resultText = resultText & "\" & character
            Case Else
                resultText = resultText & character
        End Select
    Next index

    EscapeRegexLiteral = resultText
End Function

' 辞書キーを文字数の降順で取得
Private Function SortedKeysByLengthDesc(ByVal dictionary As Object) As Variant
    Dim keys As Variant
    Dim outerIndex As Long
    Dim innerIndex As Long
    Dim tempValue As Variant

    If dictionary.Count = 0 Then
        SortedKeysByLengthDesc = Array()
        Exit Function
    End If

    ' 長いキーを先に処理し、包含関係にあるIDの誤置換を防ぐ
    keys = dictionary.Keys
    For outerIndex = LBound(keys) To UBound(keys) - 1
        For innerIndex = outerIndex + 1 To UBound(keys)
            If Len(CStr(keys(innerIndex))) > Len(CStr(keys(outerIndex))) Then
                tempValue = keys(outerIndex)
                keys(outerIndex) = keys(innerIndex)
                keys(innerIndex) = tempValue
            End If
        Next innerIndex
    Next outerIndex

    SortedKeysByLengthDesc = keys
End Function

' 辞書キーを値の昇順で取得
Private Function SortedKeysByValueAsc(ByVal dictionary As Object) As Variant
    Dim keys As Variant
    Dim outerIndex As Long
    Dim innerIndex As Long
    Dim tempValue As Variant

    If dictionary.Count = 0 Then
        SortedKeysByValueAsc = Array()
        Exit Function
    End If

    keys = dictionary.Keys
    For outerIndex = LBound(keys) To UBound(keys) - 1
        For innerIndex = outerIndex + 1 To UBound(keys)
            If CLng(dictionary(CStr(keys(innerIndex)))) < CLng(dictionary(CStr(keys(outerIndex)))) Then
                tempValue = keys(outerIndex)
                keys(outerIndex) = keys(innerIndex)
                keys(innerIndex) = tempValue
            End If
        Next innerIndex
    Next outerIndex

    SortedKeysByValueAsc = keys
End Function

' 行内の変換後値を出現位置付きで保持
Private Sub AddReplacementValue(ByVal replacementValues As Object, ByVal replacementText As String, ByVal firstMatchIndex As Long)
    ' 同じ変換後値は1行内で重複表示しない
    If replacementValues.Exists(replacementText) Then
        If firstMatchIndex >= 0 And CLng(replacementValues(replacementText)) > firstMatchIndex Then
            replacementValues(replacementText) = firstMatchIndex
        End If
    Else
        replacementValues(replacementText) = firstMatchIndex
    End If
End Sub

' 変換キー用の値を前後空白なしの文字列へ正規化
Private Function NormalizeKey(ByVal value As Variant) As String
    NormalizeKey = Trim$(CStr(value))
End Function

' 表示名用の値を前後空白なしの文字列へ正規化
Private Function NormalizeName(ByVal value As Variant) As String
    NormalizeName = Trim$(CStr(value))
End Function

' 変換に使える和名か判定
Private Function IsUsableJapaneseName(ByVal value As String) As Boolean
    Dim normalized As String

    normalized = Trim$(value)
    If Len(normalized) = 0 Then Exit Function
    If normalized = "-" Then Exit Function
    If InStr(1, normalized, MissingNameText(), vbBinaryCompare) > 0 Then Exit Function

    IsUsableJapaneseName = True
End Function

' 変換定義行に対応するparser専用フィールドIDを生成
Private Function ParserFieldIdentifier(ByVal rowNumber As Long) As String
    ParserFieldIdentifier = "__SAF_FIELD_R" & Format$(rowNumber, "000000") & "__"
End Function

' 変換定義シートを取得
Private Function GetReferenceSheet() As Worksheet
    Set GetReferenceSheet = ResolveOrCreateSheet(ReferenceSheetName(), REF_LEGACY_SHEET, 1)
End Function

' SQL解析シートを取得
Private Function GetSqlSheet() As Worksheet
    Set GetSqlSheet = ResolveOrCreateSheet(SqlSheetName(), SQL_LEGACY_SHEET, 2)
End Function

' アウトプットシートを取得
Private Function GetOutputSheet(Optional ByVal applyLayout As Boolean = True) As Worksheet
    Dim ws As Worksheet

    Set ws = ResolveOrCreateOutputSheet()
    If applyLayout Then ApplyOutputSheetLayout ws
    Set GetOutputSheet = ws
End Function

' アウトプット②を取得
Private Function GetOutputTwoSheet(Optional ByVal applyLayout As Boolean = True) As Worksheet
    Dim ws As Worksheet

    Set ws = ResolveOrCreateNamedSheet(OutputSheetTwoName())
    If applyLayout Then ApplyOutputTwoSheetLayout ws
    Set GetOutputTwoSheet = ws
End Function

' テーブル一覧を取得
Private Function GetTableListSheet(Optional ByVal applyLayout As Boolean = True) As Worksheet
    Dim ws As Worksheet

    Set ws = ResolveOrCreateNamedSheet(TableListSheetName())
    If applyLayout Then ApplyTableListLayout ws
    Set GetTableListSheet = ws
End Function

' アウトプットシートの表示と書式を適用
Private Sub ApplyOutputSheetLayout(ByVal ws As Worksheet)
    Dim lastRow As Long

    lastRow = LastOutputRow(ws)
    ApplyOutputColumnFont ws, OUTPUT_LAST_COLUMN
    ApplyOutputSheetDimensions ws, lastRow, OUTPUT_LAST_COLUMN, True, True
    ApplyOutputSheetView ws
End Sub

' アウトプットシートの成果物をA列からCL列までコピー
Public Sub CopyOutput(Optional ByVal showMessage As Boolean = True)
    CopyOutputRange GetOutputSheet(), OUTPUT_LAST_COLUMN, showMessage, "CopyOutput"
End Sub

' アウトプット②の使用範囲をクリップボードへコピー
Public Sub CopyOutputTwo(Optional ByVal showMessage As Boolean = True)
    CopyOutputRange GetOutputTwoSheet(), OUTPUT_TWO_LAST_COLUMN, showMessage, "CopyOutputTwo"
End Sub

' 指定したアウトプットシートの使用範囲をクリップボードへコピー
Private Sub CopyOutputRange( _
    ByVal wsOutput As Worksheet, _
    ByVal lastColumn As Long, _
    ByVal showMessage As Boolean, _
    ByVal errorSource As String)

    Dim lastRow As Long
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo CopyFail

    lastRow = LastOutputRow(wsOutput, lastColumn)
    If Application.CountA(wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(lastRow, lastColumn))) = 0 Then
        If showMessage Then MsgBox NoOutputToCopyMessage(), vbInformation
        Exit Sub
    End If

    wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(lastRow, lastColumn)).Copy
    If showMessage Then MsgBox CopyDoneMessage(), vbInformation
    Exit Sub

CopyFail:
    errorNumber = Err.Number
    errorDescription = Err.Description
    If showMessage Then
        MsgBox CopyFailedMessage() & errorDescription, vbExclamation
    Else
        Err.Raise errorNumber, errorSource, errorDescription
    End If
End Sub

' アウトプット範囲へ既定フォントを設定
Private Sub ApplyOutputSheetFont( _
    ByVal ws As Worksheet, _
    Optional ByVal lastRow As Long = 0, _
    Optional ByVal lastColumn As Long = OUTPUT_LAST_COLUMN)

    lastRow = MaxLong(lastRow, 1)
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn)).Font
        .Name = OutputFontName()
        .Size = OutputFontSize()
    End With
End Sub

' アウトプットシートの列幅、行高、文字配置を設定
' Parser value columns need wrapping disabled after multiline text is written.
Private Sub ApplyOutputValueLayout( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long, _
    ByVal usedOutputColumns As Object)

    Dim usedColumn As Variant

    lastRow = MaxLong(lastRow, 1)
    For Each usedColumn In usedOutputColumns.Items
        With ws.Range( _
            ws.Cells(1, CLng(usedColumn)), _
            ws.Cells(lastRow, CLng(usedColumn)))
            .WrapText = False
            .ShrinkToFit = False
        End With
    Next usedColumn
End Sub

' Setup and Clear establish defaults for cells that do not yet contain output.
Private Sub ApplyOutputColumnFont( _
    ByVal ws As Worksheet, _
    ByVal lastColumn As Long)

    With ws.Range(ws.Columns(1), ws.Columns(lastColumn)).Font
        .Name = OutputFontName()
        .Size = OutputFontSize()
    End With
End Sub

Private Sub ApplyOutputSheetDimensions( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long, _
    Optional ByVal lastColumn As Long = OUTPUT_LAST_COLUMN, _
    Optional ByVal applyTextLayout As Boolean = True, _
    Optional ByVal resetEntireColumns As Boolean = False)
    Dim layoutLastRow As Long
    Dim outputColumns As Range

    layoutLastRow = MaxLong(lastRow, 1)
    Set outputColumns = ws.Range(ws.Columns(1), ws.Columns(lastColumn))
    outputColumns.ColumnWidth = OUTPUT_COLUMN_WIDTH
    If applyTextLayout Then
        If resetEntireColumns Then
            With outputColumns
                .WrapText = False
                .ShrinkToFit = False
            End With
        Else
            With ws.Range(ws.Cells(1, 1), ws.Cells(layoutLastRow, lastColumn))
                .WrapText = False
                .ShrinkToFit = False
            End With
        End If
    End If
    ws.Range(ws.Rows(1), ws.Rows(layoutLastRow)).RowHeight = OUTPUT_ROW_HEIGHT
End Sub

' アウトプットシートの表示設定を適用
Private Sub ApplyOutputSheetView(ByVal ws As Worksheet)
    Dim previousSheet As Object
    Dim previousScreenUpdating As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo RestoreApplicationState
    previousScreenUpdating = Application.ScreenUpdating
    Set previousSheet = ActiveSheet
    Application.ScreenUpdating = False

    ' 目盛り線はウィンドウ単位のため、一時的に対象シートを表示
    ws.Activate
    ActiveWindow.DisplayGridlines = False

RestoreApplicationState:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description

    On Error Resume Next
    previousSheet.Activate
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0

    If errorNumber <> 0 Then
        Err.Raise errorNumber, errorSource, errorDescription
    End If
End Sub

' 既存名または旧シート名からシートを解決し、なければ作成
Private Function ResolveOrCreateSheet(ByVal primaryName As String, ByVal fallbackName As String, ByVal desiredIndex As Long) As Worksheet
    Dim ws As Worksheet

    Set ws = TryGetWorksheet(primaryName)
    If ws Is Nothing Then
        Set ws = TryGetWorksheet(fallbackName)
    End If
    If ws Is Nothing Then
        If ThisWorkbook.Worksheets.Count >= desiredIndex Then
            Set ws = ThisWorkbook.Worksheets(desiredIndex)
        Else
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        End If
    End If
    If ws.Name <> primaryName Then
        ws.Name = primaryName
    End If

    Set ResolveOrCreateSheet = ws
End Function

' 名前が一致するシートを取得し、存在しない場合は末尾へ安全に追加
Private Function ResolveOrCreateNamedSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    Set ws = TryGetWorksheet(sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set ResolveOrCreateNamedSheet = ws
End Function

' アウトプット①だけは任意の既存シートを流用せず、旧名移行または新規作成する
Private Function ResolveOrCreateOutputSheet() As Worksheet
    Dim ws As Worksheet

    Set ws = TryGetWorksheet(OutputSheetName())
    If ws Is Nothing Then Set ws = TryGetWorksheet(LegacyOutputSheetName())
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    End If
    If ws.Name <> OutputSheetName() Then ws.Name = OutputSheetName()

    Set ResolveOrCreateOutputSheet = ws
End Function

' シートを指定位置へ移動
Private Sub MoveWorksheetToIndex(ByVal ws As Worksheet, ByVal desiredIndex As Long)
    If ws.Index = desiredIndex Then Exit Sub

    If desiredIndex <= 1 Then
        ws.Move Before:=ThisWorkbook.Worksheets(1)
    Else
        ws.Move After:=ThisWorkbook.Worksheets(desiredIndex - 1)
    End If
End Sub

' 指定名のシートを存在する場合だけ取得
Private Function TryGetWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set TryGetWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

' 文字列キー用の辞書を作成
Private Function CreateTextDictionary() As Object
    Set CreateTextDictionary = CreateObject("Scripting.Dictionary")
    CreateTextDictionary.CompareMode = vbBinaryCompare
End Function

' 変換定義シートの見出しと列幅を設定
Private Sub ApplyReferenceHeader(ByVal ws As Worksheet)
    ResetHeaderRange ws.Range("A1:D1")
    ws.Cells(1, COL_TABLE_ID).Value = TableIdHeader()
    ws.Cells(1, COL_TABLE_NAME).Value = TableNameHeader()
    ws.Cells(1, COL_FIELD_ID).Value = FieldIdHeader()
    ws.Cells(1, COL_FIELD_NAME).Value = FieldNameHeader()
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:D").AutoFit
End Sub

' アウトプット②の表示と基本書式を適用
Private Sub ApplyOutputTwoSheetLayout( _
    ByVal ws As Worksheet, _
    Optional ByVal resetEntireColumns As Boolean = True)
    Dim lastRow As Long

    lastRow = LastOutputRow(ws, OUTPUT_TWO_LAST_COLUMN)
    If resetEntireColumns Then
        ApplyOutputColumnFont ws, OUTPUT_TWO_LAST_COLUMN
    Else
        ApplyOutputSheetFont ws, lastRow, OUTPUT_TWO_LAST_COLUMN
    End If
    ApplyOutputSheetDimensions _
        ws, lastRow, OUTPUT_TWO_LAST_COLUMN, True, resetEntireColumns
    ApplyOutputTwoHeaderFontWeight ws
    ApplyOutputSheetView ws
End Sub

' アウトプット②の見出しは背景色と罫線だけで区別し、太字にしない
Private Sub ApplyOutputTwoHeaderFontWeight(ByVal ws As Worksheet)
    ws.Range(ws.Cells(3, 1), ws.Cells(3, 48)).Font.Bold = False
    ws.Range(ws.Cells(3, 53), ws.Cells(3, OUTPUT_TWO_LAST_COLUMN)).Font.Bold = False
End Sub

' 変換定義A～D列の和名未取得表示を値変更へ追従させる
Private Sub ApplyMissingNameConditionalFormatting(ByVal ws As Worksheet)
    Dim targetRange As Range
    Dim currentDataRange As Range
    Dim missingNameCondition As FormatCondition
    Dim missingName As String
    Dim formulaText As String
    Dim lastRow As Long

    missingName = "(" & MissingNameText() & ")"
    lastRow = MaxLong(LastUsedRow(ws), 2)
    Set currentDataRange = ws.Range( _
        ws.Cells(2, COL_TABLE_ID), _
        ws.Cells(lastRow, COL_FIELD_NAME))
    currentDataRange.Interior.Pattern = xlPatternNone

    Set targetRange = ws.Range( _
        ws.Cells(2, COL_TABLE_ID), _
        ws.Cells(ws.Rows.Count, COL_FIELD_NAME))
    targetRange.FormatConditions.Delete
    formulaText = "=" & targetRange.Cells(1, 1).Address(False, False) & _
        "=""" & missingName & """"
    Set missingNameCondition = targetRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:=formulaText)
    missingNameCondition.Interior.Pattern = xlSolid
    missingNameCondition.Interior.Color = MISSING_NAME_FILL_COLOR
    missingNameCondition.StopIfTrue = True
End Sub

' SQL解析シートの見出しと列幅を設定
Private Sub ApplySqlHeader(ByVal ws As Worksheet)
    ResetHeaderRange ws.Range("A1:Z1")
    ws.Cells(1, COL_SQL).Value = SqlHeader()
    ws.Cells(1, COL_RESULT).Value = ResultHeader()
    ws.Cells(1, COL_REPLACEMENT).Value = ReplacementHeader()
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).RowHeight = 30
    ws.Columns("A:B").ColumnWidth = 42
    ws.Columns("C:Z").ColumnWidth = 24
    SetReplacementColumnsWrapText ws, False, 26, MaxLong(LastUsedRowInColumn(ws, COL_SQL), 1)
End Sub

' ヘッダー範囲の結合と内容を初期化
Private Sub ResetHeaderRange(ByVal headerRange As Range)
    headerRange.UnMerge
    headerRange.ClearContents
End Sub

' 変換内容列以降の折り返し設定を変更
Private Sub SetReplacementColumnsWrapText( _
    ByVal ws As Worksheet, _
    ByVal wrapEnabled As Boolean, _
    ByVal lastColumn As Long, _
    ByVal lastRow As Long)

    lastColumn = MaxLong(lastColumn, COL_REPLACEMENT)
    lastRow = MaxLong(lastRow, 1)
    ws.Range(ws.Cells(1, COL_REPLACEMENT), ws.Cells(lastRow, lastColumn)).WrapText = wrapEnabled
End Sub

' 変換定義シートとSQL解析シートのヘッダーを既定値に復元
Private Sub RestoreHeaders(ByVal wsRef As Worksheet, ByVal wsSql As Worksheet)
    ApplyReferenceHeader wsRef
    ApplySqlHeader wsSql
End Sub

' SQL解析シートへ解析ボタンとクリアボタンを配置
Private Sub EnsureSqlActionButtons(ByVal ws As Worksheet)
    Dim buttonTop As Double
    Dim buttonLeft As Double

    ThisWorkbook.DisplayDrawingObjects = xlDisplayShapes
    buttonTop = ws.Rows(1).Top + 2
    buttonLeft = ws.Columns("E").Left

    EnsureSqlActionButton _
        ws, "btnAnalyzeQueries", AnalyzeButtonText(), "AnalyzeQueries", _
        buttonLeft, buttonTop
    EnsureSqlActionButton _
        ws, "btnClearData", ClearButtonText(), "ClearData", _
        buttonLeft + 82, buttonTop
End Sub

' SQL action buttons are repaired in place so a long-running analysis cannot leave them hidden or collapsed.
Private Sub EnsureSqlActionButton( _
    ByVal ws As Worksheet, _
    ByVal buttonName As String, _
    ByVal captionText As String, _
    ByVal macroName As String, _
    ByVal buttonLeft As Double, _
    ByVal buttonTop As Double)

    Dim actionButton As Object

    On Error Resume Next
    Set actionButton = ws.Buttons(buttonName)
    On Error GoTo 0

    If actionButton Is Nothing Then
        DeleteShapeIfExists ws, buttonName
        Set actionButton = ws.Buttons.Add(buttonLeft, buttonTop, 72, 24)
        actionButton.Name = buttonName
    End If

    With actionButton
        .Placement = xlFreeFloating
        .Visible = True
        .Caption = captionText
        .OnAction = macroName
        .Left = buttonLeft
        .Top = buttonTop
        .Width = 72
        .Height = 24
    End With
    On Error Resume Next
    ws.Shapes(buttonName).ZOrder msoBringToFront
    On Error GoTo 0
End Sub

' 指定名の図形があれば削除
Private Sub DeleteShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
End Sub

' 解析結果の出力範囲をクリア
Private Sub ClearAnalyzeOutput(ByVal wsSql As Worksheet, ByVal lastInputRow As Long)
    Dim clearLastRow As Long
    Dim clearLastColumn As Long

    ' 前回の出力だけを消し、入力SQLは残す
    clearLastRow = MaxLong(lastInputRow, LastUsedRow(wsSql))
    clearLastColumn = MaxLong(LastUsedColumn(wsSql), COL_REPLACEMENT)
    If clearLastRow >= 2 Then
        wsSql.Range(wsSql.Cells(2, COL_RESULT), wsSql.Cells(clearLastRow, clearLastColumn)).ClearContents
    End If
End Sub

' 変換後値をC列以降へ出現順に出力
Private Sub WriteReplacementValues(ByVal wsSql As Worksheet, ByVal rowNumber As Long, ByVal replacementValues As Object)
    Dim index As Long
    Dim keys As Variant

    If replacementValues.Count = 0 Then
        Exit Sub
    End If

    keys = SortedKeysByValueAsc(replacementValues)
    For index = LBound(keys) To UBound(keys)
        SetOutputCellText wsSql.Cells(rowNumber, COL_REPLACEMENT + index), CStr(keys(index))
    Next index
End Sub

' Write all replacement values in one Excel call; this avoids one COM call per populated cell.
Private Sub WriteReplacementValuesBatch( _
    ByVal wsSql As Worksheet, _
    ByVal lastRow As Long, _
    ByVal replacementValuesByRow As Object, _
    ByRef maxReplacementCount As Long)

    Dim replacementValues As Object
    Dim replacementData() As Variant
    Dim keys As Variant
    Dim rowNumber As Long
    Dim index As Long

    maxReplacementCount = 0
    For rowNumber = 2 To lastRow
        If replacementValuesByRow.Exists(CStr(rowNumber)) Then
            Set replacementValues = replacementValuesByRow(CStr(rowNumber))
            maxReplacementCount = MaxLong(maxReplacementCount, replacementValues.Count)
        End If
    Next rowNumber

    If maxReplacementCount = 0 Or lastRow < 2 Then Exit Sub

    ReDim replacementData(1 To lastRow - 1, 1 To maxReplacementCount)
    For rowNumber = 2 To lastRow
        If replacementValuesByRow.Exists(CStr(rowNumber)) Then
            Set replacementValues = replacementValuesByRow(CStr(rowNumber))
            If replacementValues.Count > 0 Then
                keys = SortedKeysByValueAsc(replacementValues)
                For index = LBound(keys) To UBound(keys)
                    replacementData(rowNumber - 1, index - LBound(keys) + 1) = _
                        OutputTextValue(CStr(keys(index)))
                Next index
            End If
        End If
    Next rowNumber

    With wsSql.Range( _
        wsSql.Cells(2, COL_REPLACEMENT), _
        wsSql.Cells(lastRow, COL_REPLACEMENT + maxReplacementCount - 1))
        .NumberFormat = "@"
        .Value = replacementData
    End With
End Sub

' parserへ渡す論理行番号とSQL解析シートの行番号を対応付け
Private Sub AddQueryLineRows( _
    ByVal queryLineRows As Collection, _
    ByVal rowNumber As Long, _
    ByVal queryText As String)

    Dim lineIndex As Long

    For lineIndex = 1 To TextLineCount(queryText)
        queryLineRows.Add rowNumber
    Next lineIndex
End Sub

' parserの論理行に対応するSQL解析シートの入力セルへ移動
Private Sub FocusFallbackQueryLine( _
    ByVal wsSql As Worksheet, _
    ByVal queryLineRows As Collection, _
    ByVal queryLine As Long)

    Dim sqlRow As Long

    If queryLine < 1 Or queryLine > queryLineRows.Count Then Exit Sub
    sqlRow = CLng(queryLineRows.Item(queryLine))
    wsSql.Activate
    Application.Goto wsSql.Cells(sqlRow, COL_SQL), True
End Sub

' アウトプットシートへコピーボタンを配置
Private Sub InstallOutputButton( _
    ByVal ws As Worksheet, _
    ByVal buttonName As String, _
    ByVal macroName As String, _
    ByVal lastColumn As Long)
    Dim copyButton As Object
    Dim buttonLeft As Double
    Dim buttonTop As Double

    DeleteShapeIfExists ws, buttonName
    buttonLeft = ws.Columns(lastColumn + 1).Left + 4
    buttonTop = ws.Rows(1).Top + 2

    Set copyButton = ws.Buttons.Add(buttonLeft, buttonTop, 72, 24)
    With copyButton
        .Name = buttonName
        .Caption = CopyButtonText()
        .OnAction = macroName
    End With
End Sub

' 外部parserの描画計画をアウトプットシートへ反映
Private Function TryWriteExternalOutputPlan( _
    ByVal wsOutput As Worksheet, _
    ByVal wsRef As Worksheet, _
    ByVal queryText As String, _
    ByRef qualifications As Collection, _
    ByRef inputTableIds As Collection, _
    ByRef outputTableIds As Collection, _
    ByRef fallbackReason As String, _
    ByRef fallbackStartLine As Long, _
    ByRef fallbackEndLine As Long, _
    ByVal expectedQueryLineCount As Long, _
    ByRef transformedQueryLines As Collection, _
    ByRef finalReplacementValues As Collection, _
    ByRef hasTransformedQueryData As Boolean, _
    ByRef tableNameReferences As Collection, _
    ByRef hasTableNameReferenceData As Boolean) As Boolean

    Dim parserPath As String
    Dim inputPath As String
    Dim mappingPath As String
    Dim outputPath As String
    Dim exitCode As Long
    Dim outputText As String
    Dim succeeded As Boolean

    On Error GoTo ParserError
    fallbackReason = ""
    fallbackStartLine = 0
    fallbackEndLine = 0
    Set qualifications = New Collection
    Set transformedQueryLines = New Collection
    Set finalReplacementValues = New Collection
    Set tableNameReferences = New Collection
    Set inputTableIds = New Collection
    Set outputTableIds = New Collection
    hasTransformedQueryData = False
    hasTableNameReferenceData = False

    parserPath = ResolveParserExePath()
    If Len(parserPath) = 0 Then
        fallbackReason = ParserNotFoundReason()
        GoTo CleanUp
    End If

    inputPath = TemporaryFilePath("saf_sql_", ".sql")
    mappingPath = TemporaryFilePath("saf_mapping_", ".txt")
    outputPath = TemporaryFilePath("saf_plan_", ".txt")
    WriteUtf8TextFile inputPath, queryText
    WriteMappingDefinitionFile mappingPath, wsRef

    exitCode = RunParserPlanProcess(parserPath, inputPath, mappingPath, outputPath)
    If exitCode <> 0 Then
        fallbackReason = ParserExitCodeReason(exitCode)
        GoTo CleanUp
    End If
    If Not FileExists(outputPath) Then
        fallbackReason = ParserOutputMissingReason()
        GoTo CleanUp
    End If

    outputText = ReadUnicodeTextFile(outputPath)
    succeeded = ApplyOutputPlan( _
        wsOutput, outputText, qualifications, inputTableIds, outputTableIds, _
        fallbackReason, fallbackStartLine, fallbackEndLine, expectedQueryLineCount, _
        transformedQueryLines, finalReplacementValues, hasTransformedQueryData, _
        tableNameReferences, hasTableNameReferenceData)
    If Not succeeded Then
        fallbackReason = ParserOutputInvalidReason()
    End If

CleanUp:
    On Error Resume Next
    DeleteFileIfExists inputPath
    DeleteFileIfExists mappingPath
    DeleteFileIfExists outputPath
    On Error GoTo 0
    TryWriteExternalOutputPlan = succeeded
    Exit Function

ParserError:
    fallbackReason = ParserIntegrationErrorReason(Err.Description)
    Resume CleanUp
End Function

' 変換定義シートをparser連携用の行形式で保存
Private Sub WriteMappingDefinitionFile(ByVal filePath As String, ByVal wsRef As Worksheet)
    Dim rowNumber As Long
    Dim lastRow As Long
    Dim mappingText As String
    Dim fieldName As String
    Dim parserFieldId As String
    Dim mappingValues As Variant
    Dim mappingLines() As String

    lastRow = LastUsedRow(wsRef)
    ReDim mappingLines(0 To MaxLong(lastRow - 1, 0))
    mappingLines(0) = "SAF_MAPPINGS" & vbTab & "2"

    If lastRow >= 2 Then
        mappingValues = wsRef.Range( _
            wsRef.Cells(2, COL_TABLE_ID), _
            wsRef.Cells(lastRow, COL_FIELD_NAME)).Value2
    End If
    For rowNumber = 2 To lastRow
        fieldName = NormalizeName(mappingValues(rowNumber - 1, COL_FIELD_NAME))
        parserFieldId = ""
        If IsUsableJapaneseName(fieldName) Then
            parserFieldId = ParserFieldIdentifier(rowNumber)
        End If
        mappingLines(rowNumber - 1) = "M" & vbTab & _
            EscapeProtocolField(CStr(mappingValues(rowNumber - 1, COL_TABLE_ID))) & vbTab & _
            EscapeProtocolField(CStr(mappingValues(rowNumber - 1, COL_TABLE_NAME))) & vbTab & _
            EscapeProtocolField(CStr(mappingValues(rowNumber - 1, COL_FIELD_ID))) & vbTab & _
            EscapeProtocolField(fieldName) & vbTab & _
            EscapeProtocolField(parserFieldId)
    Next rowNumber

    mappingText = Join(mappingLines, vbCrLf)
    WriteUtf8TextFile filePath, mappingText
End Sub

' parserを描画計画形式で同期実行
Private Function RunParserPlanProcess( _
    ByVal parserPath As String, _
    ByVal inputPath As String, _
    ByVal mappingPath As String, _
    ByVal outputPath As String) As Long

    Dim shell As Object
    Dim commandText As String

    commandText = QuoteCommandArgument(parserPath) & _
        " --input " & QuoteCommandArgument(inputPath) & _
        " --mappings " & QuoteCommandArgument(mappingPath) & _
        " --output " & QuoteCommandArgument(outputPath) & _
        " --format vba-plan"

    Set shell = CreateObject("WScript.Shell")
    RunParserPlanProcess = CLng(shell.Run(commandText, 0, True))
End Function

' parserの描画計画を検証してセルと書式へ反映
Private Function ApplyOutputPlan( _
    ByVal ws As Worksheet, _
    ByVal planText As String, _
    ByRef qualifications As Collection, _
    ByRef inputTableIds As Collection, _
    ByRef outputTableIds As Collection, _
    ByRef fallbackReason As String, _
    ByRef fallbackStartLine As Long, _
    ByRef fallbackEndLine As Long, _
    ByVal expectedQueryLineCount As Long, _
    ByRef transformedQueryLines As Collection, _
    ByRef finalReplacementValues As Collection, _
    ByRef hasTransformedQueryData As Boolean, _
    ByRef tableNameReferences As Collection, _
    ByRef hasTableNameReferenceData As Boolean) As Boolean
    Dim lines As Variant
    Dim fields As Variant
    Dim cellValues As Variant
    Dim section As Variant
    Dim sections As Collection
    Dim parsedQualifications As Collection
    Dim parsedTransformedQueryLines As Collection
    Dim parsedFinalReplacementValues As Collection
    Dim parsedTableNameReferences As Collection
    Dim parsedInputTableIds As Collection
    Dim parsedOutputTableIds As Collection
    Dim usedOutputColumns As Object
    Dim transformedQueryLineSeen As Object
    Dim finalReplacementValueSeen As Object
    Dim tableNameReferenceSeen As Object
    Dim usedColumn As Variant
    Dim lineText As String
    Dim normalizedText As String
    Dim cellValue As String
    Dim cellText As String
    Dim originalValue As String
    Dim qualifiedValue As String
    Dim transformedValue As String
    Dim finalReplacementValue As String
    Dim sourceValue As String
    Dim physicalTableId As String
    Dim replacementSuffix As String
    Dim tableId As String
    Dim recordKey As String
    Dim parsedFallbackReason As String
    Dim lineIndex As Long
    Dim planVersion As Long
    Dim fallbackFlag As Long
    Dim rowCount As Long
    Dim rowNumber As Long
    Dim columnNumber As Long
    Dim queryLine As Long
    Dim qualificationOrder As Long
    Dim replacementOrder As Long
    Dim previousReplacementQueryLine As Long
    Dim previousReplacementOrder As Long
    Dim replacementValueSeen As Boolean
    Dim startRow As Long
    Dim endRow As Long
    Dim parsedFallbackStartLine As Long
    Dim parsedFallbackEndLine As Long
    Dim fallbackDiagnosticSeen As Boolean
    Dim tableNameReference As Variant
    Dim outputRange As Range

    On Error GoTo InvalidPlan

    normalizedText = Replace(planText, vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)
    lines = Split(normalizedText, vbLf)
    fields = Split(CStr(lines(0)), vbTab)
    If UBound(fields) <> 3 Then Exit Function
    If CStr(fields(0)) <> "SAF_OUTPUT_PLAN" Then Exit Function
    If Not IsNumeric(fields(1)) Then Exit Function
    planVersion = CLng(fields(1))
    If planVersion < 1 Or planVersion > 6 Then Exit Function
    If Not IsNumeric(fields(2)) Then Exit Function
    rowCount = CLng(fields(2))
    If rowCount < 0 Then Exit Function
    If Not IsNumeric(fields(3)) Then Exit Function
    fallbackFlag = CLng(fields(3))
    If fallbackFlag <> 0 And fallbackFlag <> 1 Then Exit Function
    If expectedQueryLineCount < 0 Then Exit Function

    Set sections = New Collection
    Set parsedQualifications = New Collection
    Set parsedTransformedQueryLines = New Collection
    Set parsedFinalReplacementValues = New Collection
    Set parsedTableNameReferences = New Collection
    Set parsedInputTableIds = New Collection
    Set parsedOutputTableIds = New Collection
    Set usedOutputColumns = CreateTextDictionary()
    Set transformedQueryLineSeen = CreateTextDictionary()
    Set finalReplacementValueSeen = CreateTextDictionary()
    Set tableNameReferenceSeen = CreateCaseInsensitiveTextDictionary()
    If rowCount > 0 Then
        ReDim cellValues(1 To rowCount, 1 To OUTPUT_LAST_COLUMN)
    End If

    ' 全行を先に検証し、セル値と書式セクションをメモリ上へ構成する
    For lineIndex = 1 To UBound(lines)
        lineText = CStr(lines(lineIndex))
        If Len(lineText) > 0 Then
            fields = Split(lineText, vbTab)

            Select Case CStr(fields(0))
                Case "C"
                    If UBound(fields) <> 3 Then GoTo InvalidPlan
                    If Not IsNumeric(fields(1)) Or Not IsNumeric(fields(2)) Then GoTo InvalidPlan
                    rowNumber = CLng(fields(1))
                    columnNumber = CLng(fields(2))
                    If rowNumber < 1 Or rowNumber > rowCount Or _
                        columnNumber < 1 Or columnNumber > OUTPUT_LAST_COLUMN Then GoTo InvalidPlan
                    cellValue = UnescapeProtocolField(CStr(fields(3)))
                    If Left$(cellValue, 1) = "'" Then
                        cellValue = "'" & cellValue
                    End If
                    cellValues(rowNumber, columnNumber) = cellValue
                    usedOutputColumns(CStr(columnNumber)) = columnNumber
                Case "S"
                    If UBound(fields) <> 3 Then GoTo InvalidPlan
                    If Not IsNumeric(fields(2)) Or Not IsNumeric(fields(3)) Then GoTo InvalidPlan
                    startRow = CLng(fields(2))
                    endRow = CLng(fields(3))
                    If startRow < 1 Or endRow < startRow Or endRow > rowCount Then GoTo InvalidPlan
                    sections.Add Array(CStr(fields(1)), startRow, endRow)
                Case "Q"
                    If planVersion < 2 Or UBound(fields) <> 4 Then GoTo InvalidPlan
                    If Not IsNumeric(fields(1)) Or Not IsNumeric(fields(2)) Then GoTo InvalidPlan
                    queryLine = CLng(fields(1))
                    qualificationOrder = CLng(fields(2))
                    If queryLine < 1 Or qualificationOrder < 1 Then GoTo InvalidPlan
                    originalValue = UnescapeProtocolField(CStr(fields(3)))
                    qualifiedValue = UnescapeProtocolField(CStr(fields(4)))
                    If Len(originalValue) = 0 Or Len(qualifiedValue) = 0 Then GoTo InvalidPlan
                    parsedQualifications.Add Array( _
                        queryLine, qualificationOrder, originalValue, qualifiedValue)
                Case "R"
                    If planVersion < 5 Or UBound(fields) <> 2 Then GoTo InvalidPlan
                    If Not IsNumeric(fields(1)) Then GoTo InvalidPlan
                    queryLine = CLng(fields(1))
                    If queryLine < 1 Or queryLine > expectedQueryLineCount Then GoTo InvalidPlan
                    recordKey = CStr(queryLine)
                    If transformedQueryLineSeen.Exists(recordKey) Then GoTo InvalidPlan
                    transformedValue = UnescapeProtocolField(CStr(fields(2)))
                    transformedQueryLineSeen.Add recordKey, True
                    parsedTransformedQueryLines.Add Array(queryLine, transformedValue)
                Case "V"
                    If planVersion < 5 Or UBound(fields) <> 3 Then GoTo InvalidPlan
                    If Not IsNumeric(fields(1)) Or Not IsNumeric(fields(2)) Then GoTo InvalidPlan
                    queryLine = CLng(fields(1))
                    replacementOrder = CLng(fields(2))
                    If queryLine < 1 Or queryLine > expectedQueryLineCount Or _
                        replacementOrder < 1 Then GoTo InvalidPlan
                    recordKey = CStr(queryLine) & vbTab & CStr(replacementOrder)
                    If finalReplacementValueSeen.Exists(recordKey) Then GoTo InvalidPlan
                    If replacementValueSeen Then
                        If queryLine < previousReplacementQueryLine Or _
                            (queryLine = previousReplacementQueryLine And _
                                replacementOrder <= previousReplacementOrder) Then GoTo InvalidPlan
                    End If
                    finalReplacementValue = UnescapeProtocolField(CStr(fields(3)))
                    If Len(finalReplacementValue) = 0 Then GoTo InvalidPlan
                    finalReplacementValueSeen.Add recordKey, True
                    parsedFinalReplacementValues.Add Array( _
                        queryLine, replacementOrder, finalReplacementValue)
                    previousReplacementQueryLine = queryLine
                    previousReplacementOrder = replacementOrder
                    replacementValueSeen = True
                Case "N"
                    If planVersion < 6 Or UBound(fields) <> 5 Then GoTo InvalidPlan
                    If Not IsNumeric(fields(1)) Or Not IsNumeric(fields(2)) Then GoTo InvalidPlan
                    rowNumber = CLng(fields(1))
                    columnNumber = CLng(fields(2))
                    If rowNumber < 1 Or rowNumber > rowCount Or _
                        columnNumber < 1 Or columnNumber > OUTPUT_LAST_COLUMN Then GoTo InvalidPlan
                    sourceValue = UnescapeProtocolField(CStr(fields(3)))
                    physicalTableId = UnescapeProtocolField(CStr(fields(4)))
                    replacementSuffix = UnescapeProtocolField(CStr(fields(5)))
                    If Len(sourceValue) = 0 Or Len(physicalTableId) = 0 Then GoTo InvalidPlan
                    If StrComp( _
                        Left$(sourceValue, Len("(" & MissingNameText() & ")[")), _
                        "(" & MissingNameText() & ")[", _
                        vbBinaryCompare) <> 0 Then GoTo InvalidPlan
                    If Len(replacementSuffix) > 0 Then
                        If Left$(replacementSuffix, 1) <> "[" Or _
                            Right$(replacementSuffix, 1) <> "]" Then GoTo InvalidPlan
                    End If
                    recordKey = CStr(rowNumber) & vbTab & CStr(columnNumber) & _
                        vbTab & sourceValue
                    If tableNameReferenceSeen.Exists(recordKey) Then GoTo InvalidPlan
                    tableNameReferenceSeen.Add recordKey, True
                    parsedTableNameReferences.Add Array( _
                        rowNumber, columnNumber, sourceValue, physicalTableId, replacementSuffix)
                Case "F"
                    If planVersion < 4 Or UBound(fields) <> 3 Then GoTo InvalidPlan
                    If fallbackFlag <> 1 Or fallbackDiagnosticSeen Then GoTo InvalidPlan
                    If Not IsNumeric(fields(1)) Or Not IsNumeric(fields(2)) Then GoTo InvalidPlan
                    parsedFallbackStartLine = CLng(fields(1))
                    parsedFallbackEndLine = CLng(fields(2))
                    If parsedFallbackStartLine < 0 Or _
                        parsedFallbackEndLine < parsedFallbackStartLine Then GoTo InvalidPlan
                    parsedFallbackReason = UnescapeProtocolField(CStr(fields(3)))
                    If Len(parsedFallbackReason) = 0 Then GoTo InvalidPlan
                    fallbackDiagnosticSeen = True
                Case "T"
                    If planVersion < 3 Or UBound(fields) <> 2 Then GoTo InvalidPlan
                    tableId = UnescapeProtocolField(CStr(fields(2)))
                    If Len(tableId) = 0 Then GoTo InvalidPlan
                    Select Case UCase$(CStr(fields(1)))
                        Case "INPUT"
                            parsedInputTableIds.Add tableId
                        Case "OUTPUT"
                            parsedOutputTableIds.Add tableId
                        Case Else
                            GoTo InvalidPlan
                    End Select
                Case Else
                    GoTo InvalidPlan
            End Select
        End If
    Next lineIndex
    If planVersion >= 4 And fallbackFlag = 1 And Not fallbackDiagnosticSeen Then GoTo InvalidPlan
    If planVersion >= 5 Then
        If parsedTransformedQueryLines.Count <> expectedQueryLineCount Then GoTo InvalidPlan
        For queryLine = 1 To expectedQueryLineCount
            If Not transformedQueryLineSeen.Exists(CStr(queryLine)) Then GoTo InvalidPlan
        Next queryLine
    End If
    If planVersion >= 6 Then
        For Each tableNameReference In parsedTableNameReferences
            rowNumber = CLng(tableNameReference(0))
            columnNumber = CLng(tableNameReference(1))
            sourceValue = CStr(tableNameReference(2))
            cellText = CStr(cellValues(rowNumber, columnNumber))
            If Not IsOutputTableNameReference( _
                cellText, columnNumber, sourceValue) Then GoTo InvalidPlan
        Next tableNameReference
    End If

    If rowCount > 0 Then
        Set outputRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, OUTPUT_LAST_COLUMN))
        For Each usedColumn In usedOutputColumns.Items
            ws.Range( _
                ws.Cells(1, CLng(usedColumn)), _
                ws.Cells(rowCount, CLng(usedColumn))).NumberFormat = "@"
        Next usedColumn
        ' Excel COM呼出しをセルごとではなく1回へまとめる
        outputRange.Value = cellValues
    End If
    ' 長文・改行を含む値の書込み後に折り返し、縮小表示、行高を確定する
    ApplyOutputSheetDimensions ws, rowCount, OUTPUT_LAST_COLUMN, False
    For Each section In sections
        ApplyOutputSectionStyle ws, CStr(section(0)), CLng(section(1)), CLng(section(2))
    Next section
    ApplyOutputValueLayout ws, rowCount, usedOutputColumns
    ApplyOutputSheetFont ws, rowCount
    ApplyOutputSheetView ws
    Set qualifications = parsedQualifications
    Set transformedQueryLines = parsedTransformedQueryLines
    Set finalReplacementValues = parsedFinalReplacementValues
    Set tableNameReferences = parsedTableNameReferences
    Set inputTableIds = parsedInputTableIds
    Set outputTableIds = parsedOutputTableIds
    hasTransformedQueryData = (planVersion >= 5)
    hasTableNameReferenceData = (planVersion >= 6)
    fallbackReason = parsedFallbackReason
    fallbackStartLine = parsedFallbackStartLine
    fallbackEndLine = parsedFallbackEndLine
    ApplyOutputPlan = True
    Exit Function

InvalidPlan:
    ApplyOutputPlan = False
End Function

' アウトプット②のタイトルと見出しがなければ初期化
Private Sub EnsureOutputTwoStructure(ByVal ws As Worksheet)
    If CStr(ws.Cells(1, 1).Value) <> InputInformationTitle() Or _
        CStr(ws.Cells(1, 53).Value) <> OutputInformationTitle() Then
        ClearOutputTwoSheet ws
    Else
        ApplyOutputTwoSheetLayout ws
    End If
End Sub

' アウトプット②をタイトルと見出しだけの状態へ戻す
Private Sub ClearOutputTwoSheet(ByVal ws As Worksheet)
    ResetOutputTwoSurface ws
    ApplyOutputTwoBlockStyle ws, 1, 3
    ApplyOutputTwoBlockStyle ws, 53, 3
    ApplyOutputTwoSheetLayout ws
End Sub

' 入出力テーブルをテーブル一覧と照合してアウトプット②へ描画
Private Sub RenderOutputTwo( _
    ByVal ws As Worksheet, _
    ByVal inputTableIds As Collection, _
    ByVal outputTableIds As Collection, _
    ByVal tableMaster As Object)

    Dim inputLastRow As Long
    Dim outputLastRow As Long

    ResetOutputTwoSurface ws
    inputLastRow = WriteOutputTwoBlock(ws, 1, inputTableIds, tableMaster)
    outputLastRow = WriteOutputTwoBlock(ws, 53, outputTableIds, tableMaster)
    ApplyOutputTwoBlockStyle ws, 1, inputLastRow
    ApplyOutputTwoBlockStyle ws, 53, outputLastRow
    ApplyOutputTwoSheetLayout ws, False
End Sub

' アウトプット②へ表示する完全一致の物理IDで未取得のテーブル表示を補完
Private Sub ApplyOutputTwoNamesToMissingTableDisplays( _
    ByVal wsOutput As Worksheet, _
    ByVal inputTableIds As Collection, _
    ByVal outputTableIds As Collection, _
    ByVal tableMaster As Object, _
    ByVal tableNameReferences As Collection, _
    ByVal hasTableNameReferenceData As Boolean)

    Dim matchedNames As Object
    Dim referenceValues As Variant
    Dim joinHeadingValues As Variant
    Dim rowNumber As Long
    Dim lastRow As Long
    Dim referencePrefix As String
    Dim cellText As String
    Dim updatedText As String

    Set matchedNames = CreateCaseInsensitiveTextDictionary()
    AddMatchedOutputTwoNames matchedNames, inputTableIds, tableMaster
    AddMatchedOutputTwoNames matchedNames, outputTableIds, tableMaster
    If matchedNames.Count = 0 Then Exit Sub

    If hasTableNameReferenceData Then
        ApplyStructuredOutputTableNames wsOutput, matchedNames, tableNameReferences
        Exit Sub
    End If

    lastRow = LastOutputRow(wsOutput, OUTPUT_LAST_COLUMN)
    If lastRow < 1 Then Exit Sub
    If lastRow = 1 Then
        ReDim referenceValues(1 To 1, 1 To 1)
        ReDim joinHeadingValues(1 To 1, 1 To 1)
        referenceValues(1, 1) = wsOutput.Cells(1, 1).Value2
        joinHeadingValues(1, 1) = wsOutput.Cells(1, 17).Value2
    Else
        referenceValues = wsOutput.Range( _
            wsOutput.Cells(1, 1), _
            wsOutput.Cells(lastRow, 1)).Value2
        joinHeadingValues = wsOutput.Range( _
            wsOutput.Cells(1, 17), _
            wsOutput.Cells(lastRow, 17)).Value2
    End If

    referencePrefix = ReferenceTablesText() & ": "
    For rowNumber = 1 To lastRow
        cellText = CStr(referenceValues(rowNumber, 1))
        If StrComp(Left$(cellText, Len(referencePrefix)), referencePrefix, vbBinaryCompare) = 0 Then
            updatedText = ReplaceMissingOutputTableNames( _
                cellText, matchedNames, False)
            If StrComp(updatedText, cellText, vbBinaryCompare) <> 0 Then
                SetOutputCellText wsOutput.Cells(rowNumber, 1), updatedText
            End If
        End If

        cellText = CStr(joinHeadingValues(rowNumber, 1))
        If IsOutputJoinHeading(cellText) Then
            ' JOIN条件セルと区別するため、物理IDと表示別名を持つ採番対象だけを補完する。
            updatedText = ReplaceMissingOutputTableNames( _
                cellText, matchedNames, True)
            If StrComp(updatedText, cellText, vbBinaryCompare) <> 0 Then
                SetOutputCellText wsOutput.Cells(rowNumber, 17), updatedText
            End If
        End If
    Next rowNumber
End Sub

' parserが特定したセル・物理ID・表示接尾辞だけを使って名称を補完
Private Sub ApplyStructuredOutputTableNames( _
    ByVal wsOutput As Worksheet, _
    ByVal matchedNames As Object, _
    ByVal tableNameReferences As Collection)

    Dim tableNameReference As Variant
    Dim rowNumber As Long
    Dim columnNumber As Long
    Dim sourceValue As String
    Dim physicalTableId As String
    Dim replacementText As String
    Dim cellText As String
    Dim updatedText As String

    For Each tableNameReference In tableNameReferences
        rowNumber = CLng(tableNameReference(0))
        columnNumber = CLng(tableNameReference(1))
        sourceValue = CStr(tableNameReference(2))
        physicalTableId = NormalizePhysicalTableId(CStr(tableNameReference(3)))
        If matchedNames.Exists(physicalTableId) Then
            replacementText = CStr(matchedNames(physicalTableId)) & _
                CStr(tableNameReference(4))
            cellText = CStr(wsOutput.Cells(rowNumber, columnNumber).Value2)
            updatedText = Replace( _
                cellText, sourceValue, replacementText, 1, -1, vbTextCompare)
            If StrComp(updatedText, cellText, vbBinaryCompare) <> 0 Then
                SetOutputCellText wsOutput.Cells(rowNumber, columnNumber), updatedText
            End If
        End If
    Next tableNameReference
End Sub

' 物理ID付きの未取得表示だけを対応するテーブル名称へ置換
Private Function ReplaceMissingOutputTableNames( _
    ByVal sourceText As String, _
    ByVal matchedNames As Object, _
    ByVal requireDisplayAlias As Boolean) As String

    Dim tableId As Variant
    Dim missingToken As String
    Dim replacementText As String
    Dim resultText As String

    resultText = sourceText
    For Each tableId In matchedNames.Keys
        missingToken = "(" & MissingNameText() & ")[" & CStr(tableId) & "]"
        replacementText = CStr(matchedNames(CStr(tableId)))
        If requireDisplayAlias Then
            missingToken = missingToken & "["
            replacementText = replacementText & "["
        End If
        resultText = Replace( _
            resultText, _
            missingToken, _
            replacementText, _
            1, _
            -1, _
            vbTextCompare)
    Next tableId
    ReplaceMissingOutputTableNames = resultText
End Function

' parserが列Qへ生成するJOIN見出しだけを条件式や文字列から識別
Private Function IsOutputJoinHeading(ByVal cellText As String) As Boolean
    If Len(cellText) < 2 Then Exit Function
    If Left$(cellText, 1) <> W(&HFF1C) Or Right$(cellText, 1) <> W(&HFF1E) Then Exit Function
    IsOutputJoinHeading = InStr(1, cellText, " JOIN ", vbBinaryCompare) > 0
End Function

' Nレコードの元表示が参照一覧またはJOIN左右の1要素と完全一致することを確認
Private Function IsOutputTableNameReference( _
    ByVal cellText As String, _
    ByVal columnNumber As Long, _
    ByVal sourceValue As String) As Boolean

    Dim referencePrefix As String
    Dim joinText As String
    Dim joinDelimiter As Variant
    Dim delimiterPosition As Long

    If columnNumber = 1 Then
        referencePrefix = ReferenceTablesText() & ": "
        If StrComp( _
            Left$(cellText, Len(referencePrefix)), _
            referencePrefix, _
            vbBinaryCompare) <> 0 Then Exit Function
        IsOutputTableNameReference = ContainsExactTableDisplay( _
            Mid$(cellText, Len(referencePrefix) + 1), sourceValue)
        Exit Function
    End If

    If columnNumber <> 17 Or Not IsOutputJoinHeading(cellText) Then Exit Function
    joinText = Mid$(cellText, 2, Len(cellText) - 2)
    For Each joinDelimiter In Array( _
        " INNER JOIN ", _
        " LEFT JOIN ", _
        " RIGHT JOIN ", _
        " FULL JOIN ", _
        " JOIN ")
        delimiterPosition = InStr( _
            1, joinText, CStr(joinDelimiter), vbBinaryCompare)
        If delimiterPosition > 0 Then
            IsOutputTableNameReference = _
                ContainsExactTableDisplay( _
                    Left$(joinText, delimiterPosition - 1), sourceValue) Or _
                ContainsExactTableDisplay( _
                    Mid$(joinText, delimiterPosition + Len(CStr(joinDelimiter))), _
                    sourceValue)
            Exit Function
        End If
    Next joinDelimiter
End Function

' 読点区切りのテーブル表示に完全一致する要素があるかを確認
Private Function ContainsExactTableDisplay( _
    ByVal displayList As String, _
    ByVal sourceValue As String) As Boolean

    Dim displayItem As Variant

    For Each displayItem In Split(displayList, W(&H3001))
        If StrComp(CStr(displayItem), sourceValue, vbTextCompare) = 0 Then
            ContainsExactTableDisplay = True
            Exit Function
        End If
    Next displayItem
End Function

' アウトプット②の明細対象かつ有効な名称を持つテーブルだけを辞書へ追加
Private Sub AddMatchedOutputTwoNames( _
    ByVal matchedNames As Object, _
    ByVal tableIds As Collection, _
    ByVal tableMaster As Object)

    Dim tableId As Variant
    Dim normalizedId As String
    Dim tableRow As Variant
    Dim tableName As String

    For Each tableId In tableIds
        normalizedId = NormalizePhysicalTableId(CStr(tableId))
        If tableMaster.Exists(normalizedId) Then
            tableRow = tableMaster(normalizedId)
            tableName = CStr(tableRow(1))
            If IsUsableJapaneseName(tableName) And Not matchedNames.Exists(normalizedId) Then
                matchedNames.Add normalizedId, tableName
            End If
        End If
    Next tableId
End Sub

' アウトプット②の旧内容・結合・書式を消して固定見出しを配置
Private Sub ResetOutputTwoSurface(ByVal ws As Worksheet)
    Dim clearLastRow As Long

    clearLastRow = MaxLong(LastOutputRow(ws, OUTPUT_TWO_LAST_COLUMN), 3)
    With ws.Range(ws.Cells(1, 1), ws.Cells(clearLastRow, OUTPUT_TWO_LAST_COLUMN))
        .UnMerge
        .ClearContents
        .ClearFormats
        .WrapText = False
        .ShrinkToFit = False
    End With

    InitializeOutputTwoBlock ws, 1, InputInformationTitle()
    InitializeOutputTwoBlock ws, 53, OutputInformationTitle()
    With ws.Range(ws.Cells(1, 49), ws.Cells(clearLastRow, 52))
        .ClearContents
        .ClearFormats
    End With
End Sub

' アウトプット②の片側ブロックへタイトルと見出しを設定
Private Sub InitializeOutputTwoBlock( _
    ByVal ws As Worksheet, _
    ByVal startColumn As Long, _
    ByVal titleText As String)

    SetOutputCellText ws.Cells(1, startColumn), titleText
    ws.Range(ws.Cells(3, startColumn), ws.Cells(3, startColumn + 1)).Merge
    SetOutputCellText ws.Cells(3, startColumn), "No"
    SetOutputCellText ws.Cells(3, startColumn + 2), TableListIdHeader()
    SetOutputCellText ws.Cells(3, startColumn + 18), TableListNameHeader()
    SetOutputCellText ws.Cells(3, startColumn + 43), TableNumberHeader()
End Sub

' 物理テーブルIDをマスターと照合し、片側ブロックへ出力
Private Function WriteOutputTwoBlock( _
    ByVal ws As Worksheet, _
    ByVal startColumn As Long, _
    ByVal tableIds As Collection, _
    ByVal tableMaster As Object) As Long

    Dim tableId As Variant
    Dim tableRow As Variant
    Dim rowNumber As Long
    Dim itemNumber As Long
    Dim emitted As Object
    Dim outputRows As Collection

    Set emitted = CreateCaseInsensitiveTextDictionary()
    Set outputRows = New Collection
    rowNumber = 3
    For Each tableId In tableIds
        tableId = NormalizePhysicalTableId(CStr(tableId))
        If tableMaster.Exists(CStr(tableId)) Then
            tableRow = tableMaster(CStr(tableId))
            If Not emitted.Exists(CStr(tableId)) Then
                emitted.Add CStr(tableId), True
                InsertOutputTwoRowByNumber outputRows, Array( _
                    CStr(tableId), CStr(tableRow(1)), CStr(tableRow(2)))
            End If
        End If
    Next tableId

    For Each tableRow In outputRows
        itemNumber = itemNumber + 1
        rowNumber = rowNumber + 1
        ws.Range( _
            ws.Cells(rowNumber, startColumn), _
            ws.Cells(rowNumber, startColumn + 1)).Merge
        SetOutputCellText ws.Cells(rowNumber, startColumn), CStr(itemNumber)
        SetOutputCellText ws.Cells(rowNumber, startColumn + 2), CStr(tableRow(0))
        SetOutputCellText ws.Cells(rowNumber, startColumn + 18), CStr(tableRow(1))
        SetOutputCellText ws.Cells(rowNumber, startColumn + 43), CStr(tableRow(2))
    Next tableRow

    WriteOutputTwoBlock = rowNumber
End Function

' 照合済みの明細を複合番号の昇順へ安定挿入する
Private Sub InsertOutputTwoRowByNumber( _
    ByVal outputRows As Collection, _
    ByVal newRow As Variant)

    Dim index As Long
    Dim existingRow As Variant

    For index = 1 To outputRows.Count
        existingRow = outputRows(index)
        If CompareCompositeTableNumbers( _
            CStr(newRow(2)), CStr(existingRow(2))) < 0 Then

            outputRows.Add Item:=newRow, Before:=index
            Exit Sub
        End If
    Next index
    outputRows.Add newRow
End Sub

' 「数値-数値」を前半、後半の順で比較する。形式外の値は有効値の後ろへ安定配置する
Private Function CompareCompositeTableNumbers( _
    ByVal leftValue As String, _
    ByVal rightValue As String) As Long

    Dim leftFirst As String
    Dim leftSecond As String
    Dim rightFirst As String
    Dim rightSecond As String
    Dim leftIsValid As Boolean
    Dim rightIsValid As Boolean

    leftIsValid = TryParseCompositeTableNumber(leftValue, leftFirst, leftSecond)
    rightIsValid = TryParseCompositeTableNumber(rightValue, rightFirst, rightSecond)
    If leftIsValid And Not rightIsValid Then
        CompareCompositeTableNumbers = -1
        Exit Function
    End If
    If Not leftIsValid And rightIsValid Then
        CompareCompositeTableNumbers = 1
        Exit Function
    End If
    If Not leftIsValid Then Exit Function

    CompareCompositeTableNumbers = CompareUnsignedIntegerText(leftFirst, rightFirst)
    If CompareCompositeTableNumbers = 0 Then
        CompareCompositeTableNumbers = CompareUnsignedIntegerText(leftSecond, rightSecond)
    End If
End Function

' 複合番号を桁あふれしない数値文字列2要素へ分解する
Private Function TryParseCompositeTableNumber( _
    ByVal value As String, _
    ByRef firstNumber As String, _
    ByRef secondNumber As String) As Boolean

    Dim parts As Variant

    parts = Split(Trim$(value), "-")
    If LBound(parts) <> 0 Or UBound(parts) <> 1 Then Exit Function
    firstNumber = Trim$(CStr(parts(0)))
    secondNumber = Trim$(CStr(parts(1)))
    If Not IsUnsignedIntegerText(firstNumber) Or _
        Not IsUnsignedIntegerText(secondNumber) Then Exit Function

    firstNumber = NormalizeUnsignedIntegerText(firstNumber)
    secondNumber = NormalizeUnsignedIntegerText(secondNumber)
    TryParseCompositeTableNumber = True
End Function

' 符号なし整数文字列を桁数、同桁なら辞書順で比較する
Private Function CompareUnsignedIntegerText( _
    ByVal leftValue As String, _
    ByVal rightValue As String) As Long

    If Len(leftValue) < Len(rightValue) Then
        CompareUnsignedIntegerText = -1
    ElseIf Len(leftValue) > Len(rightValue) Then
        CompareUnsignedIntegerText = 1
    Else
        CompareUnsignedIntegerText = Sgn(StrComp(leftValue, rightValue, vbBinaryCompare))
    End If
End Function

' 数値比較用に先頭ゼロを除去し、ゼロ自体は1文字残す
Private Function NormalizeUnsignedIntegerText(ByVal value As String) As String
    Dim index As Long

    index = 1
    Do While index < Len(value) And Mid$(value, index, 1) = "0"
        index = index + 1
    Loop
    NormalizeUnsignedIntegerText = Mid$(value, index)
End Function

' 文字列がASCII数字だけで構成されているか判定する
Private Function IsUnsignedIntegerText(ByVal value As String) As Boolean
    Dim index As Long
    Dim character As String

    If Len(value) = 0 Then Exit Function
    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        If character < "0" Or character > "9" Then Exit Function
    Next index
    IsUnsignedIntegerText = True
End Function

' アウトプット②の論理4項目へ外枠と行境界を設定
Private Sub ApplyOutputTwoBlockStyle( _
    ByVal ws As Worksheet, _
    ByVal startColumn As Long, _
    ByVal lastRow As Long)

    Dim noRange As Range
    Dim idRange As Range
    Dim nameRange As Range
    Dim numberRange As Range
    Dim logicalRange As Variant

    lastRow = MaxLong(lastRow, 3)
    Set noRange = ws.Range(ws.Cells(3, startColumn), ws.Cells(lastRow, startColumn + 1))
    Set idRange = ws.Range(ws.Cells(3, startColumn + 2), ws.Cells(lastRow, startColumn + 17))
    Set nameRange = ws.Range(ws.Cells(3, startColumn + 18), ws.Cells(lastRow, startColumn + 42))
    Set numberRange = ws.Range(ws.Cells(3, startColumn + 43), ws.Cells(lastRow, startColumn + 47))

    For Each logicalRange In Array(noRange, idRange, nameRange, numberRange)
        logicalRange.Borders.LineStyle = xlNone
        logicalRange.Interior.Pattern = xlPatternNone
        ApplyOuterBorder logicalRange
        ApplyInsideHorizontalBorder logicalRange
    Next logicalRange

    noRange.HorizontalAlignment = xlCenter
    idRange.HorizontalAlignment = xlLeft
    nameRange.HorizontalAlignment = xlLeft
    numberRange.HorizontalAlignment = xlLeft
    noRange.VerticalAlignment = xlCenter
    idRange.VerticalAlignment = xlCenter
    nameRange.VerticalAlignment = xlCenter
    numberRange.VerticalAlignment = xlCenter

    ws.Range(ws.Cells(3, startColumn), ws.Cells(3, startColumn + 47)).Interior.Color = OUTPUT_FILL_COLOR
    ws.Range(ws.Cells(3, startColumn), ws.Cells(3, startColumn + 47)).Font.Bold = False
    ApplyBottomBorder ws.Range(ws.Cells(3, startColumn), ws.Cells(3, startColumn + 1))
    ApplyBottomBorder ws.Range(ws.Cells(3, startColumn + 2), ws.Cells(3, startColumn + 17))
    ApplyBottomBorder ws.Range(ws.Cells(3, startColumn + 18), ws.Cells(3, startColumn + 42))
    ApplyBottomBorder ws.Range(ws.Cells(3, startColumn + 43), ws.Cells(3, startColumn + 47))
End Sub

' テーブル一覧の見出しと格子罫線を整える（データは保持）
Private Sub ApplyTableListLayout(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim tableRange As Range

    SetOutputCellText ws.Cells(1, 1), TableListIdHeader()
    SetOutputCellText ws.Cells(1, 2), TableListNameHeader()
    SetOutputCellText ws.Cells(1, 3), TableNumberHeader()
    lastRow = MaxLong(LastUsedRowInColumn(ws, 1), 1)
    Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 3))
    tableRange.Borders.LineStyle = xlNone
    ApplyOuterBorder tableRange
    ApplyInsideHorizontalBorder tableRange
    If tableRange.Columns.Count > 1 Then
        With tableRange.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = vbBlack
        End With
    End If
    tableRange.Interior.Pattern = xlPatternNone
    ws.Range("A1:C1").Interior.Color = OUTPUT_FILL_COLOR
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A:C").Font.Name = OutputFontName()
    ws.Range("A:C").Font.Size = OutputFontSize()
    ws.Range("A:C").WrapText = False
    ws.Range("A:C").ShrinkToFit = False
    ws.Columns("C").NumberFormat = "@"
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 24
    ws.Columns("C").ColumnWidth = 12
    ApplyOutputSheetView ws
End Sub

' テーブル一覧を大文字小文字を無視するマスターへ読み込み、重複を収集
Private Sub LoadTableMaster( _
    ByVal ws As Worksheet, _
    ByRef tableMaster As Object, _
    ByRef duplicateTableIds As Collection)

    Dim duplicateSeen As Object
    Dim rowNumber As Long
    Dim lastRow As Long
    Dim tableId As String

    Set tableMaster = CreateCaseInsensitiveTextDictionary()
    Set duplicateSeen = CreateCaseInsensitiveTextDictionary()
    Set duplicateTableIds = New Collection
    lastRow = LastUsedRowInColumn(ws, 1)

    For rowNumber = 2 To lastRow
        tableId = NormalizePhysicalTableId(CStr(ws.Cells(rowNumber, 1).Value))
        If Len(tableId) > 0 Then
            If tableMaster.Exists(tableId) Then
                If Not duplicateSeen.Exists(tableId) Then
                    duplicateSeen.Add tableId, True
                    duplicateTableIds.Add tableId
                End If
            Else
                tableMaster.Add tableId, Array( _
                    tableId, _
                    CStr(ws.Cells(rowNumber, 2).Value), _
                    CStr(ws.Cells(rowNumber, 3).Value))
            End If
        End If
    Next rowNumber
End Sub

' 物理テーブルIDの前後空白と外側の角括弧を正規化
Private Function NormalizePhysicalTableId(ByVal tableId As String) As String
    tableId = Trim$(tableId)
    If Len(tableId) >= 2 And Left$(tableId, 1) = "[" And Right$(tableId, 1) = "]" Then
        tableId = Replace(Mid$(tableId, 2, Len(tableId) - 2), "]]", "]")
    End If
    NormalizePhysicalTableId = Trim$(tableId)
End Function

' テーブルID照合用の大文字小文字を無視する辞書を作成
Private Function CreateCaseInsensitiveTextDictionary() As Object
    Set CreateCaseInsensitiveTextDictionary = CreateObject("Scripting.Dictionary")
    CreateCaseInsensitiveTextDictionary.CompareMode = vbTextCompare
End Function

' テーブル一覧の重複を解析継続方針とともに1メッセージへまとめる
Public Function DuplicateTableWarningMessage(ByVal duplicateTableIds As Collection) As String
    Dim tableId As Variant
    Dim messageText As String

    messageText = W(&H30C6, &H30FC, &H30D6, &H30EB, &H4E00, &H89A7, &H306B, &H91CD, &H8907, &H3059, &H308B) & _
        TableListIdHeader() & W(&H304C, &H3042, &H308A, &H307E, &H3059, &H3002) & vbCrLf & _
        W(&H5148, &H982D, &H884C, &H3092, &H4F7F, &H7528, &H3057, &H3066, &H89E3, &H6790, &H3092, &H7D9A, &H884C, &H3057, &H307E, &H3059, &H3002)
    For Each tableId In duplicateTableIds
        messageText = messageText & vbCrLf & "- " & CStr(tableId)
    Next tableId

    DuplicateTableWarningMessage = messageText
End Function

' parserが返した全論理行をSQL解析シートの元の入力行単位へ再結合
Private Sub ApplyTransformedQueryLines( _
    ByVal wsSql As Worksheet, _
    ByVal lastRow As Long, _
    ByVal queryLineRows As Collection, _
    ByVal transformedQueryLines As Collection)

    Dim transformedByLine As Object
    Dim transformedByRow As Object
    Dim lineRecord As Variant
    Dim transformedData() As Variant
    Dim queryLine As Long
    Dim rowNumber As Long
    Dim rowKey As String

    If lastRow < 2 Then Exit Sub

    Set transformedByLine = CreateTextDictionary()
    Set transformedByRow = CreateTextDictionary()
    For Each lineRecord In transformedQueryLines
        transformedByLine(CStr(CLng(lineRecord(0)))) = CStr(lineRecord(1))
    Next lineRecord

    For queryLine = 1 To queryLineRows.Count
        rowNumber = CLng(queryLineRows.Item(queryLine))
        rowKey = CStr(rowNumber)
        If transformedByRow.Exists(rowKey) Then
            transformedByRow(rowKey) = CStr(transformedByRow(rowKey)) & vbLf & _
                CStr(transformedByLine(CStr(queryLine)))
        Else
            transformedByRow.Add rowKey, CStr(transformedByLine(CStr(queryLine)))
        End If
    Next queryLine

    ReDim transformedData(1 To lastRow - 1, 1 To 1)
    For rowNumber = 2 To lastRow
        rowKey = CStr(rowNumber)
        If transformedByRow.Exists(rowKey) Then
            transformedData(rowNumber - 1, 1) = _
                OutputTextValue(CStr(transformedByRow(rowKey)))
        End If
    Next rowNumber

    With wsSql.Range(wsSql.Cells(2, COL_RESULT), wsSql.Cells(lastRow, COL_RESULT))
        .NumberFormat = "@"
        .Value = transformedData
    End With
End Sub

' parserが返した変換値一式で行別のC列出力を置き換え
Private Sub ApplyFinalReplacementValues( _
    ByVal replacementValuesByRow As Object, _
    ByVal queryLineRows As Collection, _
    ByVal finalReplacementValues As Collection)

    Dim rowDictionary As Object
    Dim valueRecord As Variant
    Dim rowKey As Variant
    Dim queryLine As Long
    Dim valueOrder As Long

    ' v5のV行はC列へ出す最終値の完全な集合として扱う
    For Each rowKey In replacementValuesByRow.Keys
        Set rowDictionary = replacementValuesByRow(CStr(rowKey))
        rowDictionary.RemoveAll
    Next rowKey

    For Each valueRecord In finalReplacementValues
        queryLine = CLng(valueRecord(0))
        rowKey = CStr(queryLineRows.Item(queryLine))
        If replacementValuesByRow.Exists(CStr(rowKey)) Then
            Set rowDictionary = replacementValuesByRow(CStr(rowKey))
        Else
            Set rowDictionary = CreateTextDictionary()
            Set replacementValuesByRow(CStr(rowKey)) = rowDictionary
        End If
        valueOrder = valueOrder + 1
        AddReplacementValue rowDictionary, CStr(valueRecord(2)), valueOrder
    Next valueRecord
End Sub

' parserが補完したプレフィックスを行別の変換内容へ反映
Private Sub ApplyReplacementQualifications( _
    ByVal replacementValuesByRow As Object, _
    ByVal queryLineRows As Collection, _
    ByVal qualifications As Collection)

    Dim qualification As Variant
    Dim queryLine As Long
    Dim rowKey As String
    Dim replacementValues As Object

    For Each qualification In qualifications
        queryLine = CLng(qualification(0))
        If queryLine <= queryLineRows.Count Then
            rowKey = CStr(queryLineRows.Item(queryLine))
            If replacementValuesByRow.Exists(rowKey) Then
                Set replacementValues = replacementValuesByRow(rowKey)
                ApplyReplacementQualification _
                    replacementValues, CStr(qualification(2)), CStr(qualification(3))
            End If
        End If
    Next qualification
End Sub

' 変換内容の初出位置を維持しながら補完前の値を補完後の値へ統合
Private Sub ApplyReplacementQualification( _
    ByVal replacementValues As Object, _
    ByVal originalValue As String, _
    ByVal qualifiedValue As String)

    Dim firstMatchIndex As Long

    If StrComp(originalValue, qualifiedValue, vbBinaryCompare) = 0 Then Exit Sub
    If Not replacementValues.Exists(originalValue) Then Exit Sub

    firstMatchIndex = CLng(replacementValues(originalValue))
    replacementValues.Remove originalValue
    AddReplacementValue replacementValues, qualifiedValue, firstMatchIndex
End Sub

' CRLF、CR、LFのいずれでも文字列内の行数を数える
Private Function TextLineCount(ByVal value As String) As Long
    Dim normalizedText As String

    normalizedText = Replace(value, vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)
    TextLineCount = UBound(Split(normalizedText, vbLf)) + 1
End Function

' 数式判定と先頭アポストロフィの消費を避けて文字列を書き込む
Private Sub SetOutputCellText(ByVal targetCell As Range, ByVal cellValue As String)
    targetCell.NumberFormat = "@"
    targetCell.Value = OutputTextValue(cellValue)
End Sub

' Preserve a leading apostrophe when values are assigned to a text-formatted range.
Private Function OutputTextValue(ByVal cellValue As String) As String
    If Left$(cellValue, 1) = "'" Then
        cellValue = "'" & cellValue
    End If
    OutputTextValue = cellValue
End Function

' セクション種別に応じて塗りと外枠を設定
Private Sub ApplyOutputSectionStyle( _
    ByVal ws As Worksheet, _
    ByVal sectionKind As String, _
    ByVal startRow As Long, _
    ByVal endRow As Long)

    Select Case UCase$(sectionKind)
        Case "REFERENCE"
            ' タイトルと参照テーブル行は表の外側として扱う
            ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, OUTPUT_LAST_COLUMN)).Borders.LineStyle = xlNone
        Case "STANDARD"
            ApplyFilledFrame ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, 6)), OUTPUT_FILL_COLOR
            ApplyFilledFrame ws.Range(ws.Cells(startRow, 7), ws.Cells(endRow, OUTPUT_LAST_COLUMN)), vbWhite
        Case "TRANSFER"
            ApplyTransferSectionStyle ws, startRow, endRow
        Case "TRANSFER_GROUP"
            ApplyTransferGroupStyle ws, startRow, endRow
        Case "SEPARATOR"
            ApplySeparatorBorder ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, OUTPUT_LAST_COLUMN))
        Case Else
            Err.Raise vbObjectError + 520, "ApplyOutputSectionStyle", "Unknown section kind: " & sectionKind
    End Select
End Sub

' データ移送表の同一項目に属する複数行を1つの枠へまとめる
Private Sub ApplyTransferGroupStyle(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long)
    Dim leftRange As Range
    Dim middleRange As Range
    Dim rightRange As Range

    Set leftRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, 18))
    Set middleRange = ws.Range(ws.Cells(startRow, 19), ws.Cells(endRow, 36))
    Set rightRange = ws.Range(ws.Cells(startRow, 37), ws.Cells(endRow, OUTPUT_LAST_COLUMN))

    ClearInsideHorizontalBorder leftRange
    ClearInsideHorizontalBorder middleRange
    ClearInsideHorizontalBorder rightRange
    ApplyOuterBorder leftRange
    ApplyOuterBorder middleRange
    ApplyOuterBorder rightRange
End Sub

' 指定範囲の行間罫線を解除
Private Sub ClearInsideHorizontalBorder(ByVal targetRange As Range)
    If targetRange.Rows.Count <= 1 Then Exit Sub
    targetRange.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

' データ移送表の3列フレームと見出し色を設定
Private Sub ApplyTransferSectionStyle(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long)
    Dim leftRange As Range
    Dim middleRange As Range
    Dim rightRange As Range

    Set leftRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, 18))
    Set middleRange = ws.Range(ws.Cells(startRow, 19), ws.Cells(endRow, 36))
    Set rightRange = ws.Range(ws.Cells(startRow, 37), ws.Cells(endRow, OUTPUT_LAST_COLUMN))

    ApplyFilledFrame leftRange, vbWhite
    ApplyFilledFrame middleRange, vbWhite
    ApplyFilledFrame rightRange, vbWhite
    ApplyInsideHorizontalBorder leftRange
    ApplyInsideHorizontalBorder middleRange
    ApplyInsideHorizontalBorder rightRange
    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 18)).Interior.Color = OUTPUT_FILL_COLOR
    ws.Range(ws.Cells(startRow, 19), ws.Cells(startRow, 36)).Interior.Color = OUTPUT_FILL_COLOR
    ws.Range(ws.Cells(startRow, 37), ws.Cells(startRow, OUTPUT_LAST_COLUMN)).Interior.Color = OUTPUT_FILL_COLOR
    ApplyBottomBorder ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 18))
    ApplyBottomBorder ws.Range(ws.Cells(startRow, 19), ws.Cells(startRow, 36))
    ApplyBottomBorder ws.Range(ws.Cells(startRow, 37), ws.Cells(startRow, OUTPUT_LAST_COLUMN))
End Sub

' 指定範囲の行間へ最細の黒い罫線を設定
Private Sub ApplyInsideHorizontalBorder(ByVal targetRange As Range)
    If targetRange.Rows.Count <= 1 Then Exit Sub

    With targetRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = vbBlack
    End With
End Sub

' 指定範囲を塗り、最細の黒い外枠を設定
Private Sub ApplyFilledFrame(ByVal targetRange As Range, ByVal fillColor As Long)
    If fillColor <> vbWhite Then targetRange.Interior.Color = fillColor
    ApplyOuterBorder targetRange
End Sub

' 指定範囲へ最細の黒い外枠を設定
Private Sub ApplyOuterBorder(ByVal targetRange As Range)
    Dim borderIndex As Variant

    For Each borderIndex In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        With targetRange.Borders(CLng(borderIndex))
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = vbBlack
        End With
    Next borderIndex
End Sub

' データ移送表の見出し行へ下罫線を設定
Private Sub ApplyBottomBorder(ByVal targetRange As Range)
    With targetRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = vbBlack
    End With
End Sub

' UNIONなどの境界行へ上下罫線を設定
Private Sub ApplySeparatorBorder(ByVal targetRange As Range)
    Dim borderIndex As Variant

    targetRange.Interior.Color = vbWhite
    For Each borderIndex In Array(xlEdgeTop, xlEdgeBottom)
        With targetRange.Borders(CLng(borderIndex))
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = vbBlack
        End With
    Next borderIndex
End Sub

' parser行形式で使用する制御文字をエスケープ
Private Function EscapeProtocolField(ByVal value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, vbCr, "\r")
    value = Replace(value, vbLf, "\n")
    value = Replace(value, vbTab, "\t")
    EscapeProtocolField = value
End Function

' parser行形式のエスケープを元へ戻す
Private Function UnescapeProtocolField(ByVal value As String) As String
    Dim resultText As String
    Dim currentChar As String
    Dim escapedChar As String
    Dim index As Long

    index = 1
    Do While index <= Len(value)
        currentChar = Mid$(value, index, 1)
        If currentChar = "\" And index < Len(value) Then
            escapedChar = Mid$(value, index + 1, 1)
            Select Case escapedChar
                Case "r": resultText = resultText & vbCr
                Case "n": resultText = resultText & vbLf
                Case "t": resultText = resultText & vbTab
                Case "\": resultText = resultText & "\"
                Case Else: resultText = resultText & "\" & escapedChar
            End Select
            index = index + 2
        Else
            resultText = resultText & currentChar
            index = index + 1
        End If
    Loop

    UnescapeProtocolField = resultText
End Function

' フォールバックSQLを行単位で出力し末尾へ原因を追加
Private Sub WriteFallbackOutput(ByVal wsOutput As Worksheet, ByVal queryText As String, ByVal reason As String)
    Dim lines As Variant
    Dim normalizedText As String
    Dim lineIndex As Long
    Dim reasonRow As Long
    Dim queryEndRow As Long

    ClearOutputSheet wsOutput, False
    normalizedText = Replace(queryText, vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)
    normalizedText = TrimOuterLineBreaks(normalizedText)
    lines = Split(normalizedText, vbLf)
    wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(UBound(lines) - LBound(lines) + 1, 1)).NumberFormat = "@"

    For lineIndex = LBound(lines) To UBound(lines)
        wsOutput.Cells(lineIndex - LBound(lines) + 1, 1).Value = CStr(lines(lineIndex))
    Next lineIndex

    queryEndRow = UBound(lines) - LBound(lines) + 1
    reasonRow = queryEndRow + 2
    wsOutput.Cells(reasonRow, 1).Value = FallbackReasonPrefix() & reason & _
        FallbackQueryLocation(1, queryEndRow)
    ApplyOutputSheetDimensions wsOutput, reasonRow
    ApplyOutputSheetFont wsOutput, reasonRow
    ApplyOutputSheetView wsOutput
End Sub

' 文字列の先頭と末尾にある改行だけを除去
Private Function TrimOuterLineBreaks(ByVal value As String) As String
    Do While Len(value) > 0 And Left$(value, 1) = vbLf
        value = Mid$(value, 2)
    Loop
    Do While Len(value) > 0 And Right$(value, 1) = vbLf
        value = Left$(value, Len(value) - 1)
    Loop

    TrimOuterLineBreaks = value
End Function

' アウトプットシートへクエリブロックを順に出力
Private Function WriteOutputQueryBlocks(ByVal wsOutput As Worksheet, ByVal startRow As Long, ByVal queryText As String) As Long
    Dim blocks As Collection
    Dim blockText As Variant
    Dim rowNumber As Long

    Set blocks = BuildOutputQueryBlocks(queryText)
    rowNumber = startRow
    For Each blockText In blocks
        wsOutput.Cells(rowNumber, 1).Value = CStr(blockText)
        rowNumber = rowNumber + 1
    Next blockText

    WriteOutputQueryBlocks = rowNumber
End Function

' サブクエリを内側から並べ、最後にクエリ全体を追加
Private Function BuildOutputQueryBlocks(ByVal queryText As String) As Collection
    Dim blocks As Collection

    Set blocks = TryBuildExternalOutputBlocks(queryText)
    If Not blocks Is Nothing Then
        Set BuildOutputQueryBlocks = blocks
        Exit Function
    End If

    Set blocks = New Collection
    CollectSubqueryBlocks queryText, blocks
    blocks.Add NormalizeOutputQueryBlock(queryText)

    Set BuildOutputQueryBlocks = blocks
End Function

' 外部parserでアウトプット用ブロックを作成
Private Function TryBuildExternalOutputBlocks(ByVal queryText As String) As Collection
    Dim parserPath As String
    Dim inputPath As String
    Dim outputPath As String
    Dim exitCode As Long
    Dim outputText As String

    On Error GoTo CleanUp

    parserPath = ResolveParserExePath()
    If Len(parserPath) = 0 Then
        Exit Function
    End If

    inputPath = TemporaryFilePath("saf_sql_", ".sql")
    outputPath = TemporaryFilePath("saf_blocks_", ".txt")
    WriteUtf8TextFile inputPath, queryText

    exitCode = RunParserProcess(parserPath, inputPath, outputPath)
    If exitCode <> 0 Or Not FileExists(outputPath) Then
        GoTo CleanUp
    End If

    outputText = ReadUnicodeTextFile(outputPath)
    Set TryBuildExternalOutputBlocks = SplitOutputBlocks(outputText)

CleanUp:
    DeleteFileIfExists inputPath
    DeleteFileIfExists outputPath
End Function

' parser exeの配置先を解決
Private Function ResolveParserExePath() As String
    Dim envPath As String
    Dim basePath As String
    Dim candidatePath As Variant

    envPath = Environ$("SQL_ANALYSIS_FORMATTER_PARSER_EXE")
    If Len(envPath) > 0 And FileExists(envPath) Then
        ResolveParserExePath = envPath
        Exit Function
    End If

    basePath = ThisWorkbook.Path
    For Each candidatePath In Array( _
        basePath & Application.PathSeparator & "SqlAnalysisFormatter.Parser.exe", _
        basePath & Application.PathSeparator & "tools" & Application.PathSeparator & "SqlAnalysisFormatter.Parser.exe", _
        basePath & Application.PathSeparator & "dist" & Application.PathSeparator & "parser" & Application.PathSeparator & "SqlAnalysisFormatter.Parser.exe")
        If FileExists(CStr(candidatePath)) Then
            ResolveParserExePath = CStr(candidatePath)
            Exit Function
        End If
    Next candidatePath
End Function

' parser exeを同期実行
Private Function RunParserProcess(ByVal parserPath As String, ByVal inputPath As String, ByVal outputPath As String) As Long
    Dim shell As Object
    Dim commandText As String

    commandText = QuoteCommandArgument(parserPath) & _
        " --input " & QuoteCommandArgument(inputPath) & _
        " --output " & QuoteCommandArgument(outputPath) & _
        " --format vba-blocks"

    Set shell = CreateObject("WScript.Shell")
    RunParserProcess = CLng(shell.Run(commandText, 0, True))
End Function

' parser出力をブロック単位へ分割
Private Function SplitOutputBlocks(ByVal outputText As String) As Collection
    Dim blocks As Collection
    Dim values As Variant
    Dim index As Long

    Set blocks = New Collection
    values = Split(outputText, OutputBlockSeparator())
    For index = LBound(values) To UBound(values)
        blocks.Add CStr(values(index))
    Next index

    Set SplitOutputBlocks = blocks
End Function

' UTF-8でテキストファイルへ書き込み
Private Sub WriteUtf8TextFile(ByVal filePath As String, ByVal contentText As String)
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText contentText
    stream.SaveToFile filePath, 2
    stream.Close
End Sub

' UTF-16でテキストファイルを読み込み
Private Function ReadUnicodeTextFile(ByVal filePath As String) As String
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "unicode"
    stream.Open
    stream.LoadFromFile filePath
    ReadUnicodeTextFile = stream.ReadText(-1)
    stream.Close
End Function

' 一時ファイルパスを作成
Private Function TemporaryFilePath(ByVal prefixText As String, ByVal extensionText As String) As String
    Dim fso As Object
    Dim tempName As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    tempName = Replace(fso.GetTempName(), ".", "_")
    TemporaryFilePath = fso.BuildPath(Environ$("TEMP"), prefixText & tempName & extensionText)
End Function

' コマンドライン引数を引用符で囲む
Private Function QuoteCommandArgument(ByVal value As String) As String
    QuoteCommandArgument = """" & Replace(value, """", """""") & """"
End Function

' ファイルが存在するか判定
Private Function FileExists(ByVal filePath As String) As Boolean
    If Len(filePath) = 0 Then Exit Function
    FileExists = (Len(Dir$(filePath, vbNormal)) > 0)
End Function

' ファイルがあれば削除
Private Sub DeleteFileIfExists(ByVal filePath As String)
    If Len(filePath) = 0 Then Exit Sub
    On Error Resume Next
    Kill filePath
    On Error GoTo 0
End Sub

' 括弧内のSELECT/WITHをサブクエリとして収集
Private Sub CollectSubqueryBlocks(ByVal queryText As String, ByVal blocks As Collection)
    Dim index As Long
    Dim closingIndex As Long
    Dim blockText As String
    Dim normalizedBlock As String
    Dim currentChar As String

    index = 1
    Do While index <= Len(queryText)
        currentChar = Mid$(queryText, index, 1)
        If currentChar = "'" Then
            index = PositionAfterSqlString(queryText, index)
        ElseIf StartsWithAt(queryText, index, "--") Then
            index = PositionAfterLineComment(queryText, index)
        ElseIf StartsWithAt(queryText, index, "/*") Then
            index = PositionAfterBlockComment(queryText, index)
        ElseIf currentChar = "(" Then
            closingIndex = MatchingClosingParenthesis(queryText, index)
            If closingIndex = 0 Then
                index = index + 1
            Else
                blockText = Mid$(queryText, index + 1, closingIndex - index - 1)
                CollectSubqueryBlocks blockText, blocks
                normalizedBlock = NormalizeOutputQueryBlock(blockText)
                If IsSubqueryBlock(normalizedBlock) Then
                    blocks.Add normalizedBlock
                End If
                index = closingIndex + 1
            End If
        Else
            index = index + 1
        End If
    Loop
End Sub

' アウトプット用に前後空白だけを除去
Private Function NormalizeOutputQueryBlock(ByVal queryText As String) As String
    NormalizeOutputQueryBlock = TrimSqlWhitespace(queryText)
End Function

' サブクエリとして扱うブロックか判定
Private Function IsSubqueryBlock(ByVal queryText As String) As Boolean
    IsSubqueryBlock = StartsWithSqlToken(queryText, "SELECT") Or StartsWithSqlToken(queryText, "WITH")
End Function

' SQL先頭トークンが指定語か判定
Private Function StartsWithSqlToken(ByVal queryText As String, ByVal tokenText As String) As Boolean
    Dim trimmedText As String
    Dim nextChar As String

    trimmedText = TrimSqlWhitespace(queryText)
    If Len(trimmedText) < Len(tokenText) Then Exit Function
    If UCase$(Left$(trimmedText, Len(tokenText))) <> tokenText Then Exit Function
    If Len(trimmedText) = Len(tokenText) Then
        StartsWithSqlToken = True
        Exit Function
    End If

    nextChar = Mid$(trimmedText, Len(tokenText) + 1, 1)
    StartsWithSqlToken = Not IsIdentifierCharacter(nextChar)
End Function

' SQL上の前後空白を除去
Private Function TrimSqlWhitespace(ByVal sourceText As String) As String
    Dim startIndex As Long
    Dim endIndex As Long

    startIndex = 1
    Do While startIndex <= Len(sourceText)
        If Not IsWhitespace(Mid$(sourceText, startIndex, 1)) Then Exit Do
        startIndex = startIndex + 1
    Loop

    endIndex = Len(sourceText)
    Do While endIndex >= startIndex
        If Not IsWhitespace(Mid$(sourceText, endIndex, 1)) Then Exit Do
        endIndex = endIndex - 1
    Loop

    If endIndex >= startIndex Then
        TrimSqlWhitespace = Mid$(sourceText, startIndex, endIndex - startIndex + 1)
    End If
End Function

' 指定位置から始まる文字列か判定
Private Function StartsWithAt(ByVal sourceText As String, ByVal startIndex As Long, ByVal searchText As String) As Boolean
    If startIndex + Len(searchText) - 1 > Len(sourceText) Then Exit Function
    StartsWithAt = (Mid$(sourceText, startIndex, Len(searchText)) = searchText)
End Function

' 対応する閉じ括弧の位置を取得
Private Function MatchingClosingParenthesis(ByVal sourceText As String, ByVal openingIndex As Long) As Long
    Dim index As Long
    Dim depth As Long
    Dim currentChar As String

    index = openingIndex
    Do While index <= Len(sourceText)
        currentChar = Mid$(sourceText, index, 1)
        If currentChar = "'" Then
            index = PositionAfterSqlString(sourceText, index)
        ElseIf StartsWithAt(sourceText, index, "--") Then
            index = PositionAfterLineComment(sourceText, index)
        ElseIf StartsWithAt(sourceText, index, "/*") Then
            index = PositionAfterBlockComment(sourceText, index)
        ElseIf currentChar = "(" Then
            depth = depth + 1
            index = index + 1
        ElseIf currentChar = ")" Then
            depth = depth - 1
            If depth = 0 Then
                MatchingClosingParenthesis = index
                Exit Function
            End If
            index = index + 1
        Else
            index = index + 1
        End If
    Loop
End Function

' 文字列リテラルの直後の位置を取得
Private Function PositionAfterSqlString(ByVal sourceText As String, ByVal quoteIndex As Long) As Long
    Dim index As Long

    index = quoteIndex + 1
    Do While index <= Len(sourceText)
        If Mid$(sourceText, index, 1) = "'" Then
            If index < Len(sourceText) And Mid$(sourceText, index + 1, 1) = "'" Then
                index = index + 2
            Else
                PositionAfterSqlString = index + 1
                Exit Function
            End If
        Else
            index = index + 1
        End If
    Loop

    PositionAfterSqlString = Len(sourceText) + 1
End Function

' 行コメントの直後の位置を取得
Private Function PositionAfterLineComment(ByVal sourceText As String, ByVal commentIndex As Long) As Long
    Dim newlineIndex As Long

    newlineIndex = InStr(commentIndex + 2, sourceText, vbLf, vbBinaryCompare)
    If newlineIndex = 0 Then
        PositionAfterLineComment = Len(sourceText) + 1
    Else
        PositionAfterLineComment = newlineIndex + 1
    End If
End Function

' ブロックコメントの直後の位置を取得
Private Function PositionAfterBlockComment(ByVal sourceText As String, ByVal commentIndex As Long) As Long
    Dim closeIndex As Long

    closeIndex = InStr(commentIndex + 2, sourceText, "*/", vbBinaryCompare)
    If closeIndex = 0 Then
        PositionAfterBlockComment = Len(sourceText) + 1
    Else
        PositionAfterBlockComment = closeIndex + 2
    End If
End Function

' Excel検索ダイアログの検索方向を行へ戻す
Private Sub RestoreFindSearchOrderByRows(ByVal ws As Worksheet)
    Dim foundCell As Range

    Application.FindFormat.Clear
    Set foundCell = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        MatchByte:=False, _
        SearchFormat:=False)
End Sub

' アウトプットシートの内容と前回の表書式をクリア
Private Sub ClearOutputSheet(ByVal ws As Worksheet, Optional ByVal applyLayout As Boolean = True)
    Dim clearLastRow As Long

    clearLastRow = LastOutputRow(ws)
    With ws.Range(ws.Cells(1, 1), ws.Cells(MaxLong(clearLastRow, 1), OUTPUT_LAST_COLUMN))
        .Clear
    End With
    If applyLayout Then ApplyOutputSheetLayout ws
End Sub

' アウトプット範囲の最終使用行を取得
Private Function LastOutputRow( _
    ByVal ws As Worksheet, _
    Optional ByVal lastColumn As Long = OUTPUT_LAST_COLUMN) As Long
    Dim outputRange As Range
    Dim foundCell As Range

    LastOutputRow = 1
    ' 値が存在し得る使用範囲だけへ絞り、全104万行の反復検索を避ける
    Set outputRange = Intersect( _
        ws.UsedRange, _
        ws.Range(ws.Columns(1), ws.Columns(lastColumn)))
    If outputRange Is Nothing Then Exit Function

    Set foundCell = outputRange.Find( _
        What:="*", _
        After:=outputRange.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False, _
        MatchByte:=False, _
        SearchFormat:=False)
    If Not foundCell Is Nothing Then
        LastOutputRow = foundCell.Row
    End If
End Function

' 指定シートの2行目以降を使用範囲に合わせてクリア
Private Sub ClearRowsBelowHeader(ByVal ws As Worksheet, ByVal minimumLastColumn As Long)
    Dim lastRow As Long
    Dim lastColumn As Long

    lastRow = LastUsedRow(ws)
    lastColumn = MaxLong(LastUsedColumn(ws), minimumLastColumn)
    If lastRow >= 2 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastColumn)).ClearContents
    End If
End Sub

' 指定列の最終使用行を取得
Private Function LastUsedRowInColumn(ByVal ws As Worksheet, ByVal columnNumber As Long) As Long
    Dim rowNumber As Long

    rowNumber = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row
    If rowNumber = 1 And Len(CStr(ws.Cells(1, columnNumber).Value)) = 0 Then
        LastUsedRowInColumn = 1
    Else
        LastUsedRowInColumn = rowNumber
    End If
End Function

' シート全体の最終使用行を取得
Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim dataRange As Range
    Dim area As Range
    Dim areaLastRow As Long

    LastUsedRow = 1
    Set dataRange = UsedValueCells(ws)
    If dataRange Is Nothing Then Exit Function

    For Each area In dataRange.Areas
        areaLastRow = area.Row + area.Rows.Count - 1
        If areaLastRow > LastUsedRow Then
            LastUsedRow = areaLastRow
        End If
    Next area
End Function

' シート全体の最終使用列を取得
Private Function LastUsedColumn(ByVal ws As Worksheet) As Long
    Dim dataRange As Range
    Dim area As Range
    Dim areaLastColumn As Long

    LastUsedColumn = 1
    Set dataRange = UsedValueCells(ws)
    If dataRange Is Nothing Then Exit Function

    For Each area In dataRange.Areas
        areaLastColumn = area.Column + area.Columns.Count - 1
        If areaLastColumn > LastUsedColumn Then
            LastUsedColumn = areaLastColumn
        End If
    Next area
End Function

' Find設定を汚さず、値または数式が入っているセルだけを取得
Private Function UsedValueCells(ByVal ws As Worksheet) As Range
    Dim constantCells As Range
    Dim formulaCells As Range

    On Error Resume Next
    Set constantCells = ws.Cells.SpecialCells(xlCellTypeConstants)
    Set formulaCells = ws.Cells.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If constantCells Is Nothing Then
        Set UsedValueCells = formulaCells
    ElseIf formulaCells Is Nothing Then
        Set UsedValueCells = constantCells
    Else
        Set UsedValueCells = Union(constantCells, formulaCells)
    End If
End Function

' 2つのLong値の大きい方を取得
Private Function MaxLong(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    If leftValue >= rightValue Then
        MaxLong = leftValue
    Else
        MaxLong = rightValue
    End If
End Function

' Unicodeコードポイント列から文字列を生成
Private Function W(ParamArray codes() As Variant) As String
    Dim index As Long
    Dim resultText As String

    ' VBEインポート時の文字化けを避けるため、UI文字列はコードポイントで保持
    For index = LBound(codes) To UBound(codes)
        resultText = resultText & ChrW$(CLng(codes(index)))
    Next index

    W = resultText
End Function

' 変換定義シート名を取得
Private Function ReferenceSheetName() As String
    ReferenceSheetName = W(&H5909, &H63DB, &H5B9A, &H7FA9)
End Function

' SQL解析シート名を取得
Private Function SqlSheetName() As String
    SqlSheetName = "SQL" & W(&H89E3, &H6790)
End Function

' アウトプット①シート名を取得
Private Function OutputSheetName() As String
    OutputSheetName = LegacyOutputSheetName() & ChrW$(&H2460)
End Function

' 旧アウトプットシート名を取得
Private Function LegacyOutputSheetName() As String
    LegacyOutputSheetName = W(&H30A2, &H30A6, &H30C8, &H30D7, &H30C3, &H30C8)
End Function

' アウトプット②シート名を取得
Private Function OutputSheetTwoName() As String
    OutputSheetTwoName = LegacyOutputSheetName() & ChrW$(&H2461)
End Function

' テーブル一覧シート名を取得
Private Function TableListSheetName() As String
    TableListSheetName = W(&H30C6, &H30FC, &H30D6, &H30EB, &H4E00, &H89A7)
End Function

' 入力情報タイトルを取得
Private Function InputInformationTitle() As String
    InputInformationTitle = ChrW$(&HFF1C) & W(&H5165, &H529B, &H60C5, &H5831) & ChrW$(&HFF1E)
End Function

' 出力情報タイトルを取得
Private Function OutputInformationTitle() As String
    OutputInformationTitle = ChrW$(&HFF1C) & W(&H51FA, &H529B, &H60C5, &H5831) & ChrW$(&HFF1E)
End Function

' テーブル一覧のテーブルID見出しを取得
Private Function TableListIdHeader() As String
    TableListIdHeader = W(&H30C6, &H30FC, &H30D6, &H30EB) & "ID"
End Function

' テーブル一覧のテーブル名称見出しを取得
Private Function TableListNameHeader() As String
    TableListNameHeader = W(&H30C6, &H30FC, &H30D6, &H30EB, &H540D, &H79F0)
End Function

' テーブル一覧の番号見出しを取得
Private Function TableNumberHeader() As String
    TableNumberHeader = W(&H756A, &H53F7)
End Function

' アウトプットシートのフォント名を取得
Private Function OutputFontName() As String
    OutputFontName = W(&HFF2D, &HFF33, &H20, &H30B4, &H30B7, &H30C3, &H30AF)
End Function

' アウトプットシートのフォントサイズを取得
Private Function OutputFontSize() As Long
    OutputFontSize = 9
End Function

' parser出力のブロック区切り文字を取得
Private Function OutputBlockSeparator() As String
    OutputBlockSeparator = ChrW$(&H1E)
End Function

' 所属テーブルID見出しを取得
Private Function TableIdHeader() As String
    TableIdHeader = W(&H6240, &H5C5E, &H30C6, &H30FC, &H30D6, &H30EB) & "ID"
End Function

' 所属テーブル和名見出しを取得
Private Function TableNameHeader() As String
    TableNameHeader = W(&H6240, &H5C5E, &H30C6, &H30FC, &H30D6, &H30EB, &H548C, &H540D)
End Function

' フィールドID見出しを取得
Private Function FieldIdHeader() As String
    FieldIdHeader = W(&H30D5, &H30A3, &H30FC, &H30EB, &H30C9) & "ID"
End Function

' フィールド和名見出しを取得
Private Function FieldNameHeader() As String
    FieldNameHeader = W(&H30D5, &H30A3, &H30FC, &H30EB, &H30C9, &H548C, &H540D)
End Function

' SQLクエリ見出しを取得
Private Function SqlHeader() As String
    SqlHeader = "SQL" & W(&H30AF, &H30A8, &H30EA)
End Function

' 変換後クエリ見出しを取得
Private Function ResultHeader() As String
    ResultHeader = W(&H5909, &H63DB, &H5F8C, &H30AF, &H30A8, &H30EA)
End Function

' 変換内容見出しを取得
Private Function ReplacementHeader() As String
    ReplacementHeader = W(&H5909, &H63DB, &H5185, &H5BB9)
End Function

' 解析ボタンの表示文字を取得
Private Function AnalyzeButtonText() As String
    AnalyzeButtonText = W(&H89E3, &H6790)
End Function

' クリアボタンの表示文字を取得
Private Function ClearButtonText() As String
    ClearButtonText = W(&H30AF, &H30EA, &H30A2)
End Function

' コピーボタンの表示文字を取得
Private Function CopyButtonText() As String
    CopyButtonText = W(&H30B3, &H30D4, &H30FC)
End Function

' フォールバック原因の見出しを取得
Private Function FallbackReasonPrefix() As String
    FallbackReasonPrefix = W(&H30D5, &H30A9, &H30FC, &H30EB, &H30D0, &H30C3, &H30AF, &H539F, &H56E0) & ": "
End Function

' フォールバック対象クエリのアウトプット行を取得
Private Function FallbackQueryLocation(ByVal startRow As Long, ByVal endRow As Long) As String
    Dim rowText As String

    rowText = CStr(startRow)
    If endRow <> startRow Then
        rowText = rowText & W(&HFF5E) & CStr(endRow)
    End If
    FallbackQueryLocation = W(&HFF08, &H5BFE, &H8C61, &H30AF, &H30A8, &H30EA) & ": " & _
        OutputSheetName() & " " & _
        rowText & W(&H884C, &H76EE, &HFF09)
End Function

' parser未配置時の原因を取得
Private Function ParserNotFoundReason() As String
    ParserNotFoundReason = "parser EXE" & W(&H304C, &H898B, &H3064, &H304B, &H308A, &H307E, &H305B, &H3093, &H3002)
End Function

' parser異常終了時の原因を取得
Private Function ParserExitCodeReason(ByVal exitCode As Long) As String
    ParserExitCodeReason = "parser EXE" & _
        W(&H306E, &H5B9F, &H884C, &H306B, &H5931, &H6557, &H3057, &H307E, &H3057, &H305F, &H3002, &H7D42, &H4E86, &H30B3, &H30FC, &H30C9) & _
        ": " & CStr(exitCode)
End Function

' parser出力ファイル未生成時の原因を取得
Private Function ParserOutputMissingReason() As String
    ParserOutputMissingReason = "parser EXE" & _
        W(&H306E, &H51FA, &H529B, &H30D5, &H30A1, &H30A4, &H30EB, &H304C, &H898B, &H3064, &H304B, &H308A, &H307E, &H305B, &H3093, &H3002)
End Function

' parser出力形式不正時の原因を取得
Private Function ParserOutputInvalidReason() As String
    ParserOutputInvalidReason = "parser EXE" & _
        W(&H306E, &H51FA, &H529B, &H5F62, &H5F0F, &H304C, &H4E0D, &H6B63, &H3067, &H3059, &H3002)
End Function

' parser連携例外の原因を取得
Private Function ParserIntegrationErrorReason(ByVal description As String) As String
    ParserIntegrationErrorReason = "parser EXE" & _
        W(&H3068, &H306E, &H9023, &H643A, &H4E2D, &H306B, &H30A8, &H30E9, &H30FC) & ": " & description
End Function

' 和名未取得判定用の文字列を取得
Private Function MissingNameText() As String
    MissingNameText = W(&H548C, &H540D, &H672A, &H53D6, &H5F97)
End Function

' parserが出力する参照テーブル見出しを取得
Private Function ReferenceTablesText() As String
    ReferenceTablesText = W(&H53C2, &H7167, &H30C6, &H30FC, &H30D6, &H30EB)
End Function

' 解析完了メッセージを取得
Private Function AnalyzeDoneMessage() As String
    AnalyzeDoneMessage = W(&H89E3, &H6790, &H304C, &H5B8C, &H4E86, &H3057, &H307E, &H3057, &H305F, &H3002)
End Function

' フォールバック理由を含む解析完了メッセージを取得
Private Function AnalyzeFallbackMessage(ByVal reason As String) As String
    AnalyzeFallbackMessage = AnalyzeDoneMessage() & vbCrLf & vbCrLf & _
        W(&H30A8, &H30E9, &H30FC, &H5185, &H5BB9) & ":" & vbCrLf & reason
End Function

' コピー完了メッセージを取得
Private Function CopyDoneMessage() As String
    CopyDoneMessage = W(&H30A2, &H30A6, &H30C8, &H30D7, &H30C3, &H30C8, &H3092, &H30AF, &H30EA, &H30C3, &H30D7, &H30DC, &H30FC, &H30C9, &H306B, &H30B3, &H30D4, &H30FC, &H3057, &H307E, &H3057, &H305F, &H3002)
End Function

' コピー対象なしメッセージを取得
Private Function NoOutputToCopyMessage() As String
    NoOutputToCopyMessage = W(&H30B3, &H30D4, &H30FC, &H3059, &H308B, &H6210, &H679C, &H7269, &H304C, &H3042, &H308A, &H307E, &H305B, &H3093, &H3002)
End Function

' コピー失敗メッセージを取得
Private Function CopyFailedMessage() As String
    CopyFailedMessage = W(&H30B3, &H30D4, &H30FC, &H306B, &H5931, &H6557, &H3057, &H307E, &H3057, &H305F) & ": "
End Function

' クリア完了メッセージを取得
Private Function ClearDoneMessage() As String
    ClearDoneMessage = W(&H30AF, &H30EA, &H30A2, &H304C, &H5B8C, &H4E86, &H3057, &H307E, &H3057, &H305F, &H3002)
End Function

' クリア確認メッセージを取得
Public Function ClearConfirmMessage() As String
    ClearConfirmMessage = W(&H89E3, &H6790, &H7D50, &H679C, &H3092, &H30AF, &H30EA, &H30A2, &H3057, &H307E, &H3059, &H3002, &H3088, &H308D, &H3057, &H3044, &H3067, &H3059, &H304B, &HFF1F)
End Function

' 確認ダイアログのタイトルを取得
Private Function ConfirmTitle() As String
    ConfirmTitle = W(&H78BA, &H8A8D)
End Function
