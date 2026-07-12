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

' ブックのシート名、見出し、操作ボタンを初期化
Public Sub SetupWorkbook()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet

    Set wsRef = ResolveOrCreateSheet(ReferenceSheetName(), REF_LEGACY_SHEET, 1)
    Set wsSql = ResolveOrCreateSheet(SqlSheetName(), SQL_LEGACY_SHEET, 2)
    Set wsOutput = ResolveOrCreateSheet(OutputSheetName(), OutputSheetName(), 3)

    RestoreHeaders wsRef, wsSql
    ApplyOutputSheetLayout wsOutput
    InstallButtons wsSql
End Sub

' SQL解析シートのA列を変換し、B列以降へ結果を出力
Public Sub AnalyzeQueries(Optional ByVal showMessage As Boolean = True)
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim qualifiedMap As Object
    Dim standaloneMap As Object
    Dim qualifiedKeys As Variant
    Dim standaloneKeys As Variant
    Dim lastRow As Long
    Dim outputRow As Long
    Dim rowNumber As Long
    Dim sourceText As String
    Dim convertedText As String
    Dim replacementValues As Object

    Set wsRef = GetReferenceSheet()
    Set wsSql = GetSqlSheet()
    Set wsOutput = GetOutputSheet()
    Set qualifiedMap = CreateTextDictionary()
    Set standaloneMap = CreateTextDictionary()

    LoadMappings wsRef, qualifiedMap, standaloneMap
    If qualifiedMap.Count = 0 And standaloneMap.Count = 0 Then
        If showMessage Then
            MsgBox NoDefinitionMessage(), vbInformation
        End If
        RestoreFindSearchOrderByRows wsSql
        Exit Sub
    End If

    qualifiedKeys = SortedKeysByLengthDesc(qualifiedMap)
    standaloneKeys = SortedKeysByLengthDesc(standaloneMap)

    lastRow = LastUsedRowInColumn(wsSql, COL_SQL)
    ClearAnalyzeOutput wsSql, lastRow
    ClearOutputSheet wsOutput
    outputRow = 1

    For rowNumber = 2 To lastRow
        sourceText = CStr(wsSql.Cells(rowNumber, COL_SQL).Value)
        If Len(sourceText) > 0 Then
            Set replacementValues = CreateTextDictionary()
            convertedText = ApplyMappings(sourceText, qualifiedMap, qualifiedKeys, standaloneMap, standaloneKeys, replacementValues)
            wsSql.Cells(rowNumber, COL_RESULT).Value = convertedText
            WriteReplacementValues wsSql, rowNumber, replacementValues
            outputRow = WriteOutputQueryBlocks(wsOutput, outputRow, convertedText)
        End If
    Next rowNumber

    wsSql.Columns(COL_RESULT).WrapText = False
    SetReplacementColumnsWrapText wsSql, False, LastUsedColumn(wsSql)
    If showMessage Then
        MsgBox AnalyzeDoneMessage(), vbInformation
    End If
    RestoreFindSearchOrderByRows wsSql
End Sub

' 確認後、入力シートの2行目以降とアウトプットシートをクリア
Public Sub ClearData(Optional ByVal showMessage As Boolean = True)
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet

    If showMessage Then
        If MsgBox(ClearConfirmMessage(), vbQuestion + vbYesNo + vbDefaultButton2, ConfirmTitle()) <> vbYes Then
            Exit Sub
        End If
    End If

    Set wsRef = GetReferenceSheet()
    Set wsSql = GetSqlSheet()
    Set wsOutput = GetOutputSheet()

    ClearRowsBelowHeader wsRef, COL_FIELD_NAME
    ClearRowsBelowHeader wsSql, COL_REPLACEMENT
    ClearOutputSheet wsOutput
    RestoreHeaders wsRef, wsSql
    If showMessage Then
        MsgBox ClearDoneMessage(), vbInformation
    End If
    RestoreFindSearchOrderByRows wsSql
End Sub

' 変換定義シートから修飾付きIDと単独IDの変換表を作成
Private Sub LoadMappings(ByVal wsRef As Worksheet, ByVal qualifiedMap As Object, ByVal standaloneMap As Object)
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim tableId As String
    Dim fieldId As String
    Dim fieldName As String

    lastRow = LastUsedRow(wsRef)
    For rowNumber = 2 To lastRow
        tableId = NormalizeKey(wsRef.Cells(rowNumber, COL_TABLE_ID).Value)
        fieldId = NormalizeKey(wsRef.Cells(rowNumber, COL_FIELD_ID).Value)
        fieldName = NormalizeName(wsRef.Cells(rowNumber, COL_FIELD_NAME).Value)

        If Len(fieldId) > 0 And IsUsableJapaneseName(fieldName) Then
            If tableId = "-" Then
                standaloneMap(fieldId) = fieldName
            ElseIf Len(tableId) > 0 Then
                qualifiedMap(tableId & "." & fieldId) = tableId & "." & fieldName
            End If
        End If
    Next rowNumber
End Sub

' 1行分のSQLへ変換表を適用し、変換後値を記録
Private Function ApplyMappings( _
    ByVal sourceText As String, _
    ByVal qualifiedMap As Object, _
    ByVal qualifiedKeys As Variant, _
    ByVal standaloneMap As Object, _
    ByVal standaloneKeys As Variant, _
    ByVal replacementValues As Object) As String
    Dim resultText As String

    resultText = sourceText

    resultText = ApplyMappingSet(resultText, qualifiedMap, qualifiedKeys, replacementValues, False)
    resultText = ApplyMappingSet(resultText, standaloneMap, standaloneKeys, replacementValues, True)

    ApplyMappings = resultText
End Function

' 指定された変換表を1行分のSQLへ適用
Private Function ApplyMappingSet( _
    ByVal sourceText As String, _
    ByVal mapping As Object, _
    ByVal sortedKeys As Variant, _
    ByVal replacementValues As Object, _
    ByVal standaloneMode As Boolean) As String
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
            resultText = ReplaceIdentifier(resultText, searchText, replacementText, changeCount, firstMatchIndex, standaloneMode)
            If changeCount > 0 Then
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
    ByVal standaloneMode As Boolean) As String
    Dim re As Object
    Dim matches As Object
    Dim matchItem As Object
    Dim resultText As String
    Dim index As Long
    Dim prefix As String
    Dim suffix As String
    Dim identifierStart As Long

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False
    ' 識別子の一部一致を避けるため、前後を英数字とアンダースコア以外に限定
    If standaloneMode Then
        re.Pattern = "(^|[^A-Za-z0-9_.])" & EscapeRegexLiteral(searchText) & "([^A-Za-z0-9_.]|$)"
    Else
        re.Pattern = "(^|[^A-Za-z0-9_])" & EscapeRegexLiteral(searchText) & "([^A-Za-z0-9_]|$)"
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

' 変換定義シートを取得
Private Function GetReferenceSheet() As Worksheet
    Set GetReferenceSheet = ResolveOrCreateSheet(ReferenceSheetName(), REF_LEGACY_SHEET, 1)
End Function

' SQL解析シートを取得
Private Function GetSqlSheet() As Worksheet
    Set GetSqlSheet = ResolveOrCreateSheet(SqlSheetName(), SQL_LEGACY_SHEET, 2)
End Function

' アウトプットシートを取得
Private Function GetOutputSheet() As Worksheet
    Dim ws As Worksheet

    Set ws = ResolveOrCreateSheet(OutputSheetName(), OutputSheetName(), 3)
    ApplyOutputSheetLayout ws
    Set GetOutputSheet = ws
End Function

' アウトプットシートの表示と書式を適用
Private Sub ApplyOutputSheetLayout(ByVal ws As Worksheet)
    ApplyOutputSheetFont ws
    ApplyOutputSheetView ws
End Sub

' アウトプットシートの既定フォントを設定
Private Sub ApplyOutputSheetFont(ByVal ws As Worksheet)
    With ws.Cells.Font
        .Name = OutputFontName()
        .Size = OutputFontSize()
    End With
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
    SetReplacementColumnsWrapText ws, False, 26
End Sub

' ヘッダー範囲の結合と内容を初期化
Private Sub ResetHeaderRange(ByVal headerRange As Range)
    headerRange.UnMerge
    headerRange.ClearContents
End Sub

' 変換内容列以降の折り返し設定を変更
Private Sub SetReplacementColumnsWrapText(ByVal ws As Worksheet, ByVal wrapEnabled As Boolean, ByVal lastColumn As Long)
    lastColumn = MaxLong(lastColumn, COL_REPLACEMENT)
    ws.Range(ws.Columns(COL_REPLACEMENT), ws.Columns(lastColumn)).WrapText = wrapEnabled
End Sub

' 変換定義シートとSQL解析シートのヘッダーを既定値に復元
Private Sub RestoreHeaders(ByVal wsRef As Worksheet, ByVal wsSql As Worksheet)
    ApplyReferenceHeader wsRef
    ApplySqlHeader wsSql
End Sub

' SQL解析シートへ解析ボタンとクリアボタンを配置
Private Sub InstallButtons(ByVal ws As Worksheet)
    Dim buttonTop As Double
    Dim buttonLeft As Double
    Dim analyzeButton As Object
    Dim clearButton As Object

    DeleteShapeIfExists ws, "btnAnalyzeQueries"
    DeleteShapeIfExists ws, "btnClearData"

    buttonTop = ws.Rows(1).Top + 2
    buttonLeft = ws.Columns("E").Left

    Set analyzeButton = ws.Buttons.Add(buttonLeft, buttonTop, 72, 24)
    With analyzeButton
        .Name = "btnAnalyzeQueries"
        .Caption = AnalyzeButtonText()
        .OnAction = "AnalyzeQueries"
    End With

    Set clearButton = ws.Buttons.Add(buttonLeft + 82, buttonTop, 72, 24)
    With clearButton
        .Name = "btnClearData"
        .Caption = ClearButtonText()
        .OnAction = "ClearData"
    End With
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
        wsSql.Cells(rowNumber, COL_REPLACEMENT + index).Value = CStr(keys(index))
    Next index
End Sub

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

' アウトプットシートの内容だけをクリア
Private Sub ClearOutputSheet(ByVal ws As Worksheet)
    ws.Cells.ClearContents
End Sub

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

' アウトプットシート名を取得
Private Function OutputSheetName() As String
    OutputSheetName = W(&H30A2, &H30A6, &H30C8, &H30D7, &H30C3, &H30C8)
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

' 和名変換後クエリ見出しを取得
Private Function ResultHeader() As String
    ResultHeader = W(&H548C, &H540D, &H5909, &H63DB, &H5F8C, &H30AF, &H30A8, &H30EA)
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

' 和名未取得判定用の文字列を取得
Private Function MissingNameText() As String
    MissingNameText = W(&H548C, &H540D, &H672A, &H53D6, &H5F97)
End Function

' 解析完了メッセージを取得
Private Function AnalyzeDoneMessage() As String
    AnalyzeDoneMessage = W(&H89E3, &H6790, &H304C, &H5B8C, &H4E86, &H3057, &H307E, &H3057, &H305F, &H3002)
End Function

' クリア完了メッセージを取得
Private Function ClearDoneMessage() As String
    ClearDoneMessage = W(&H30AF, &H30EA, &H30A2, &H304C, &H5B8C, &H4E86, &H3057, &H307E, &H3057, &H305F, &H3002)
End Function

' クリア確認メッセージを取得
Private Function ClearConfirmMessage() As String
    ClearConfirmMessage = W(&H0032, &H884C, &H76EE, &H4EE5, &H964D, &H3092, &H30AF, &H30EA, &H30A2, &H3057, &H307E, &H3059, &H3002, &H3088, &H308D, &H3057, &H3044, &H3067, &H3059, &H304B, &HFF1F)
End Function

' 確認ダイアログのタイトルを取得
Private Function ConfirmTitle() As String
    ConfirmTitle = W(&H78BA, &H8A8D)
End Function

' 変換定義なしメッセージを取得
Private Function NoDefinitionMessage() As String
    NoDefinitionMessage = W(&H5909, &H63DB, &H5B9A, &H7FA9, &H30B7, &H30FC, &H30C8, &H306B, &H6709, &H52B9, &H306A, &H5909, &H63DB, &H5B9A, &H7FA9, &H304C, &H3042, &H308A, &H307E, &H305B, &H3093, &H3002)
End Function
