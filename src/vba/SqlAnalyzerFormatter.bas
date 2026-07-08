Attribute VB_Name = "SqlAnalyzerFormatter"
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

Public Sub SetupWorkbook()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet

    Set wsRef = ResolveOrCreateSheet(ReferenceSheetName(), REF_LEGACY_SHEET, 1)
    Set wsSql = ResolveOrCreateSheet(SqlSheetName(), SQL_LEGACY_SHEET, 2)

    ApplyReferenceHeader wsRef
    ApplySqlHeader wsSql
    InstallButtons wsSql
End Sub

Public Sub AnalyzeQueries(Optional ByVal showMessage As Boolean = True)
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim qualifiedMap As Object
    Dim tokenMap As Object
    Dim lastRow As Long
    Dim clearLastRow As Long
    Dim rowNumber As Long
    Dim sourceText As String
    Dim convertedText As String
    Dim details As String

    Set wsRef = GetReferenceSheet()
    Set wsSql = GetSqlSheet()
    Set qualifiedMap = CreateObject("Scripting.Dictionary")
    Set tokenMap = CreateObject("Scripting.Dictionary")
    qualifiedMap.CompareMode = vbBinaryCompare
    tokenMap.CompareMode = vbBinaryCompare

    LoadMappings wsRef, qualifiedMap, tokenMap
    If qualifiedMap.Count = 0 And tokenMap.Count = 0 Then
        If showMessage Then
            MsgBox NoDefinitionMessage(), vbInformation
        End If
        Exit Sub
    End If

    lastRow = LastUsedRowInColumn(wsSql, COL_SQL)
    ' 前回の出力だけを消し、入力SQLは残す
    clearLastRow = MaxLong(lastRow, LastUsedRowInColumn(wsSql, COL_REPLACEMENT))
    If clearLastRow >= 2 Then
        wsSql.Range(wsSql.Cells(2, COL_RESULT), wsSql.Cells(clearLastRow, COL_REPLACEMENT)).ClearContents
    End If

    For rowNumber = 2 To lastRow
        sourceText = CStr(wsSql.Cells(rowNumber, COL_SQL).Value)
        If Len(sourceText) > 0 Then
            details = vbNullString
            convertedText = ApplyMappings(sourceText, qualifiedMap, tokenMap, details)
            wsSql.Cells(rowNumber, COL_RESULT).Value = convertedText
            If Len(details) > 0 Then
                wsSql.Cells(rowNumber, COL_REPLACEMENT).Value = details
            End If
        End If
    Next rowNumber

    wsSql.Columns(COL_RESULT).WrapText = False
    wsSql.Columns(COL_REPLACEMENT).WrapText = True
    If showMessage Then
        MsgBox AnalyzeDoneMessage(), vbInformation
    End If
End Sub

Public Sub ClearData(Optional ByVal showMessage As Boolean = True)
    ClearRowsBelowHeader GetReferenceSheet(), COL_FIELD_NAME
    ClearRowsBelowHeader GetSqlSheet(), COL_REPLACEMENT
    If showMessage Then
        MsgBox ClearDoneMessage(), vbInformation
    End If
End Sub

Private Sub LoadMappings(ByVal wsRef As Worksheet, ByVal qualifiedMap As Object, ByVal tokenMap As Object)
    Dim uniqueTokens As Object
    Dim conflictTokens As Object
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim tableId As String
    Dim tableName As String
    Dim fieldId As String
    Dim fieldName As String
    Dim key As Variant

    Set uniqueTokens = CreateObject("Scripting.Dictionary")
    Set conflictTokens = CreateObject("Scripting.Dictionary")
    uniqueTokens.CompareMode = vbBinaryCompare
    conflictTokens.CompareMode = vbBinaryCompare

    ' 単独フィールドIDは和名が一意に決まる場合だけ採用
    lastRow = LastUsedRow(wsRef)
    For rowNumber = 2 To lastRow
        tableId = NormalizeKey(wsRef.Cells(rowNumber, COL_TABLE_ID).Value)
        tableName = NormalizeName(wsRef.Cells(rowNumber, COL_TABLE_NAME).Value)
        fieldId = NormalizeKey(wsRef.Cells(rowNumber, COL_FIELD_ID).Value)
        fieldName = NormalizeName(wsRef.Cells(rowNumber, COL_FIELD_NAME).Value)

        If Len(fieldId) > 0 And IsUsableJapaneseName(fieldName) Then
            If Len(tableId) > 0 And tableId <> "-" And IsUsableJapaneseName(tableName) Then
                qualifiedMap(tableId & "." & fieldId) = tableName & "." & fieldName
            End If

            If Not conflictTokens.Exists(fieldId) Then
                If uniqueTokens.Exists(fieldId) Then
                    If CStr(uniqueTokens(fieldId)) <> fieldName Then
                        conflictTokens(fieldId) = True
                    End If
                Else
                    uniqueTokens(fieldId) = fieldName
                End If
            End If
        End If
    Next rowNumber

    For Each key In uniqueTokens.Keys
        If Not conflictTokens.Exists(CStr(key)) Then
            tokenMap(CStr(key)) = CStr(uniqueTokens(key))
        End If
    Next key
End Sub

Private Function ApplyMappings(ByVal sourceText As String, ByVal qualifiedMap As Object, ByVal tokenMap As Object, ByRef details As String) As String
    Dim resultText As String
    Dim notes As Object
    Dim key As Variant
    Dim changeCount As Long
    Dim replacementText As String

    resultText = sourceText
    Set notes = CreateObject("Scripting.Dictionary")
    notes.CompareMode = vbBinaryCompare

    ' 修飾付き識別子を先に処理し、単独IDとの重複置換を避ける
    For Each key In SortedKeysByLengthDesc(qualifiedMap)
        replacementText = CStr(qualifiedMap(CStr(key)))
        changeCount = 0
        resultText = ReplaceIdentifier(resultText, CStr(key), replacementText, changeCount)
        If changeCount > 0 Then
            AddNote notes, CStr(key) & " -> " & replacementText
        End If
    Next key

    For Each key In SortedKeysByLengthDesc(tokenMap)
        replacementText = CStr(tokenMap(CStr(key)))
        changeCount = 0
        resultText = ReplaceIdentifier(resultText, CStr(key), replacementText, changeCount)
        If changeCount > 0 Then
            AddNote notes, CStr(key) & " -> " & replacementText
        End If
    Next key

    If notes.Count > 0 Then
        details = Join(notes.Keys, "; ")
    Else
        details = vbNullString
    End If
    ApplyMappings = resultText
End Function

Private Function ReplaceIdentifier(ByVal sourceText As String, ByVal searchText As String, ByVal replacementText As String, ByRef changeCount As Long) As String
    Dim re As Object
    Dim matches As Object
    Dim matchItem As Object
    Dim resultText As String
    Dim index As Long
    Dim prefix As String
    Dim suffix As String

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False
    ' 識別子の一部一致を避けるため、前後を英数字とアンダースコア以外に限定
    re.Pattern = "(^|[^A-Za-z0-9_])" & EscapeRegexLiteral(searchText) & "([^A-Za-z0-9_]|$)"

    Set matches = re.Execute(sourceText)
    changeCount = matches.Count
    resultText = sourceText

    ' 後方から置換し、FirstIndexのずれを防ぐ
    For index = matches.Count - 1 To 0 Step -1
        Set matchItem = matches.Item(index)
        prefix = CStr(matchItem.SubMatches(0))
        suffix = CStr(matchItem.SubMatches(1))
        resultText = Left$(resultText, matchItem.FirstIndex) _
            & prefix & replacementText & suffix _
            & Mid$(resultText, matchItem.FirstIndex + matchItem.Length + 1)
    Next index

    ReplaceIdentifier = resultText
End Function

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

Private Sub AddNote(ByVal notes As Object, ByVal noteText As String)
    If Not notes.Exists(noteText) Then
        notes(noteText) = True
    End If
End Sub

Private Function NormalizeKey(ByVal value As Variant) As String
    NormalizeKey = Trim$(CStr(value))
End Function

Private Function NormalizeName(ByVal value As Variant) As String
    NormalizeName = Trim$(CStr(value))
End Function

Private Function IsUsableJapaneseName(ByVal value As String) As Boolean
    Dim normalized As String

    normalized = Trim$(value)
    If Len(normalized) = 0 Then Exit Function
    If normalized = "-" Then Exit Function
    If InStr(1, normalized, MissingNameText(), vbBinaryCompare) > 0 Then Exit Function

    IsUsableJapaneseName = True
End Function

Private Function GetReferenceSheet() As Worksheet
    Set GetReferenceSheet = ResolveOrCreateSheet(ReferenceSheetName(), REF_LEGACY_SHEET, 1)
End Function

Private Function GetSqlSheet() As Worksheet
    Set GetSqlSheet = ResolveOrCreateSheet(SqlSheetName(), SQL_LEGACY_SHEET, 2)
End Function

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

Private Function TryGetWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set TryGetWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Sub ApplyReferenceHeader(ByVal ws As Worksheet)
    ws.Cells(1, COL_TABLE_ID).Value = TableIdHeader()
    ws.Cells(1, COL_TABLE_NAME).Value = TableNameHeader()
    ws.Cells(1, COL_FIELD_ID).Value = FieldIdHeader()
    ws.Cells(1, COL_FIELD_NAME).Value = FieldNameHeader()
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:D").AutoFit
End Sub

Private Sub ApplySqlHeader(ByVal ws As Worksheet)
    ws.Cells(1, COL_SQL).Value = SqlHeader()
    ws.Cells(1, COL_RESULT).Value = ResultHeader()
    ws.Cells(1, COL_REPLACEMENT).Value = ReplacementHeader()
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:C").ColumnWidth = 42
    ws.Columns("C").ColumnWidth = 50
End Sub

Private Sub InstallButtons(ByVal ws As Worksheet)
    Dim buttonTop As Double
    Dim buttonLeft As Double
    Dim analyzeButton As Object
    Dim clearButton As Object

    DeleteShapeIfExists ws, "btnAnalyzeQueries"
    DeleteShapeIfExists ws, "btnClearData"

    ws.Columns("E:F").ColumnWidth = 12
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

Private Sub DeleteShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
End Sub

Private Sub ClearRowsBelowHeader(ByVal ws As Worksheet, ByVal minimumLastColumn As Long)
    Dim lastRow As Long
    Dim lastColumn As Long

    lastRow = LastUsedRow(ws)
    lastColumn = MaxLong(LastUsedColumn(ws), minimumLastColumn)
    If lastRow >= 2 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastColumn)).ClearContents
    End If
End Sub

Private Function LastUsedRowInColumn(ByVal ws As Worksheet, ByVal columnNumber As Long) As Long
    Dim rowNumber As Long

    rowNumber = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row
    If rowNumber = 1 And Len(CStr(ws.Cells(1, columnNumber).Value)) = 0 Then
        LastUsedRowInColumn = 1
    Else
        LastUsedRowInColumn = rowNumber
    End If
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim foundCell As Range

    Set foundCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If foundCell Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = foundCell.Row
    End If
End Function

Private Function LastUsedColumn(ByVal ws As Worksheet) As Long
    Dim foundCell As Range

    Set foundCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If foundCell Is Nothing Then
        LastUsedColumn = 1
    Else
        LastUsedColumn = foundCell.Column
    End If
End Function

Private Function MaxLong(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    If leftValue >= rightValue Then
        MaxLong = leftValue
    Else
        MaxLong = rightValue
    End If
End Function

Private Function W(ParamArray codes() As Variant) As String
    Dim index As Long
    Dim resultText As String

    ' VBEインポート時の文字化けを避けるため、UI文字列はコードポイントで保持
    For index = LBound(codes) To UBound(codes)
        resultText = resultText & ChrW$(CLng(codes(index)))
    Next index

    W = resultText
End Function

Private Function ReferenceSheetName() As String
    ReferenceSheetName = W(&H5909, &H63DB, &H5B9A, &H7FA9)
End Function

Private Function SqlSheetName() As String
    SqlSheetName = "SQL" & W(&H89E3, &H6790)
End Function

Private Function TableIdHeader() As String
    TableIdHeader = W(&H6240, &H5C5E, &H30C6, &H30FC, &H30D6, &H30EB) & "ID"
End Function

Private Function TableNameHeader() As String
    TableNameHeader = W(&H6240, &H5C5E, &H30C6, &H30FC, &H30D6, &H30EB, &H548C, &H540D)
End Function

Private Function FieldIdHeader() As String
    FieldIdHeader = W(&H30D5, &H30A3, &H30FC, &H30EB, &H30C9) & "ID"
End Function

Private Function FieldNameHeader() As String
    FieldNameHeader = W(&H30D5, &H30A3, &H30FC, &H30EB, &H30C9, &H548C, &H540D)
End Function

Private Function SqlHeader() As String
    SqlHeader = "SQL" & W(&H30AF, &H30A8, &H30EA)
End Function

Private Function ResultHeader() As String
    ResultHeader = W(&H548C, &H540D, &H5909, &H63DB, &H5F8C, &H30AF, &H30A8, &H30EA)
End Function

Private Function ReplacementHeader() As String
    ReplacementHeader = W(&H5909, &H63DB, &H5185, &H5BB9)
End Function

Private Function AnalyzeButtonText() As String
    AnalyzeButtonText = W(&H89E3, &H6790)
End Function

Private Function ClearButtonText() As String
    ClearButtonText = W(&H30AF, &H30EA, &H30A2)
End Function

Private Function MissingNameText() As String
    MissingNameText = W(&H548C, &H540D, &H672A, &H53D6, &H5F97)
End Function

Private Function AnalyzeDoneMessage() As String
    AnalyzeDoneMessage = W(&H89E3, &H6790, &H304C, &H5B8C, &H4E86, &H3057, &H307E, &H3057, &H305F, &H3002)
End Function

Private Function ClearDoneMessage() As String
    ClearDoneMessage = W(&H30AF, &H30EA, &H30A2, &H304C, &H5B8C, &H4E86, &H3057, &H307E, &H3057, &H305F, &H3002)
End Function

Private Function NoDefinitionMessage() As String
    NoDefinitionMessage = W(&H5909, &H63DB, &H5B9A, &H7FA9, &H30B7, &H30FC, &H30C8, &H306B, &H6709, &H52B9, &H306A, &H5909, &H63DB, &H5B9A, &H7FA9, &H304C, &H3042, &H308A, &H307E, &H305B, &H3093, &H3002)
End Function
