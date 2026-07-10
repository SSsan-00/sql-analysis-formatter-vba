Attribute VB_Name = "SqlAnalysisFormatterTests"
Option Explicit

'@TestModule
'@Folder("Tests")

Private Const COL_SQL As Long = 1
Private Const COL_RESULT As Long = 2
Private Const COL_REPLACEMENT As Long = 3

' テスト一式をまとめて実行
Public Sub RunAllSqlAnalysisFormatterTests(Optional ByVal showMessage As Boolean = True)
    On Error GoTo TestFail

    SetupWorkbook_CreatesOutputSheet
    AnalyzeQueries_ConvertsCrudFixtures

    If showMessage Then
        MsgBox "SqlAnalysisFormatter tests passed.", vbInformation
    End If
    Exit Sub

TestFail:
    If showMessage Then
        MsgBox Err.Description, vbCritical
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'@TestMethod("SetupWorkbook")
Public Sub SetupWorkbook_CreatesOutputSheet()
    SetupWorkbook

    AssertWorksheetExists OutputSheetName()
End Sub

'@TestMethod("AnalyzeQueries")
Public Sub AnalyzeQueries_ConvertsCrudFixtures()
    Dim wsSql As Worksheet

    ArrangeCrudFixtures
    AnalyzeQueries False

    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    AssertAnalyzeRow wsSql, 2, ExpectedSelectSql(), _
        Array("users." & FullNameText(), "users." & UserIdText(), "orders." & OrderIdText(), StatusText(), "orders." & OrderUserIdText())
    AssertAnalyzeRow wsSql, 3, ExpectedInsertSql(), _
        Array(StatusText(), CreatedAtText(), "orders." & OrderIdText(), "users." & UserIdText(), "orders." & AmountText())
    AssertAnalyzeRow wsSql, 4, ExpectedUpdateSql(), _
        Array("users." & FullNameText(), UpdatedAtText(), StatusText(), "users." & UserIdText())
    AssertAnalyzeRow wsSql, 5, ExpectedDeleteSql(), _
        Array("orders." & OrderIdText(), "order_items." & DetailOrderIdText(), "order_items." & ProductIdText(), StatusText())
    AssertAnalyzeRow wsSql, 6, ExpectedComplexSelectSql(), _
        Array("users." & UserIdText(), "orders." & AmountText(), StatusText(), "orders." & OrderUserIdText(), "order_items." & DetailOrderIdText(), "orders." & OrderIdText(), "order_items." & QuantityText())
    AssertAnalyzeRow wsSql, 7, ExpectedSelfJoinSql(), _
        Array("users." & UserIdText(), "users." & FullNameText(), "manager." & FullNameText(), "users." & ManagerIdText(), "manager." & UserIdText(), "manager." & StatusText(), StatusText())
    AssertAnalyzeRow wsSql, 8, ExpectedSelectIntoSql(), _
        Array("users." & UserIdText(), "users." & MailText(), StatusText())
    AssertAnalyzeRow wsSql, 9, ExpectedUpdateFromSql(), _
        Array("orders." & AmountText(), UpdatedAtText(), "orders." & OrderUserIdText(), "users." & UserIdText(), "users." & MailText(), StatusText(), "order_items." & DetailOrderIdText(), "orders." & OrderIdText())
    AssertAnalyzeRow wsSql, 10, ExpectedDeleteExistsSql(), _
        Array("orders." & OrderIdText(), "order_items." & DetailOrderIdText(), StatusText(), "orders." & AmountText())
End Sub

' CRUDを含む解析テスト用データを作成
Private Sub ArrangeCrudFixtures()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents

    SeedReferenceDefinitions wsRef
    SeedCrudQueries wsSql
End Sub

' CRUDサンプルで使う変換定義を投入
Private Sub SeedReferenceDefinitions(ByVal ws As Worksheet)
    PutDefinition ws, 2, "users", UserTableText(), "user_id", UserIdText()
    PutDefinition ws, 3, "users", UserTableText(), "name", FullNameText()
    PutDefinition ws, 4, "users", UserTableText(), "email", MailText()
    PutDefinition ws, 5, "orders", OrderTableText(), "order_id", OrderIdText()
    PutDefinition ws, 6, "orders", OrderTableText(), "user_id", OrderUserIdText()
    PutDefinition ws, 7, "orders", OrderTableText(), "amount", AmountText()
    PutDefinition ws, 8, "order_items", OrderItemTableText(), "order_id", DetailOrderIdText()
    PutDefinition ws, 9, "order_items", OrderItemTableText(), "product_id", ProductIdText()
    PutDefinition ws, 10, "order_items", OrderItemTableText(), "quantity", QuantityText()
    PutDefinition ws, 11, "users", UserTableText(), "manager_id", ManagerIdText()
    PutDefinition ws, 12, "manager", ManagerTableText(), "user_id", UserIdText()
    PutDefinition ws, 13, "manager", ManagerTableText(), "name", FullNameText()
    PutDefinition ws, 14, "manager", ManagerTableText(), "status", StatusText()
    PutDefinition ws, 15, "-", "", "status", StatusText()
    PutDefinition ws, 16, "-", "", "created_at", CreatedAtText()
    PutDefinition ws, 17, "-", "", "updated_at", UpdatedAtText()
End Sub

' CRUDサンプルクエリを投入
Private Sub SeedCrudQueries(ByVal ws As Worksheet)
    ws.Cells(2, COL_SQL).Value = InputSelectSql()
    ws.Cells(3, COL_SQL).Value = InputInsertSql()
    ws.Cells(4, COL_SQL).Value = InputUpdateSql()
    ws.Cells(5, COL_SQL).Value = InputDeleteSql()
    ws.Cells(6, COL_SQL).Value = InputComplexSelectSql()
    ws.Cells(7, COL_SQL).Value = InputSelfJoinSql()
    ws.Cells(8, COL_SQL).Value = InputSelectIntoSql()
    ws.Cells(9, COL_SQL).Value = InputUpdateFromSql()
    ws.Cells(10, COL_SQL).Value = InputDeleteExistsSql()
End Sub

' A5M2で整形したSELECTの入力SQLを返す
Private Function InputSelectSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "    trim(users.name) as name"
    AppendA5M2Line resultText, "    , users.user_id"
    AppendA5M2Line resultText, "    , orders.order_id"
    AppendA5M2Line resultText, TS("    , status")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, TS("    users")
    AppendA5M2Line resultText, TS("    inner join orders")
    AppendA5M2Line resultText, TS("        on users.user_id = orders.user_id")
    AppendA5M2Line resultText, "where"
    AppendA5M2Line resultText, "    status = 'ACTIVE'"

    InputSelectSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したINSERTの入力SQLを返す
Private Function InputInsertSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, TS("insert")
    AppendA5M2Line resultText, TS("into orders(order_id, user_id, amount, status, created_at)")
    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "    orders.order_id"
    AppendA5M2Line resultText, "    , users.user_id"
    AppendA5M2Line resultText, "    , orders.amount"
    AppendA5M2Line resultText, "    , status"
    AppendA5M2Line resultText, TS("    , created_at")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, "    users"

    InputInsertSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したUPDATEの入力SQLを返す
Private Function InputUpdateSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, TS("update users")
    AppendA5M2Line resultText, "set"
    AppendA5M2Line resultText, "    users.name = 'Taro'"
    AppendA5M2Line resultText, "    , updated_at = CURRENT_TIMESTAMP"
    AppendA5M2Line resultText, TS("    , status = 'ACTIVE'")
    AppendA5M2Line resultText, "where"
    AppendA5M2Line resultText, "    users.user_id = :user_id"

    InputUpdateSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したDELETEの入力SQLを返す
Private Function InputDeleteSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, TS("delete")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, TS("    orders")
    AppendA5M2Line resultText, "where"
    AppendA5M2Line resultText, TS("    orders.order_id in (")
    AppendA5M2Line resultText, "        select"
    AppendA5M2Line resultText, TS("            order_items.order_id")
    AppendA5M2Line resultText, "        from"
    AppendA5M2Line resultText, TS("            order_items")
    AppendA5M2Line resultText, "        where"
    AppendA5M2Line resultText, "            order_items.product_id = :product_id"
    AppendA5M2Line resultText, TS("    )")
    AppendA5M2Line resultText, "    and status = 'CANCELLED'"

    InputDeleteSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形した複合SELECTの入力SQLを返す
Private Function InputComplexSelectSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "    users.user_id"
    AppendA5M2Line resultText, TS("    , case")
    AppendA5M2Line resultText, TS("        when sum(orders.amount) > 100000")
    AppendA5M2Line resultText, TS("            then 'VIP'")
    AppendA5M2Line resultText, TS("        when sum(orders.amount) between 50000 and 100000")
    AppendA5M2Line resultText, TS("            then 'STANDARD'")
    AppendA5M2Line resultText, TS("        else status")
    AppendA5M2Line resultText, TS("        end as rank_name")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, TS("    users")
    AppendA5M2Line resultText, TS("    left join orders")
    AppendA5M2Line resultText, TS("        on users.user_id = orders.user_id")
    AppendA5M2Line resultText, "where"
    AppendA5M2Line resultText, TS("    (")
    AppendA5M2Line resultText, TS("        (status = 'ACTIVE' and orders.amount > 0)")
    AppendA5M2Line resultText, TS("        or (")
    AppendA5M2Line resultText, TS("            status = 'PENDING'")
    AppendA5M2Line resultText, TS("            and exists (")
    AppendA5M2Line resultText, "                select"
    AppendA5M2Line resultText, TS("                    1")
    AppendA5M2Line resultText, "                from"
    AppendA5M2Line resultText, TS("                    order_items")
    AppendA5M2Line resultText, "                where"
    AppendA5M2Line resultText, TS("                    order_items.order_id = orders.order_id")
    AppendA5M2Line resultText, "                    and order_items.quantity > 1"
    AppendA5M2Line resultText, "            )"
    AppendA5M2Line resultText, "        )"
    AppendA5M2Line resultText, TS("    )")
    AppendA5M2Line resultText, "group by"
    AppendA5M2Line resultText, "    users.user_id"
    AppendA5M2Line resultText, TS("    , status")
    AppendA5M2Line resultText, "having"
    AppendA5M2Line resultText, TS("    count(orders.order_id) > 0")
    AppendA5M2Line resultText, "order by"
    AppendA5M2Line resultText, "    users.user_id"
    AppendA5M2Line resultText, "    , status"

    InputComplexSelectSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形した自己結合の入力SQLを返す
Private Function InputSelfJoinSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "    users.user_id"
    AppendA5M2Line resultText, "    , users.name"
    AppendA5M2Line resultText, TS("    , manager.name as manager_name")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, TS("    users")
    AppendA5M2Line resultText, TS("    inner join users manager")
    AppendA5M2Line resultText, TS("        on users.manager_id = manager.user_id")
    AppendA5M2Line resultText, "where"
    AppendA5M2Line resultText, TS("    manager.status = status")
    AppendA5M2Line resultText, "order by"
    AppendA5M2Line resultText, "    manager.name"

    InputSelfJoinSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したSELECT-INTOの入力SQLを返す
Private Function InputSelectIntoSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "    users.user_id"
    AppendA5M2Line resultText, "    , users.email"
    AppendA5M2Line resultText, TS("    , status")
    AppendA5M2Line resultText, TS("into user_export")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, TS("    users")
    AppendA5M2Line resultText, "where"
    AppendA5M2Line resultText, TS("    users.email is not null")
    AppendA5M2Line resultText, TS("    and status in ('ACTIVE', 'LOCKED')")
    AppendA5M2Line resultText, "order by"
    AppendA5M2Line resultText, "    users.email"

    InputSelectIntoSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したUPDATE-FROMの入力SQLを返す
Private Function InputUpdateFromSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, TS("update orders")
    AppendA5M2Line resultText, "set"
    AppendA5M2Line resultText, "    orders.amount = orders.amount * 1.1"
    AppendA5M2Line resultText, TS("    , updated_at = CURRENT_TIMESTAMP")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, TS("    orders")
    AppendA5M2Line resultText, TS("    inner join users")
    AppendA5M2Line resultText, TS("        on orders.user_id = users.user_id")
    AppendA5M2Line resultText, "where"
    AppendA5M2Line resultText, TS("    (users.email like :domain or status = 'PENDING')")
    AppendA5M2Line resultText, TS("    and (")
    AppendA5M2Line resultText, TS("        orders.amount > 1000")
    AppendA5M2Line resultText, TS("        or exists (")
    AppendA5M2Line resultText, "            select"
    AppendA5M2Line resultText, TS("                1")
    AppendA5M2Line resultText, "            from"
    AppendA5M2Line resultText, TS("                order_items")
    AppendA5M2Line resultText, "            where"
    AppendA5M2Line resultText, "                order_items.order_id = orders.order_id"
    AppendA5M2Line resultText, "        )"
    AppendA5M2Line resultText, "    )"

    InputUpdateFromSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したDELETE EXISTSの入力SQLを返す
Private Function InputDeleteExistsSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, TS("delete")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, TS("    order_items")
    AppendA5M2Line resultText, "where"
    AppendA5M2Line resultText, TS("    exists (")
    AppendA5M2Line resultText, "        select"
    AppendA5M2Line resultText, TS("            1")
    AppendA5M2Line resultText, "        from"
    AppendA5M2Line resultText, TS("            orders")
    AppendA5M2Line resultText, "        where"
    AppendA5M2Line resultText, TS("            orders.order_id = order_items.order_id")
    AppendA5M2Line resultText, "            and (status = 'CANCELLED' or orders.amount <= 0)"
    AppendA5M2Line resultText, "    )"

    InputDeleteExistsSql = FinishA5M2Sql(resultText)
End Function

' SELECTの和名変換後期待値を返す
Private Function ExpectedSelectSql() As String
    ExpectedSelectSql = ConvertFixtureSql(InputSelectSql())
End Function

' INSERTの和名変換後期待値を返す
Private Function ExpectedInsertSql() As String
    ExpectedInsertSql = ConvertFixtureSql(InputInsertSql())
End Function

' UPDATEの和名変換後期待値を返す
Private Function ExpectedUpdateSql() As String
    ExpectedUpdateSql = ConvertFixtureSql(InputUpdateSql())
End Function

' DELETEの和名変換後期待値を返す
Private Function ExpectedDeleteSql() As String
    ExpectedDeleteSql = ConvertFixtureSql(InputDeleteSql())
End Function

' 複合SELECTの和名変換後期待値を返す
Private Function ExpectedComplexSelectSql() As String
    ExpectedComplexSelectSql = ConvertFixtureSql(InputComplexSelectSql())
End Function

' 自己結合の和名変換後期待値を返す
Private Function ExpectedSelfJoinSql() As String
    ExpectedSelfJoinSql = ConvertFixtureSql(InputSelfJoinSql())
End Function

' SELECT-INTOの和名変換後期待値を返す
Private Function ExpectedSelectIntoSql() As String
    ExpectedSelectIntoSql = ConvertFixtureSql(InputSelectIntoSql())
End Function

' UPDATE-FROMの和名変換後期待値を返す
Private Function ExpectedUpdateFromSql() As String
    ExpectedUpdateFromSql = ConvertFixtureSql(InputUpdateFromSql())
End Function

' DELETE EXISTSの和名変換後期待値を返す
Private Function ExpectedDeleteExistsSql() As String
    ExpectedDeleteExistsSql = ConvertFixtureSql(InputDeleteExistsSql())
End Function

' A5M2のCRLFをExcelセル内改行に合わせてLFで保持
Private Sub AppendA5M2Line(ByRef resultText As String, ByVal lineText As String)
    If Len(resultText) > 0 Then
        resultText = resultText & vbLf
    End If
    resultText = resultText & lineText
End Sub

' A5M2出力末尾の改行を付与
Private Function FinishA5M2Sql(ByVal resultText As String) As String
    FinishA5M2Sql = resultText & vbLf
End Function

' A5M2が付ける行末スペースを明示
Private Function TS(ByVal lineText As String) As String
    TS = lineText & " "
End Function

' 入力SQLから和名変換後の期待値を作成
Private Function ConvertFixtureSql(ByVal sourceText As String) As String
    Dim resultText As String

    resultText = sourceText
    resultText = Replace(resultText, "order_items.product_id", "order_items." & ProductIdText(), , , vbBinaryCompare)
    resultText = Replace(resultText, "order_items.quantity", "order_items." & QuantityText(), , , vbBinaryCompare)
    resultText = Replace(resultText, "order_items.order_id", "order_items." & DetailOrderIdText(), , , vbBinaryCompare)
    resultText = Replace(resultText, "users.manager_id", "users." & ManagerIdText(), , , vbBinaryCompare)
    resultText = Replace(resultText, "manager.user_id", "manager." & UserIdText(), , , vbBinaryCompare)
    resultText = Replace(resultText, "manager.status", "manager." & StatusText(), , , vbBinaryCompare)
    resultText = Replace(resultText, "manager.name", "manager." & FullNameText(), , , vbBinaryCompare)
    resultText = Replace(resultText, "orders.order_id", "orders." & OrderIdText(), , , vbBinaryCompare)
    resultText = Replace(resultText, "orders.user_id", "orders." & OrderUserIdText(), , , vbBinaryCompare)
    resultText = Replace(resultText, "orders.amount", "orders." & AmountText(), , , vbBinaryCompare)
    resultText = Replace(resultText, "users.user_id", "users." & UserIdText(), , , vbBinaryCompare)
    resultText = Replace(resultText, "users.email", "users." & MailText(), , , vbBinaryCompare)
    resultText = Replace(resultText, "users.name", "users." & FullNameText(), , , vbBinaryCompare)
    resultText = ReplaceStandaloneFixture(resultText, "created_at", CreatedAtText())
    resultText = ReplaceStandaloneFixture(resultText, "updated_at", UpdatedAtText())
    resultText = ReplaceStandaloneFixture(resultText, "status", StatusText())

    ConvertFixtureSql = resultText
End Function

' 単体フィールド定義を識別子単位で置換
Private Function ReplaceStandaloneFixture(ByVal sourceText As String, ByVal searchText As String, ByVal replacementText As String) As String
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
    re.Pattern = "(^|[^A-Za-z0-9_.])" & searchText & "([^A-Za-z0-9_.]|$)"

    Set matches = re.Execute(sourceText)
    resultText = sourceText
    For index = matches.Count - 1 To 0 Step -1
        Set matchItem = matches.Item(index)
        prefix = CStr(matchItem.SubMatches(0))
        suffix = CStr(matchItem.SubMatches(1))
        resultText = Left$(resultText, matchItem.FirstIndex) _
            & prefix & replacementText & suffix _
            & Mid$(resultText, matchItem.FirstIndex + matchItem.Length + 1)
    Next index

    ReplaceStandaloneFixture = resultText
End Function

Private Sub PutDefinition(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal tableId As String, ByVal tableName As String, ByVal fieldId As String, ByVal fieldName As String)
    ws.Cells(rowNumber, 1).Value = tableId
    ws.Cells(rowNumber, 2).Value = tableName
    ws.Cells(rowNumber, 3).Value = fieldId
    ws.Cells(rowNumber, 4).Value = fieldName
End Sub

Private Sub AssertAnalyzeRow(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal expectedSql As String, ByVal expectedReplacements As Variant)
    Dim index As Long
    Dim replacementCount As Long

    AssertCellValue ws.Cells(rowNumber, COL_RESULT), expectedSql
    For index = LBound(expectedReplacements) To UBound(expectedReplacements)
        AssertCellValue ws.Cells(rowNumber, COL_REPLACEMENT + index), CStr(expectedReplacements(index))
    Next index

    replacementCount = UBound(expectedReplacements) - LBound(expectedReplacements) + 1
    AssertCellValue ws.Cells(rowNumber, COL_REPLACEMENT + replacementCount), ""
End Sub

Private Sub AssertWorksheetExists(ByVal sheetName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Fail "Worksheet not found: " & sheetName
    End If
End Sub

Private Sub AssertCellValue(ByVal cell As Range, ByVal expected As String)
    Dim actual As String

    actual = CStr(cell.Value)
    If actual <> expected Then
        Fail cell.Worksheet.Name & "!" & cell.Address(False, False) & _
            " expected=[" & expected & "] actual=[" & actual & "]"
    End If
End Sub

Private Sub Fail(ByVal message As String)
    Err.Raise vbObjectError + 513, "SqlAnalysisFormatterTests", message
End Sub

' VBEインポート時の文字化けを避けるため、テスト文字列もコードポイントで保持
Private Function W(ParamArray codes() As Variant) As String
    Dim index As Long
    Dim resultText As String

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

Private Function OutputSheetName() As String
    OutputSheetName = W(&H30A2, &H30A6, &H30C8, &H30D7, &H30C3, &H30C8)
End Function

Private Function UserTableText() As String
    UserTableText = W(&H30E6, &H30FC, &H30B6, &H30FC)
End Function

Private Function OrderTableText() As String
    OrderTableText = W(&H6CE8, &H6587)
End Function

Private Function OrderItemTableText() As String
    OrderItemTableText = W(&H6CE8, &H6587, &H660E, &H7D30)
End Function

Private Function ManagerTableText() As String
    ManagerTableText = W(&H7BA1, &H7406, &H8005)
End Function

Private Function UserIdText() As String
    UserIdText = UserTableText() & "ID"
End Function

Private Function FullNameText() As String
    FullNameText = W(&H6C0F, &H540D)
End Function

Private Function MailText() As String
    MailText = W(&H30E1, &H30FC, &H30EB)
End Function

Private Function ManagerIdText() As String
    ManagerIdText = ManagerTableText() & "ID"
End Function

Private Function OrderIdText() As String
    OrderIdText = OrderTableText() & "ID"
End Function

Private Function OrderUserIdText() As String
    OrderUserIdText = OrderTableText() & UserTableText() & "ID"
End Function

Private Function AmountText() As String
    AmountText = W(&H91D1, &H984D)
End Function

Private Function DetailOrderIdText() As String
    DetailOrderIdText = W(&H660E, &H7D30, &H6CE8, &H6587) & "ID"
End Function

Private Function ProductIdText() As String
    ProductIdText = W(&H5546, &H54C1) & "ID"
End Function

Private Function QuantityText() As String
    QuantityText = W(&H6570, &H91CF)
End Function

Private Function StatusText() As String
    StatusText = W(&H72B6, &H614B)
End Function

Private Function CreatedAtText() As String
    CreatedAtText = W(&H4F5C, &H6210, &H65E5, &H6642)
End Function

Private Function UpdatedAtText() As String
    UpdatedAtText = W(&H66F4, &H65B0, &H65E5, &H6642)
End Function
