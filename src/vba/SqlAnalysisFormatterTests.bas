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
    AnalyzeQueries_ConvertsTsqlFunctionFixtures
    ClearData_ClearsOutputSheet

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
    AssertOutputSheetGridlinesHidden
    AssertOutputSheetFont
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

'@TestMethod("AnalyzeQueries")
Public Sub AnalyzeQueries_ConvertsTsqlFunctionFixtures()
    Dim wsSql As Worksheet

    ArrangeTsqlFunctionFixtures
    AnalyzeQueries False

    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    AssertAnalyzeRow wsSql, 2, ExpectedTsqlTrimFromSql(), _
        Array("users." & FullNameText(), "users." & UserIdText())
    AssertAnalyzeRow wsSql, 3, ExpectedTsqlInSql(), _
        Array("users." & UserIdText(), StatusText(), "orders." & OrderUserIdText(), "orders." & AmountText())
    AssertAnalyzeRow wsSql, 4, ExpectedTsqlCoalesceSql(), _
        Array("users." & UserIdText(), "users." & MailText(), "users." & FullNameText())
    AssertAnalyzeRow wsSql, 5, ExpectedTsqlFormatSql(), _
        Array("orders." & OrderIdText(), "orders." & AmountText(), CreatedAtText())
    AssertAnalyzeRow wsSql, 6, ExpectedTsqlWithSql(), _
        Array("users." & UserIdText(), StatusText())
    AssertAnalyzeRow wsSql, 7, ExpectedTsqlCastSql(), _
        Array("users." & UserIdText(), "orders." & AmountText(), CreatedAtText(), UpdatedAtText(), StatusText(), "orders." & OrderUserIdText())
    AssertAnalyzeRow wsSql, 8, ExpectedTsqlIsNullSql(), _
        Array("users." & UserIdText(), "users." & MailText(), StatusText())
    AssertAnalyzeRow wsSql, 9, ExpectedTsqlSubstringSql(), _
        Array("users." & UserIdText(), "users." & MailText())
    AssertAnalyzeRow wsSql, 10, ExpectedTsqlRoundSql(), _
        Array("orders." & OrderIdText(), "orders." & AmountText())
    AssertAnalyzeRow wsSql, 11, ExpectedTsqlSumSql(), _
        Array("orders." & OrderUserIdText(), "orders." & AmountText())
    AssertAnalyzeRow wsSql, 12, ExpectedTsqlReplaceSql(), _
        Array("users." & UserIdText(), "users." & MailText())
    AssertAnalyzeRow wsSql, 13, ExpectedTsqlDateAddSql(), _
        Array("orders." & OrderIdText(), CreatedAtText())
    AssertAnalyzeRow wsSql, 14, ExpectedTsqlDateDiffSql(), _
        Array("orders." & OrderIdText(), CreatedAtText(), UpdatedAtText())
    AssertAnalyzeRow wsSql, 15, ExpectedTsqlCountSql(), _
        Array("users." & UserIdText(), "orders." & OrderIdText(), "orders." & OrderUserIdText())
    AssertAnalyzeRow wsSql, 16, ExpectedTsqlExistsSql(), _
        Array("users." & UserIdText(), "orders." & OrderUserIdText(), "orders." & AmountText())
End Sub

'@TestMethod("ClearData")
Public Sub ClearData_ClearsOutputSheet()
    Dim wsOutput As Worksheet

    SetupWorkbook
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())
    wsOutput.Cells(1, 1).Value = "output header"
    wsOutput.Cells(3, 2).Value = "output detail"

    ClearData False

    AssertCellValue wsOutput.Cells(1, 1), ""
    AssertCellValue wsOutput.Cells(3, 2), ""
    AssertOutputSheetFont
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

' T-SQL関数サンプルの解析テスト用データを作成
Private Sub ArrangeTsqlFunctionFixtures()
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
    SeedTsqlFunctionQueries wsSql
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

' T-SQLの主要関数・構文を独立した行へ投入
Private Sub SeedTsqlFunctionQueries(ByVal ws As Worksheet)
    ws.Cells(2, COL_SQL).Value = InputTsqlTrimFromSql()
    ws.Cells(3, COL_SQL).Value = InputTsqlInSql()
    ws.Cells(4, COL_SQL).Value = InputTsqlCoalesceSql()
    ws.Cells(5, COL_SQL).Value = InputTsqlFormatSql()
    ws.Cells(6, COL_SQL).Value = InputTsqlWithSql()
    ws.Cells(7, COL_SQL).Value = InputTsqlCastSql()
    ws.Cells(8, COL_SQL).Value = InputTsqlIsNullSql()
    ws.Cells(9, COL_SQL).Value = InputTsqlSubstringSql()
    ws.Cells(10, COL_SQL).Value = InputTsqlRoundSql()
    ws.Cells(11, COL_SQL).Value = InputTsqlSumSql()
    ws.Cells(12, COL_SQL).Value = InputTsqlReplaceSql()
    ws.Cells(13, COL_SQL).Value = InputTsqlDateAddSql()
    ws.Cells(14, COL_SQL).Value = InputTsqlDateDiffSql()
    ws.Cells(15, COL_SQL).Value = InputTsqlCountSql()
    ws.Cells(16, COL_SQL).Value = InputTsqlExistsSql()
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

' A5M2で整形したT-SQL TRIM FROMの入力SQLを返す
Private Function InputTsqlTrimFromSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, TS("    trim('.' from users.name) as trimmed_name")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, TS("    users")
    AppendA5M2Line resultText, "where"
    AppendA5M2Line resultText, "    users.user_id = @user_id"

    InputTsqlTrimFromSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したT-SQL INの入力SQLを返す
Private Function InputTsqlInSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "    users.user_id"
    AppendA5M2Line resultText, TS("    , status")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, TS("    users")
    AppendA5M2Line resultText, "where"
    AppendA5M2Line resultText, TS("    status in ('ACTIVE', 'LOCKED', 'PENDING')")
    AppendA5M2Line resultText, TS("    and users.user_id in (")
    AppendA5M2Line resultText, "        select"
    AppendA5M2Line resultText, TS("            orders.user_id")
    AppendA5M2Line resultText, "        from"
    AppendA5M2Line resultText, TS("            orders")
    AppendA5M2Line resultText, "        where"
    AppendA5M2Line resultText, "            orders.amount > 0"
    AppendA5M2Line resultText, "    )"

    InputTsqlInSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したT-SQL COALESCEの入力SQLを返す
Private Function InputTsqlCoalesceSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "    users.user_id"
    AppendA5M2Line resultText, TS("    , coalesce(users.email, users.name, 'unknown') as contact_text")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, "    users"

    InputTsqlCoalesceSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したT-SQL FORMATの入力SQLを返す
Private Function InputTsqlFormatSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "    orders.order_id"
    AppendA5M2Line resultText, "    , format(orders.amount, 'N2', 'ja-JP') as amount_n2"
    AppendA5M2Line resultText, "    , format(created_at, 'yyyy/MM/dd') as created_date"
    AppendA5M2Line resultText, "    , format(created_at, 'yyyyMMddHHmmss') as created_stamp"
    AppendA5M2Line resultText, TS("    , format(orders.amount, 'C', 'ja-JP') as amount_currency")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, "    orders"

    InputTsqlFormatSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したT-SQL WITHの入力SQLを返す
Private Function InputTsqlWithSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, TS("with target_users as (")
    AppendA5M2Line resultText, "    select"
    AppendA5M2Line resultText, TS("        users.user_id")
    AppendA5M2Line resultText, "    from"
    AppendA5M2Line resultText, TS("        users")
    AppendA5M2Line resultText, "    where"
    AppendA5M2Line resultText, "        status = 'ACTIVE'"
    AppendA5M2Line resultText, TS(")")
    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, TS("    target_users.user_id")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, "    target_users"

    InputTsqlWithSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したT-SQL CASTの入力SQLを返す
Private Function InputTsqlCastSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "    cast(users.user_id as int) as user_id_int"
    AppendA5M2Line resultText, "    , cast(orders.amount as decimal (18, 2)) as amount_decimal"
    AppendA5M2Line resultText, "    , cast(created_at as date) as created_date"
    AppendA5M2Line resultText, "    , cast(updated_at as datetime2(3)) as updated_at_dt"
    AppendA5M2Line resultText, TS("    , cast(status as nvarchar(20)) as status_text")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, TS("    users")
    AppendA5M2Line resultText, TS("    inner join orders")
    AppendA5M2Line resultText, "        on users.user_id = orders.user_id"

    InputTsqlCastSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したT-SQL ISNULLの入力SQLを返す
Private Function InputTsqlIsNullSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "    users.user_id"
    AppendA5M2Line resultText, "    , isnull(users.email, 'unknown') as email_text"
    AppendA5M2Line resultText, TS("    , isnull(status, 'UNKNOWN') as status_text")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, "    users"

    InputTsqlIsNullSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したT-SQL SUBSTRINGの入力SQLを返す
Private Function InputTsqlSubstringSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "    users.user_id"
    AppendA5M2Line resultText, TS("    , substring(users.email, 1, 3) as email_prefix")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, "    users"

    InputTsqlSubstringSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したT-SQL ROUNDの入力SQLを返す
Private Function InputTsqlRoundSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "    orders.order_id"
    AppendA5M2Line resultText, "    , round(orders.amount, 0) as amount_round0"
    AppendA5M2Line resultText, TS("    , round(orders.amount, 2, 1) as amount_truncate2")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, "    orders"

    InputTsqlRoundSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したT-SQL SUMの入力SQLを返す
Private Function InputTsqlSumSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "    orders.user_id"
    AppendA5M2Line resultText, TS("    , sum(orders.amount) as total_amount")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, TS("    orders")
    AppendA5M2Line resultText, "group by"
    AppendA5M2Line resultText, TS("    orders.user_id")
    AppendA5M2Line resultText, "having"
    AppendA5M2Line resultText, "    sum(orders.amount) > 0"

    InputTsqlSumSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したT-SQL REPLACEの入力SQLを返す
Private Function InputTsqlReplaceSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "    users.user_id"
    AppendA5M2Line resultText, TS("    , replace (users.email, '@old.example', '@new.example') as normalized_email")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, "    users"

    InputTsqlReplaceSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したT-SQL DATEADDの入力SQLを返す
Private Function InputTsqlDateAddSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "    orders.order_id"
    AppendA5M2Line resultText, "    , dateadd(day, 7, created_at) as due_date"
    AppendA5M2Line resultText, TS("    , dateadd(month, 1, created_at) as next_month_date")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, "    orders"

    InputTsqlDateAddSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したT-SQL DATEDIFFの入力SQLを返す
Private Function InputTsqlDateDiffSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "    orders.order_id"
    AppendA5M2Line resultText, "    , datediff(day, created_at, updated_at) as elapsed_days"
    AppendA5M2Line resultText, TS("    , datediff(minute, created_at, updated_at) as elapsed_minutes")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, "    orders"

    InputTsqlDateDiffSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したT-SQL COUNTの入力SQLを返す
Private Function InputTsqlCountSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "    users.user_id"
    AppendA5M2Line resultText, TS("    , count(orders.order_id) as order_count")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, TS("    users")
    AppendA5M2Line resultText, TS("    left join orders")
    AppendA5M2Line resultText, TS("        on users.user_id = orders.user_id")
    AppendA5M2Line resultText, "group by"
    AppendA5M2Line resultText, "    users.user_id"

    InputTsqlCountSql = FinishA5M2Sql(resultText)
End Function

' A5M2で整形したT-SQL EXISTSの入力SQLを返す
Private Function InputTsqlExistsSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, TS("    users.user_id")
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, TS("    users")
    AppendA5M2Line resultText, "where"
    AppendA5M2Line resultText, TS("    exists (")
    AppendA5M2Line resultText, "        select"
    AppendA5M2Line resultText, TS("            1")
    AppendA5M2Line resultText, "        from"
    AppendA5M2Line resultText, TS("            orders")
    AppendA5M2Line resultText, "        where"
    AppendA5M2Line resultText, TS("            orders.user_id = users.user_id")
    AppendA5M2Line resultText, "            and orders.amount > 0"
    AppendA5M2Line resultText, "    )"

    InputTsqlExistsSql = FinishA5M2Sql(resultText)
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

' T-SQL TRIM FROMの和名変換後期待値を返す
Private Function ExpectedTsqlTrimFromSql() As String
    ExpectedTsqlTrimFromSql = ConvertFixtureSql(InputTsqlTrimFromSql())
End Function

' T-SQL INの和名変換後期待値を返す
Private Function ExpectedTsqlInSql() As String
    ExpectedTsqlInSql = ConvertFixtureSql(InputTsqlInSql())
End Function

' T-SQL COALESCEの和名変換後期待値を返す
Private Function ExpectedTsqlCoalesceSql() As String
    ExpectedTsqlCoalesceSql = ConvertFixtureSql(InputTsqlCoalesceSql())
End Function

' T-SQL FORMATの和名変換後期待値を返す
Private Function ExpectedTsqlFormatSql() As String
    ExpectedTsqlFormatSql = ConvertFixtureSql(InputTsqlFormatSql())
End Function

' T-SQL WITHの和名変換後期待値を返す
Private Function ExpectedTsqlWithSql() As String
    ExpectedTsqlWithSql = ConvertFixtureSql(InputTsqlWithSql())
End Function

' T-SQL CASTの和名変換後期待値を返す
Private Function ExpectedTsqlCastSql() As String
    ExpectedTsqlCastSql = ConvertFixtureSql(InputTsqlCastSql())
End Function

' T-SQL ISNULLの和名変換後期待値を返す
Private Function ExpectedTsqlIsNullSql() As String
    ExpectedTsqlIsNullSql = ConvertFixtureSql(InputTsqlIsNullSql())
End Function

' T-SQL SUBSTRINGの和名変換後期待値を返す
Private Function ExpectedTsqlSubstringSql() As String
    ExpectedTsqlSubstringSql = ConvertFixtureSql(InputTsqlSubstringSql())
End Function

' T-SQL ROUNDの和名変換後期待値を返す
Private Function ExpectedTsqlRoundSql() As String
    ExpectedTsqlRoundSql = ConvertFixtureSql(InputTsqlRoundSql())
End Function

' T-SQL SUMの和名変換後期待値を返す
Private Function ExpectedTsqlSumSql() As String
    ExpectedTsqlSumSql = ConvertFixtureSql(InputTsqlSumSql())
End Function

' T-SQL REPLACEの和名変換後期待値を返す
Private Function ExpectedTsqlReplaceSql() As String
    ExpectedTsqlReplaceSql = ConvertFixtureSql(InputTsqlReplaceSql())
End Function

' T-SQL DATEADDの和名変換後期待値を返す
Private Function ExpectedTsqlDateAddSql() As String
    ExpectedTsqlDateAddSql = ConvertFixtureSql(InputTsqlDateAddSql())
End Function

' T-SQL DATEDIFFの和名変換後期待値を返す
Private Function ExpectedTsqlDateDiffSql() As String
    ExpectedTsqlDateDiffSql = ConvertFixtureSql(InputTsqlDateDiffSql())
End Function

' T-SQL COUNTの和名変換後期待値を返す
Private Function ExpectedTsqlCountSql() As String
    ExpectedTsqlCountSql = ConvertFixtureSql(InputTsqlCountSql())
End Function

' T-SQL EXISTSの和名変換後期待値を返す
Private Function ExpectedTsqlExistsSql() As String
    ExpectedTsqlExistsSql = ConvertFixtureSql(InputTsqlExistsSql())
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
    resultText = ReplaceQualifiedFixture(resultText, "order_items.product_id", "order_items." & ProductIdText())
    resultText = ReplaceQualifiedFixture(resultText, "order_items.quantity", "order_items." & QuantityText())
    resultText = ReplaceQualifiedFixture(resultText, "order_items.order_id", "order_items." & DetailOrderIdText())
    resultText = ReplaceQualifiedFixture(resultText, "users.manager_id", "users." & ManagerIdText())
    resultText = ReplaceQualifiedFixture(resultText, "manager.user_id", "manager." & UserIdText())
    resultText = ReplaceQualifiedFixture(resultText, "manager.status", "manager." & StatusText())
    resultText = ReplaceQualifiedFixture(resultText, "manager.name", "manager." & FullNameText())
    resultText = ReplaceQualifiedFixture(resultText, "orders.order_id", "orders." & OrderIdText())
    resultText = ReplaceQualifiedFixture(resultText, "orders.user_id", "orders." & OrderUserIdText())
    resultText = ReplaceQualifiedFixture(resultText, "orders.amount", "orders." & AmountText())
    resultText = ReplaceQualifiedFixture(resultText, "users.user_id", "users." & UserIdText())
    resultText = ReplaceQualifiedFixture(resultText, "users.email", "users." & MailText())
    resultText = ReplaceQualifiedFixture(resultText, "users.name", "users." & FullNameText())
    resultText = ReplaceStandaloneFixture(resultText, "created_at", CreatedAtText())
    resultText = ReplaceStandaloneFixture(resultText, "updated_at", UpdatedAtText())
    resultText = ReplaceStandaloneFixture(resultText, "status", StatusText())

    ConvertFixtureSql = resultText
End Function

' テーブル修飾付き定義を識別子単位で置換
Private Function ReplaceQualifiedFixture(ByVal sourceText As String, ByVal searchText As String, ByVal replacementText As String) As String
    ReplaceQualifiedFixture = ReplaceRegexFixture(sourceText, "(^|[^A-Za-z0-9_])" & EscapeRegexFixture(searchText) & "([^A-Za-z0-9_]|$)", replacementText)
End Function

' 単体フィールド定義を識別子単位で置換
Private Function ReplaceStandaloneFixture(ByVal sourceText As String, ByVal searchText As String, ByVal replacementText As String) As String
    ReplaceStandaloneFixture = ReplaceRegexFixture(sourceText, "(^|[^A-Za-z0-9_.])" & EscapeRegexFixture(searchText) & "([^A-Za-z0-9_.]|$)", replacementText)
End Function

' 前後の区切り文字を残して正規表現置換
Private Function ReplaceRegexFixture(ByVal sourceText As String, ByVal patternText As String, ByVal replacementText As String) As String
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
    re.Pattern = patternText

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

    ReplaceRegexFixture = resultText
End Function

' テスト期待値用の正規表現リテラルをエスケープ
Private Function EscapeRegexFixture(ByVal value As String) As String
    Dim index As Long
    Dim resultText As String
    Dim currentChar As String

    For index = 1 To Len(value)
        currentChar = Mid$(value, index, 1)
        If InStr(1, "\.^$|?*+()[]{}", currentChar, vbBinaryCompare) > 0 Then
            resultText = resultText & "\"
        End If
        resultText = resultText & currentChar
    Next index

    EscapeRegexFixture = resultText
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

Private Sub AssertOutputSheetGridlinesHidden()
    Dim previousSheet As Object
    Dim wsOutput As Worksheet

    Set previousSheet = ActiveSheet
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    ' 目盛り線の表示状態は対象シートを表示して確認
    wsOutput.Activate
    If ActiveWindow.DisplayGridlines Then
        Fail "Output sheet gridlines should be hidden."
    End If
    previousSheet.Activate
End Sub

Private Sub AssertOutputSheetFont()
    Dim wsOutput As Worksheet

    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    AssertCellFont wsOutput.Cells(1, 1), OutputFontName(), 9
    AssertCellFont wsOutput.Cells(20, 5), OutputFontName(), 9
End Sub

Private Sub AssertCellFont(ByVal cell As Range, ByVal expectedName As String, ByVal expectedSize As Double)
    If CStr(cell.Font.Name) <> expectedName Then
        Fail cell.Worksheet.Name & "!" & cell.Address(False, False) & _
            " font expected=[" & expectedName & "] actual=[" & CStr(cell.Font.Name) & "]"
    End If
    If CDbl(cell.Font.Size) <> expectedSize Then
        Fail cell.Worksheet.Name & "!" & cell.Address(False, False) & _
            " font size expected=[" & CStr(expectedSize) & "] actual=[" & CStr(cell.Font.Size) & "]"
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

Private Function OutputFontName() As String
    OutputFontName = W(&HFF2D, &HFF33, &H20, &H30B4, &H30B7, &H30C3, &H30AF)
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
