Attribute VB_Name = "SqlAnalysisFormatterTests"
Option Explicit

'@TestModule
'@Folder("Tests")

Private Const COL_SQL As Long = 1
Private Const COL_RESULT As Long = 2
Private Const COL_REPLACEMENT As Long = 3
Private Const OUTPUT_FILL_COLOR As Long = &HEFCEF2

' テスト一式をまとめて実行
Public Sub RunAllSqlAnalysisFormatterTests(Optional ByVal showMessage As Boolean = True)
    On Error GoTo TestFail

    SetupWorkbook_CreatesOutputSheet
    SetupWorkbook_CreatesOutputTwoLayout
    SetupWorkbook_CreatesStableSqlActionButtons
    SetupWorkbook_ProvidesToastForm
    SetupWorkbook_TracksMissingNameFillColor
    CompletionToast_UsesTwoSecondDuration
    CompletionToast_UsesReadableBlueStyle
    CompletionToast_ShowsWithoutChangingSelection
    CompletionToast_ReplacesExistingNotification
    CompletionToast_DismissesImmediately
    CompletionToast_AutoDismissesAfterTwoSeconds
    AnalyzeQueries_ShowsCompletionToastOnSuccess
    AnalyzeQueries_ShowMessageFalseDoesNotShowToast
    CopyOutput_CopiesRenderedRange
    CopyOutputTwo_CopiesRenderedRange
    AnalyzeQueries_ConvertsCrudFixtures
    AnalyzeQueries_ConvertsTsqlFunctionFixtures
    AnalyzeQueries_ProcessesQueriesWithoutMappings
    AnalyzeQueries_RendersMatchedInputAndOutputTables
    AnalyzeQueries_ClassifiesModificationTargetsByRole
    AnalyzeQueries_UsesOutputTwoNameForExactMissingReference
    AnalyzeQueries_RenamesDuplicateUnionAliases
    AnalyzeQueries_RendersSupportedTablesDuringPartialFallback
    AnalyzeQueries_SortsOutputTwoByCompositeTableNumber
    AnalyzeQueries_LeavesOutputTwoHeaderOnlyOnFallback
    AnalyzeQueries_WritesWithSubqueriesInsideOut
    AnalyzeQueries_PreservesLeadingApostropheInOutput
    AnalyzeQueries_DisablesWrappingAfterWritingLongText
    AnalyzeQueries_RendersDeeplyNestedCaseConditions
    AnalyzeQueries_NormalizesInvisibleOutputWhitespace
    AnalyzeQueries_ResolvesQualifiedStarAndMatchingAlias
    AnalyzeQueries_QualifiesUnqualifiedSelectColumns
    AnalyzeQueries_QualifiesStandaloneColumnThroughTableName
    AnalyzeQueries_WritesQualifiedReplacementValuesOnce
    AnalyzeQueries_MapsQualifiedReplacementsToMultilineRows
    AnalyzeQueries_PreservesReplacementValuesOnParserFallback
    AnalyzeQueries_FocusesSyntaxFallbackSqlRow
    AnalyzeQueries_RendersInsertSelectWithoutColumnList
    AnalyzeQueries_ResolvesMatchingTemporaryTableDefinition
    AnalyzeQueries_PreservesUnmatchedTemporaryTableDefinition
    AnalyzeQueries_SeparatesTransferExpressionsFromColumns
    AnalyzeQueries_RendersWrappedUpdateCaseAsTransferMethod
    AnalyzeQueries_HandlesSyntaxCharactersInFieldNames
    AnalyzeQueries_UsesStandaloneTableNameForSingleTable
    AnalyzeQueries_WritesUnsupportedQueryAsIs
    AnalyzeQueries_FramesOnlyTableBody
    AnalyzeQueries_RestoresSqlActionButtonsAfterLargeInput
    AnalyzeQueries_RestoresApplicationStateAfterOutputError
    ClearConfirmMessage_UsesAnalysisResultWording
    ClearData_ShowMessageFalseDismissesExistingToast
    ClearData_ClearsOutputSheet
    ClearData_InitializesOutputTwoAndPreservesTableList
    ClearData_RestoresSqlActionButtons

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
' 完了通知用のモードレスUserFormが配布ブックに組み込まれていることを確認
Public Sub SetupWorkbook_ProvidesToastForm()
    Dim toastComponent As Object

    On Error Resume Next
    Set toastComponent = ThisWorkbook.VBProject.VBComponents.Item("SqlAnalysisToast")
    On Error GoTo 0

    If toastComponent Is Nothing Then
        Fail "SqlAnalysisToast form is missing from the workbook."
    End If
    If CLng(toastComponent.Type) <> 3 Then
        Fail "SqlAnalysisToast should be a VBA UserForm."
    End If
End Sub

'@TestMethod("CompletionToast")
Public Sub CompletionToast_UsesTwoSecondDuration()
    If SqlAnalysisToastManager.ToastDurationSeconds() <> 2 Then
        Fail "Completion toast duration should be two seconds."
    End If
End Sub

'@TestMethod("CompletionToast")
Public Sub CompletionToast_UsesReadableBlueStyle()
    Dim toastForm As Object
    Dim loadedForm As Object
    Dim messageLabel As Object
    Dim verticalOffset As Double

    On Error GoTo TestFail
    SqlAnalysisToastManager.ShowToast "Toast style test"
    For Each loadedForm In VBA.UserForms
        If TypeName(loadedForm) = "SqlAnalysisToast" Then
            Set toastForm = loadedForm
            Exit For
        End If
    Next loadedForm
    If toastForm Is Nothing Then
        Fail "SqlAnalysisToast should be loaded after ShowToast."
    End If
    Set messageLabel = toastForm.Controls("MessageLabel")

    If CLng(toastForm.BackColor) <> RGB(221, 235, 247) Then
        Fail "Completion toast background should use #DDEBF7."
    End If
    If CLng(messageLabel.ForeColor) <> RGB(31, 78, 120) Then
        Fail "Completion toast text should use #1F4E78."
    End If
    If CStr(messageLabel.Font.Name) <> "Yu Gothic UI" Then
        Fail "Completion toast should use Yu Gothic UI."
    End If
    If CDbl(messageLabel.Font.Size) <> 12# Then
        Fail "Completion toast font size should be 12 points."
    End If
    If CDbl(messageLabel.Height) < 24# Then
        Fail "Completion toast label should be tall enough for 12-point text."
    End If
    If Not CBool(messageLabel.Font.Bold) Then
        Fail "Completion toast text should be bold."
    End If
    verticalOffset = Abs( _
        (CDbl(messageLabel.Top) + CDbl(messageLabel.Height) / 2#) - _
        CDbl(toastForm.InsideHeight) / 2#)
    If verticalOffset > 1# Then
        Fail "Completion toast text should be vertically centered."
    End If

TestCleanUp:
    SqlAnalysisToastManager.DismissToast
    Exit Sub

TestFail:
    SqlAnalysisToastManager.DismissToast
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'@TestMethod("CompletionToast")
Public Sub CompletionToast_ShowsWithoutChangingSelection()
    Dim wsSql As Worksheet
    Dim selectedAddress As String
    Dim message As String

    On Error GoTo TestFail
    SetupWorkbook
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    wsSql.Activate
    wsSql.Range("D10").Select
    selectedAddress = ActiveCell.Address
    message = "Toast focus test"

    SqlAnalysisToastManager.ShowToast message

    If Not ActiveSheet Is wsSql Then
        Fail "Completion toast should not change the active sheet."
    End If
    If ActiveCell.Address <> selectedAddress Then
        Fail "Completion toast should not change the selected cell."
    End If
    If Not SqlAnalysisToastManager.ToastIsVisible() Then
        Fail "Completion toast should be visible after ShowToast."
    End If
    If SqlAnalysisToastManager.CurrentToastMessage() <> message Then
        Fail "Completion toast should expose the displayed message."
    End If
    If Not SqlAnalysisToastManager.ToastWindowStyleIsValid() Then
        Fail "Completion toast should be borderless and must not activate itself."
    End If

TestCleanUp:
    SqlAnalysisToastManager.DismissToast
    Exit Sub

TestFail:
    SqlAnalysisToastManager.DismissToast
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'@TestMethod("CompletionToast")
Public Sub CompletionToast_ReplacesExistingNotification()
    Dim firstDismissalTime As Date
    Dim secondDismissalTime As Date
    Dim restartAt As Date

    On Error GoTo TestFail
    SqlAnalysisToastManager.DismissToast

    SqlAnalysisToastManager.ShowToast "First toast"
    firstDismissalTime = SqlAnalysisToastManager.CurrentToastDismissalTime()
    restartAt = DateAdd("s", 1, Now)
    Do While Now < restartAt
        DoEvents
    Loop
    SqlAnalysisToastManager.ShowToast "Second toast"
    secondDismissalTime = SqlAnalysisToastManager.CurrentToastDismissalTime()

    If VBA.UserForms.Count <> 1 Then
        Fail "Repeated completion notices should reuse one toast form."
    End If
    If SqlAnalysisToastManager.CurrentToastMessage() <> "Second toast" Then
        Fail "Repeated completion notices should replace the displayed message."
    End If
    If secondDismissalTime <= firstDismissalTime Then
        Fail "Repeated completion notices should restart the two-second timer."
    End If

TestCleanUp:
    SqlAnalysisToastManager.DismissToast
    Exit Sub

TestFail:
    SqlAnalysisToastManager.DismissToast
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'@TestMethod("CompletionToast")
Public Sub CompletionToast_DismissesImmediately()
    SqlAnalysisToastManager.ShowToast "Dismiss toast"
    SqlAnalysisToastManager.DismissToast

    If SqlAnalysisToastManager.ToastIsVisible() Then
        Fail "DismissToast should immediately hide the completion toast."
    End If
    If VBA.UserForms.Count <> 0 Then
        Fail "DismissToast should unload the completion toast form."
    End If
End Sub

'@TestMethod("CompletionToast")
Public Sub CompletionToast_AutoDismissesAfterTwoSeconds()
    Dim timeoutAt As Date

    On Error GoTo TestFail
    SqlAnalysisToastManager.ShowToast "Auto dismiss toast"
    timeoutAt = DateAdd("s", 4, Now)
    Do While SqlAnalysisToastManager.ToastIsVisible() And Now < timeoutAt
        DoEvents
    Loop

    If SqlAnalysisToastManager.ToastIsVisible() Then
        Fail "Completion toast should automatically close after two seconds."
    End If
    Exit Sub

TestFail:
    SqlAnalysisToastManager.DismissToast
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'@TestMethod("AnalyzeQueries")
Public Sub AnalyzeQueries_ShowMessageFalseDoesNotShowToast()
    Dim wsSql As Worksheet

    SetupWorkbook
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    wsSql.Range("A2:Z20").ClearContents
    SqlAnalysisToastManager.ShowToast "Existing toast"

    AnalyzeQueries False

    If SqlAnalysisToastManager.ToastIsVisible() Then
        Fail "AnalyzeQueries False should not display a completion toast."
    End If
End Sub

'@TestMethod("ClearData")
Public Sub ClearData_ShowMessageFalseDismissesExistingToast()
    SetupWorkbook
    SqlAnalysisToastManager.ShowToast "Existing toast"

    ClearData False

    If SqlAnalysisToastManager.ToastIsVisible() Then
        Fail "ClearData False should dismiss an existing completion toast."
    End If
End Sub

'@TestMethod("AnalyzeQueries")
Public Sub AnalyzeQueries_ShowsCompletionToastOnSuccess()
    Dim wsSql As Worksheet
    Dim expectedMessage As String

    On Error GoTo TestFail
    SqlAnalysisToastManager.DismissToast
    SetupWorkbook
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    wsSql.Range("A2:Z20").ClearContents
    expectedMessage = W(&H89E3, &H6790, &H304C, &H5B8C, &H4E86, _
        &H3057, &H307E, &H3057, &H305F, &H3002)

    AnalyzeQueries True

    If Not SqlAnalysisToastManager.ToastIsVisible() Then
        Fail "Successful analysis should display a completion toast."
    End If
    If SqlAnalysisToastManager.CurrentToastMessage() <> expectedMessage Then
        Fail "Successful analysis should display the analysis completion message."
    End If

TestCleanUp:
    SqlAnalysisToastManager.DismissToast
    Exit Sub

TestFail:
    SqlAnalysisToastManager.DismissToast
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'@TestMethod("SetupWorkbook")
Public Sub SetupWorkbook_CreatesStableSqlActionButtons()
    Dim wsSql As Worksheet

    SetupWorkbook
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())

    AssertSqlActionButtons wsSql
End Sub

'@TestMethod("AnalyzeQueries")
Public Sub AnalyzeQueries_RestoresSqlActionButtonsAfterLargeInput()
    Const SELECT_ITEM_COUNT As Long = 500

    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim itemNumber As Long
    Dim lastRow As Long
    Dim testStage As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean

    On Error GoTo TestFail

    If Not ExternalParserConfigured() Then Exit Sub

    testStage = "setup"
    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    wsRef.Range("A2:D1000").ClearContents
    wsSql.Range("A2:Z2000").ClearContents

    testStage = "prepare large SQL"
    wsSql.Cells(2, COL_SQL).Value = "SELECT"
    For itemNumber = 1 To SELECT_ITEM_COUNT
        wsSql.Cells(itemNumber + 2, COL_SQL).Value = _
            IIf(itemNumber = 1, "    ", "    , ") & _
            "tb1.col" & CStr(itemNumber) & " AS item" & CStr(itemNumber)
    Next itemNumber
    lastRow = SELECT_ITEM_COUNT + 3
    wsSql.Cells(lastRow, COL_SQL).Value = "FROM users AS tb1;"

    testStage = "simulate hidden buttons"
    wsSql.Shapes("btnAnalyzeQueries").Delete
    wsSql.Columns("E:F").Hidden = True
    ThisWorkbook.DisplayDrawingObjects = xlHide

    testStage = "analyze large SQL"
    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    AnalyzeQueries False

    testStage = "verify completed analysis"
    AssertCellValue wsSql.Cells(lastRow, COL_RESULT), "FROM users AS tb1;"
    If Application.ScreenUpdating <> previousScreenUpdating Then
        Fail "AnalyzeQueries should restore ScreenUpdating after a large input."
    End If
    If Application.EnableEvents <> previousEnableEvents Then
        Fail "AnalyzeQueries should restore EnableEvents after a large input."
    End If
    testStage = "verify repaired buttons"
    AssertSqlActionButtons wsSql
    wsSql.Columns("E:F").Hidden = False
    Exit Sub

TestFail:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    On Error Resume Next
    wsSql.Columns("E:F").Hidden = False
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, testStage & ": " & errorDescription
End Sub

'@TestMethod("AnalyzeQueries")
Public Sub AnalyzeQueries_RestoresApplicationStateAfterOutputError()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim previousStatusBar As Variant
    Dim analyzeErrorNumber As Long
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo TestFail
    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())
    wsRef.Range("A2:D20").ClearContents
    wsSql.Range("A2:Z20").ClearContents
    wsSql.Cells(2, COL_SQL).Value = "SELECT 1;"
    wsSql.Shapes("btnAnalyzeQueries").Delete
    ThisWorkbook.DisplayDrawingObjects = xlHide
    wsOutput.Protect

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    previousStatusBar = Application.StatusBar
    On Error Resume Next
    AnalyzeQueries False
    analyzeErrorNumber = Err.Number
    Err.Clear
    On Error GoTo TestFail
    wsOutput.Unprotect

    If analyzeErrorNumber = 0 Then
        Fail "AnalyzeQueries should report a protected output sheet."
    End If
    If Application.ScreenUpdating <> previousScreenUpdating Then
        Fail "AnalyzeQueries should restore ScreenUpdating after an error."
    End If
    If Application.EnableEvents <> previousEnableEvents Then
        Fail "AnalyzeQueries should restore EnableEvents after an error."
    End If
    If StrComp(CStr(Application.StatusBar), CStr(previousStatusBar), vbTextCompare) <> 0 Then
        Fail "AnalyzeQueries should restore StatusBar after an error. expected=[" & _
            CStr(previousStatusBar) & "] actual=[" & CStr(Application.StatusBar) & "]"
    End If
    AssertSqlActionButtons wsSql
    Exit Sub

TestFail:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    On Error Resume Next
    If Not wsOutput Is Nothing Then wsOutput.Unprotect
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

'@TestMethod("ClearData")
Public Sub ClearData_RestoresSqlActionButtons()
    Dim wsSql As Worksheet

    SetupWorkbook
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    wsSql.Shapes("btnClearData").Delete
    ThisWorkbook.DisplayDrawingObjects = xlHide

    ClearData False

    AssertSqlActionButtons wsSql
End Sub

'@TestMethod("AnalyzeQueries")
Public Sub AnalyzeQueries_LeavesOutputTwoHeaderOnlyOnFallback()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim wsTableList As Worksheet

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetTwoName())
    Set wsTableList = ThisWorkbook.Worksheets(TableListSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsTableList.Range("A2:C200").ClearContents
    PutTableListRow wsTableList, 2, "users", UserTableText(), "1-1"
    wsSql.Cells(2, COL_SQL).Value = "CREATE TABLE dbo.users (id int)"

    AnalyzeQueries False

    AssertCellValue wsOutput.Range("A1"), InputInformationTitle()
    AssertCellValue wsOutput.Range("BA1"), OutputInformationTitle()
    AssertCellValue wsOutput.Range("A3"), "No"
    AssertCellValue wsOutput.Range("BA3"), "No"
    AssertCellValue wsOutput.Range("A4"), ""
    AssertCellValue wsOutput.Range("C4"), ""
    AssertCellValue wsOutput.Range("BA4"), ""
    AssertCellValue wsOutput.Range("BC4"), ""
    wsTableList.Range("A2:C200").ClearContents
End Sub

'@TestMethod("AnalyzeQueries")
' 変換定義が空でも物理名のまま解析表とアウトプットを作成できることを確認
Public Sub AnalyzeQueries_ProcessesQueriesWithoutMappings()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim missingNameText As String
    Dim sqlText As String

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    sqlText = "SELECT tb1.name FROM users AS tb1"
    wsSql.Cells(2, COL_SQL).Value = sqlText

    AnalyzeQueries False

    missingNameText = "(" & W(&H548C, &H540D, &H672A, &H53D6, &H5F97) & ")"
    AssertCellValue wsSql.Cells(2, COL_RESULT), sqlText
    AssertCellValue wsOutput.Cells(1, 1), SelectOutputTitle()
    AssertCellValue wsOutput.Cells(2, 1), _
        ReferenceTablesText() & ": " & missingNameText & "[tb1]"
    AssertCellValue wsOutput.Cells(3, 17), "tb1.name"
End Sub

'@TestMethod("AnalyzeQueries")
Public Sub AnalyzeQueries_RendersMatchedInputAndOutputTables()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim wsTableList As Worksheet

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetTwoName())
    Set wsTableList = ThisWorkbook.Worksheets(TableListSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsTableList.Range("A2:C200").ClearContents
    PutTableListRow wsTableList, 2, "[users]", UserTableText(), "1-1"
    PutTableListRow wsTableList, 3, "ORDERS", OrderTableText(), "2-1"
    PutTableListRow wsTableList, 4, "[archive_users]", "archive users", "3-1"
    ' Duplicate matching is case-insensitive and the first row wins.
    PutTableListRow wsTableList, 5, "USERS", "duplicate user", "9-9"
    wsSql.Cells(2, COL_SQL).Value = _
        "INSERT INTO [archive].[archive_users] (name) " & _
        "SELECT u.name FROM [dbo].[users] AS u " & _
        "INNER JOIN [sales].[orders] AS o ON u.user_id = o.user_id " & _
        "LEFT JOIN [audit].[audit_logs] AS l ON u.user_id = l.user_id"

    AnalyzeQueries False

    AssertCellValue wsOutput.Range("A4"), "1"
    AssertCellValue wsOutput.Range("C4"), "users"
    AssertCellValue wsOutput.Range("S4"), UserTableText()
    AssertCellValue wsOutput.Range("AR4"), "1-1"
    AssertCellValue wsOutput.Range("A5"), "2"
    AssertCellValue wsOutput.Range("C5"), "orders"
    AssertCellValue wsOutput.Range("S5"), OrderTableText()
    AssertCellValue wsOutput.Range("AR5"), "2-1"
    AssertCellValue wsOutput.Range("C6"), ""
    AssertCellValue wsOutput.Range("BA4"), "1"
    AssertCellValue wsOutput.Range("BC4"), "archive_users"
    AssertCellValue wsOutput.Range("BS4"), "archive users"
    AssertCellValue wsOutput.Range("CR4"), "3-1"
    AssertCellValue wsOutput.Range("BC5"), ""

    AssertMergedArea wsOutput.Range("A4"), "$A$4:$B$4"
    AssertMergedArea wsOutput.Range("A5"), "$A$5:$B$5"
    AssertMergedArea wsOutput.Range("BA4"), "$BA$4:$BB$4"
    AssertCellNotMerged wsOutput.Range("C4")
    AssertCellNotMerged wsOutput.Range("S4")
    AssertCellNotMerged wsOutput.Range("AR4")
    AssertCellNotMerged wsOutput.Range("BC4")
    AssertCellNotMerged wsOutput.Range("BS4")
    AssertCellNotMerged wsOutput.Range("CR4")
    AssertHorizontalAlignment wsOutput.Range("A4"), xlCenter
    AssertHorizontalAlignment wsOutput.Range("BA4"), xlCenter
    AssertHorizontalAlignment wsOutput.Range("C4"), xlLeft
    AssertHorizontalAlignment wsOutput.Range("BC4"), xlLeft
    AssertDataBlock wsOutput.Range("A4:B4"), False
    AssertDataBlock wsOutput.Range("A5:B5"), False
    AssertDataBlock wsOutput.Range("C4:R4")
    AssertDataBlock wsOutput.Range("S4:AQ4")
    AssertDataBlock wsOutput.Range("AR4:AV4")
    AssertDataBlock wsOutput.Range("BA4:BB4"), False
    AssertDataBlock wsOutput.Range("BC4:BR4")
    AssertDataBlock wsOutput.Range("BS4:CQ4")
    AssertDataBlock wsOutput.Range("CR4:CV4")
    AssertBlankSeparatorRange wsOutput.Range("AW1:AZ6")
    wsTableList.Range("A2:C200").ClearContents
End Sub

'@TestMethod("AnalyzeQueries")
' UPDATEの対象行セットを入力から除外し、独立した自己参照だけを入力へ含めることを確認
Public Sub AnalyzeQueries_ClassifiesModificationTargetsByRole()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim wsTableList As Worksheet

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetTwoName())
    Set wsTableList = ThisWorkbook.Worksheets(TableListSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsTableList.Range("A2:C200").ClearContents
    PutTableListRow wsTableList, 2, "locations", "location master", "1-1"
    PutTableListRow wsTableList, 3, "users", "user master", "2-1"
    wsSql.Cells(2, COL_SQL).Value = _
        "UPDATE u SET location_name = l.name " & _
        "FROM users AS u JOIN locations AS l ON l.id = u.location_id;"

    AnalyzeQueries False

    AssertOutputTwoRow wsOutput, 4, 1, "locations", "location master", "1-1"
    AssertCellValue wsOutput.Range("C5"), ""
    AssertOutputTwoRow wsOutput, 4, 53, "users", "user master", "2-1"
    AssertCellValue wsOutput.Range("BC5"), ""

    wsSql.Range("A2:Z200").ClearContents
    wsSql.Cells(2, COL_SQL).Value = _
        "UPDATE users SET name = " & _
        "(SELECT TOP (1) source.name FROM users AS source " & _
        "WHERE source.id <> users.id);"

    AnalyzeQueries False

    AssertOutputTwoRow wsOutput, 4, 1, "users", "user master", "2-1"
    AssertCellValue wsOutput.Range("C5"), ""
    AssertOutputTwoRow wsOutput, 4, 53, "users", "user master", "2-1"
    AssertCellValue wsOutput.Range("BC5"), ""
    wsTableList.Range("A2:C200").ClearContents
End Sub

'@TestMethod("AnalyzeQueries")
' アウトプット②とIDが完全一致する未取得参照だけをテーブル名称へ置換することを確認
Public Sub AnalyzeQueries_UsesOutputTwoNameForExactMissingReference()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim wsOutputTwo As Worksheet
    Dim wsTableList As Worksheet
    Dim missingNameText As String
    Dim definitionLocationName As String
    Dim expectedReference As String

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())
    Set wsOutputTwo = ThisWorkbook.Worksheets(OutputSheetTwoName())
    Set wsTableList = ThisWorkbook.Worksheets(TableListSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    wsTableList.Range("A2:C200").ClearContents
    definitionLocationName = "definition location"
    PutDefinition wsRef, 2, "locations", definitionLocationName, "", ""
    PutTableListRow wsTableList, 2, "[USERS]", "user master", "1-1"
    PutTableListRow wsTableList, 3, "orders", "order master", "2-1"
    PutTableListRow wsTableList, 4, "audit_logs", "", "3-1"
    PutTableListRow wsTableList, 5, "locations", "table-list location", "4-1"
    wsSql.Cells(2, COL_SQL).Value = _
        "SELECT users.name FROM users " & _
        "JOIN orders AS o ON users.id = o.user_id " & _
        "JOIN audit_logs ON users.id = audit_logs.user_id " & _
        "JOIN locations ON users.id = locations.user_id;"

    AnalyzeQueries False

    missingNameText = "(" & W(&H548C, &H540D, &H672A, &H53D6, &H5F97) & ")"
    expectedReference = ReferenceTablesText() & ": user master" & W(&H3001) & _
        missingNameText & "[o]" & W(&H3001) & _
        missingNameText & "[audit_logs]" & W(&H3001) & _
        definitionLocationName & "[locations]"
    AssertCellValue wsOutput.Cells(2, 1), expectedReference
    AssertCellValue wsOutputTwo.Range("C4"), "users"
    AssertCellValue wsOutputTwo.Range("S4"), "user master"
    AssertCellValue wsOutputTwo.Range("C5"), "orders"
    AssertCellValue wsOutputTwo.Range("S6"), ""
    AssertCellValue wsOutputTwo.Range("S7"), "table-list location"
    wsTableList.Range("A2:C200").ClearContents
End Sub

'@TestMethod("AnalyzeQueries")
' UNION分岐で異なる物理表が同じ別名を使う場合に表示別名と物理表一覧を分離することを確認
Public Sub AnalyzeQueries_RenamesDuplicateUnionAliases()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim wsOutputTwo As Worksheet
    Dim wsTableList As Worksheet
    Dim expectedReference As String

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())
    Set wsOutputTwo = ThisWorkbook.Worksheets(OutputSheetTwoName())
    Set wsTableList = ThisWorkbook.Worksheets(TableListSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    wsTableList.Range("A2:C200").ClearContents
    PutTableListRow wsTableList, 2, "city1", "city one", "1-1"
    PutTableListRow wsTableList, 3, "city2", "city two", "1-2"
    wsSql.Cells(2, COL_SQL).Value = "SELECT tb1.* FROM city1 AS tb1"
    wsSql.Cells(3, COL_SQL).Value = "UNION"
    wsSql.Cells(4, COL_SQL).Value = "SELECT tb1.* FROM city2 AS tb1;"

    AnalyzeQueries False

    expectedReference = ReferenceTablesText() & ": city one[tb1]" & _
        W(&H3001) & "city two[tb2]"
    AssertCellValue wsOutput.Cells(2, 1), expectedReference
    AssertCellValue wsOutput.Cells(3, 17), "tb1." & W(&H5168, &H9805, &H76EE)
    AssertCellValue wsOutput.Cells(5, 17), "tb2." & W(&H5168, &H9805, &H76EE)
    AssertOutputTwoRow wsOutputTwo, 4, 1, "city1", "city one", "1-1"
    AssertOutputTwoRow wsOutputTwo, 5, 1, "city2", "city two", "1-2"
    AssertCellValue wsOutputTwo.Range("BC4"), ""
    wsTableList.Range("A2:C200").ClearContents
End Sub

'@TestMethod("AnalyzeQueries")
' 未対応ステートメントを除外し、対応できたステートメントのテーブルだけを描画することを確認
Public Sub AnalyzeQueries_RendersSupportedTablesDuringPartialFallback()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim wsTableList As Worksheet

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetTwoName())
    Set wsTableList = ThisWorkbook.Worksheets(TableListSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsTableList.Range("A2:C200").ClearContents
    PutTableListRow wsTableList, 2, "users", UserTableText(), "2-10"
    PutTableListRow wsTableList, 3, "active_users", "active users", "1-2"
    PutTableListRow wsTableList, 4, "audit_log", "audit log", "3-1"
    PutTableListRow wsTableList, 5, "ignored_table", "ignored table", "0-1"
    wsSql.Cells(2, COL_SQL).Value = _
        "SELECT * FROM dbo.users;" & vbCrLf & _
        "CREATE INDEX IX_ignored_id ON dbo.ignored_table(id);" & vbCrLf & _
        "INSERT INTO dbo.audit_log(user_id) " & _
        "SELECT id FROM dbo.active_users;"

    AnalyzeQueries False

    AssertCellValue wsOutput.Range("A4"), "1"
    AssertCellValue wsOutput.Range("C4"), "active_users"
    AssertCellValue wsOutput.Range("AR4"), "1-2"
    AssertCellValue wsOutput.Range("A5"), "2"
    AssertCellValue wsOutput.Range("C5"), "users"
    AssertCellValue wsOutput.Range("AR5"), "2-10"
    AssertCellValue wsOutput.Range("C6"), ""
    AssertCellValue wsOutput.Range("BA4"), "1"
    AssertCellValue wsOutput.Range("BC4"), "audit_log"
    AssertCellValue wsOutput.Range("CR4"), "3-1"
    AssertCellValue wsOutput.Range("BC5"), ""
    wsTableList.Range("A2:C200").ClearContents
End Sub

'@TestMethod("AnalyzeQueries")
' 番号を数値2要素として比較し、入力・出力の各ブロックを昇順へ並べることを確認
Public Sub AnalyzeQueries_SortsOutputTwoByCompositeTableNumber()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim wsTableList As Worksheet

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetTwoName())
    Set wsTableList = ThisWorkbook.Worksheets(TableListSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsTableList.Range("A2:C200").ClearContents
    PutTableListRow wsTableList, 2, "input_late", "input late", "10-1"
    PutTableListRow wsTableList, 3, "input_early", "input early", "1-10"
    PutTableListRow wsTableList, 4, "input_middle", "input middle", "1-2"
    PutTableListRow wsTableList, 5, "output_late", "output late", "10-2"
    PutTableListRow wsTableList, 6, "output_early", "output early", "2-10"
    PutTableListRow wsTableList, 7, "output_middle", "output middle", "2-2"
    PutTableListRow wsTableList, 8, "input_middle_later", "input middle later", "1-2"
    PutTableListRow wsTableList, 9, "output_middle_later", "output middle later", "2-2"
    wsSql.Cells(2, COL_SQL).Value = _
        "INSERT INTO dbo.output_late(id) SELECT id FROM dbo.input_late;" & vbCrLf & _
        "INSERT INTO dbo.output_early(id) SELECT id FROM dbo.input_early;" & vbCrLf & _
        "INSERT INTO dbo.output_middle(id) SELECT id FROM dbo.input_middle;" & vbCrLf & _
        "INSERT INTO dbo.output_middle_later(id) " & _
        "SELECT id FROM dbo.input_middle_later;"

    AnalyzeQueries False

    AssertOutputTwoRow wsOutput, 4, 1, "input_middle", "input middle", "1-2"
    AssertOutputTwoRow wsOutput, 5, 1, _
        "input_middle_later", "input middle later", "1-2"
    AssertOutputTwoRow wsOutput, 6, 1, "input_early", "input early", "1-10"
    AssertOutputTwoRow wsOutput, 7, 1, "input_late", "input late", "10-1"
    AssertOutputTwoRow wsOutput, 4, 53, "output_middle", "output middle", "2-2"
    AssertOutputTwoRow wsOutput, 5, 53, _
        "output_middle_later", "output middle later", "2-2"
    AssertOutputTwoRow wsOutput, 6, 53, "output_early", "output early", "2-10"
    AssertOutputTwoRow wsOutput, 7, 53, "output_late", "output late", "10-2"
    wsTableList.Range("A2:C200").ClearContents
End Sub

'@TestMethod("AnalyzeQueries")
' 表本体がない場合にタイトルと参照テーブル行へ罫線を付けないことを確認
Public Sub AnalyzeQueries_FramesOnlyTableBody()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    ' 前のテストが残した表本体の上罫線だけを初期化する
    wsOutput.Range("A1:CL200").Borders.LineStyle = xlNone
    PutDefinition wsRef, 2, "users", UserTableText(), "user_id", UserIdText()
    wsSql.Cells(2, COL_SQL).Value = "DELETE FROM users"

    AnalyzeQueries False

    AssertCellValue wsOutput.Cells(1, 1), DataTransferTitle()
    AssertCellValue wsOutput.Cells(2, 1), ReferenceTablesText() & ": " & UserTableText() & "[users]"
    AssertCellHasNoEdgeBorders wsOutput.Cells(1, 1)
    AssertCellHasNoEdgeBorders wsOutput.Cells(1, 90)
    AssertCellHasNoEdgeBorders wsOutput.Cells(2, 1)
    AssertCellHasNoEdgeBorders wsOutput.Cells(2, 90)
End Sub

'@TestMethod("AnalyzeQueries")
' 単独フィールド定義のテーブル和名を単一参照テーブルへ使用することを確認
Public Sub AnalyzeQueries_UsesStandaloneTableNameForSingleTable()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    PutDefinition wsRef, 2, "-", UserTableText(), "name", FullNameText()
    wsSql.Cells(2, COL_SQL).Value = "SELECT name FROM [user]"

    AnalyzeQueries False

    AssertCellValue wsSql.Cells(2, COL_RESULT), "SELECT " & FullNameText() & " FROM [user]"
    AssertCellValue wsOutput.Cells(2, 1), ReferenceTablesText() & ": " & UserTableText()
End Sub

'@TestMethod("AnalyzeQueries")
' 構文文字を含むフィールド和名でも表を出力できることを確認
Public Sub AnalyzeQueries_HandlesSyntaxCharactersInFieldNames()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim bracketSlashName As String
    Dim operatorName As String
    Dim quoteCommentName As String

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    bracketSlashName = UserIdText() & "[" & AmountText() & "]/" & StatusText()
    operatorName = "=1+1 / [" & StatusText() & "]"
    quoteCommentName = FullNameText() & "'/*main*/--current"
    PutDefinition wsRef, 2, "tb1", OrderTableText(), "amount", "  " & bracketSlashName & "  "
    PutDefinition wsRef, 3, "-", "", "status", operatorName
    PutDefinition wsRef, 4, "tb1", OrderTableText(), "owner", quoteCommentName
    wsSql.Cells(2, COL_SQL).Value = _
        "SELECT tb1.amount, status, tb1.owner FROM invoices AS tb1"

    AnalyzeQueries False

    AssertCellValue wsSql.Cells(2, COL_RESULT), _
        "SELECT tb1." & bracketSlashName & ", " & operatorName & _
        ", tb1." & quoteCommentName & " FROM invoices AS tb1"
    AssertCellValue wsOutput.Cells(1, 1), SelectOutputTitle()
    AssertCellValue wsOutput.Cells(3, 17), "tb1." & bracketSlashName
    AssertCellValue wsOutput.Cells(4, 17), operatorName
    AssertCellValue wsOutput.Cells(5, 17), "tb1." & quoteCommentName
    AssertCellValue wsSql.Cells(2, COL_REPLACEMENT + 1), operatorName
    If wsSql.Cells(2, COL_REPLACEMENT + 1).HasFormula Then
        Fail "Syntax-like field name was written as an Excel formula."
    End If
End Sub

' 自動実行時の結果をダイアログなしで返す
Public Function RunAllSqlAnalysisFormatterTestsForAutomation() As String
    On Error GoTo TestFail

    RunAllSqlAnalysisFormatterTests False
    RunAllSqlAnalysisFormatterTestsForAutomation = "OK"
    Exit Function

TestFail:
    RunAllSqlAnalysisFormatterTestsForAutomation = CStr(Err.Number) & ": " & Err.Description
End Function

'@TestMethod("AnalyzeQueries")
' 改行を含む長い文字列を書き込んだ後も折り返しと縮小表示を無効にすることを確認
Public Sub AnalyzeQueries_DisablesWrappingAfterWritingLongText()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim longText As String

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    PutDefinition wsRef, 2, "-", "", "long_text", "長文"
    longText = String(180, "A") & vbLf & String(180, "B")
    wsSql.Cells(2, COL_SQL).Value = "SELECT '" & longText & "' AS long_text"

    AnalyzeQueries False

    If CBool(wsOutput.Cells(3, 32).WrapText) Then
        Fail "Long output text should not enable wrapping."
    End If
    If CBool(wsOutput.Cells(3, 32).ShrinkToFit) Then
        Fail "Long output text should not enable shrink to fit."
    End If
End Sub

'@TestMethod("AnalyzeQueries")
' 括弧とAND/ORが混在するCASE条件を再帰的な階層で描画することを確認
Public Sub AnalyzeQueries_RendersDeeplyNestedCaseConditions()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    PutDefinition wsRef, 2, "tb1", "conditions", "a", "a"
    wsSql.Cells(2, COL_SQL).Value = _
        "SELECT CASE WHEN ((tb1.a = 1 OR tb1.b = 1) " & _
        "AND (tb1.c = 1 OR tb1.d = 1 OR tb1.e = 1)) " & _
        "OR (tb1.f = 1 AND (tb1.g = 1 OR tb1.h = 1)) " & _
        "THEN 'X' ELSE 'Y' END AS result_code FROM conditions AS tb1"

    AnalyzeQueries False

    AssertCellValue wsOutput.Cells(3, 38), "((tb1.a = 1"
    AssertCellValue wsOutput.Cells(4, 36), "OR"
    AssertCellValue wsOutput.Cells(5, 34), "AND"
    AssertCellValue wsOutput.Cells(7, 38), "tb1.e = 1))"
    AssertCellValue wsOutput.Cells(8, 32), "OR"
    AssertCellValue wsOutput.Cells(9, 34), "AND"
    AssertCellValue wsOutput.Cells(10, 36), "OR"
    AssertCellValue wsOutput.Cells(10, 38), "tb1.h = 1)) " & W(&H2192) & " 'X'"
    AssertCellValue wsOutput.Cells(11, 32), "ELSE " & W(&H2192) & " 'Y'"
End Sub

'@TestMethod("AnalyzeQueries")
' 出力するSQL断片の不可視空白を1スペースへ統一し、リテラル内の連続空白は保持することを確認
Public Sub AnalyzeQueries_NormalizesInvisibleOutputWhitespace()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    PutDefinition wsRef, 2, "tb1", "users", "name", "name"
    wsSql.Cells(2, COL_SQL).Value = _
        "SELECT CASE WHEN tb1.name" & vbTab & "  IS" & vbCrLf & _
        " NULL THEN 'A  B' ELSE 'C' END AS result_name " & _
        "FROM users AS tb1 WHERE tb1.name              = '1'"

    AnalyzeQueries False

    AssertCellValue wsOutput.Cells(3, 32), _
        "tb1.name IS NULL " & W(&H2192) & " 'A  B'"
    AssertCellValue wsOutput.Cells(4, 32), _
        "ELSE " & W(&H2192) & " 'C'"
    AssertCellValue wsOutput.Cells(5, 17), "tb1.name = '1'"
End Sub

'@TestMethod("AnalyzeQueries")
' 修飾付き全項目と参照列名に一致する式エイリアスを和名表示することを確認
Public Sub AnalyzeQueries_ResolvesQualifiedStarAndMatchingAlias()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim nameText As String
    Dim allFieldsText As String

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    nameText = W(&H540D, &H524D)
    allFieldsText = W(&H5168, &H9805, &H76EE)
    PutDefinition wsRef, 2, "tb1", UserTableText(), "name", nameText
    wsSql.Cells(2, COL_SQL).Value = _
        "SELECT tb1.*, TRIM(tb1.name) AS name FROM users AS tb1"

    AnalyzeQueries False

    AssertCellValue wsOutput.Cells(3, 17), "tb1." & allFieldsText
    AssertCellValue wsOutput.Cells(4, 17), nameText
    AssertCellValue wsOutput.Cells(4, 31), W(&H203B)
    AssertCellValue wsOutput.Cells(4, 32), "TRIM(tb1." & nameText & ")"
End Sub

'@TestMethod("AnalyzeQueries")
' SELECT INTOの一時テーブル名と和名定義が一致する場合に全項目を移送できることを確認
Public Sub AnalyzeQueries_ResolvesMatchingTemporaryTableDefinition()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim nameText As String
    Dim temporaryTableText As String
    Dim allFieldsText As String

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    nameText = W(&H540D, &H524D)
    temporaryTableText = W(&H4E00, &H6642) & UserTableText()
    allFieldsText = W(&H5168, &H9805, &H76EE)
    PutDefinition wsRef, 2, "tb1", UserTableText(), "name", nameText
    PutDefinition wsRef, 3, "#wkuser", temporaryTableText, "name", nameText
    wsSql.Cells(2, COL_SQL).Value = _
        "SELECT tb1.* INTO #wkuser FROM users AS tb1"

    AnalyzeQueries False

    AssertCellValue wsOutput.Cells(3, 17), "tb1." & allFieldsText
    AssertCellValue wsOutput.Cells(6, 1), _
        ReferenceTablesText() & ": " & temporaryTableText & W(&H3001) & _
        UserTableText() & "[tb1]"
    AssertCellValue wsOutput.Cells(8, 1), allFieldsText
    AssertCellValue wsOutput.Cells(8, 19), "tb1." & allFieldsText
End Sub

'@TestMethod("AnalyzeQueries")
' SELECT INTOの一時テーブル名と和名定義が一致しない場合に未解決表示を使用することを確認
Public Sub AnalyzeQueries_PreservesUnmatchedTemporaryTableDefinition()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim nameText As String
    Dim missingNameText As String
    Dim allFieldsText As String

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    nameText = W(&H540D, &H524D)
    missingNameText = "(" & W(&H548C, &H540D, &H672A, &H53D6, &H5F97) & ")"
    allFieldsText = W(&H5168, &H9805, &H76EE)
    PutDefinition wsRef, 2, "tb1", UserTableText(), "name", nameText
    PutDefinition wsRef, 3, "#other_work", "other", "name", nameText
    wsSql.Cells(2, COL_SQL).Value = _
        "SELECT tb1.* INTO #wkuser FROM users AS tb1"

    AnalyzeQueries False

    AssertCellValue wsOutput.Cells(3, 17), "tb1." & allFieldsText
    AssertCellValue wsOutput.Cells(6, 1), _
        ReferenceTablesText() & ": " & missingNameText & "[#wkuser]" & W(&H3001) & _
        UserTableText() & "[tb1]"
    AssertCellValue wsOutput.Cells(8, 1), allFieldsText
    AssertCellValue wsOutput.Cells(8, 19), "tb1." & allFieldsText
End Sub

'@TestMethod("AnalyzeQueries")
' 更新系の集計式と算術式を移送方法へ、参照列を読点区切りで移送元へ出力することを確認
Public Sub AnalyzeQueries_SeparatesTransferExpressionsFromColumns()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    PutDefinition wsRef, 2, "tb1", "orders", "amount", "amount"
    wsSql.Cells(2, COL_SQL).Value = _
        "INSERT INTO order_summary(total_amount, gross_amount) " & _
        "SELECT SUM(tb1.amount), tb1.amount + tb1.tax + tb1.amount " & _
        "FROM orders AS tb1"

    AnalyzeQueries False

    AssertCellValue wsOutput.Cells(9, 19), "tb1.amount"
    AssertCellValue wsOutput.Cells(9, 37), "SUM(tb1.amount)"
    AssertCellValue wsOutput.Cells(10, 19), _
        "tb1.amount" & W(&H3001) & "tb1.tax"
    AssertCellValue wsOutput.Cells(10, 37), _
        "tb1.amount + tb1.tax + tb1.amount"
End Sub

'@TestMethod("AnalyzeQueries")
' 外側の式で包まれたUPDATE CASEを移送方法へ展開し、戻り値の列を読点区切りで移送元へ出力することを確認
Public Sub AnalyzeQueries_RendersWrappedUpdateCaseAsTransferMethod()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    PutDefinition wsRef, 2, "tb1", "users", "name", "name"
    PutDefinition wsRef, 3, "tb2", "import_users", "name", "name"
    wsSql.Cells(2, COL_SQL).Value = _
        "UPDATE tb1 SET display_name = CAST(CASE " & _
        "WHEN tb2.status IS NULL THEN tb1.name ELSE tb2.name END " & _
        "AS NVARCHAR(100)) FROM users AS tb1 LEFT JOIN import_users AS tb2 " & _
        "ON tb1.user_id = tb2.user_id"

    AnalyzeQueries False

    AssertCellValue wsOutput.Cells(4, 19), _
        "tb1.name" & W(&H3001) & "tb2.name"
    AssertCellValue wsOutput.Cells(4, 37), _
        "CAST(CASE" & W(&H7D50, &H679C) & " AS NVARCHAR(100))"
    AssertCellValue wsOutput.Cells(4, 51), W(&H203B)
    AssertCellValue wsOutput.Cells(4, 52), _
        "tb2.status IS NULL " & W(&H2192) & " tb1.name"
    AssertCellValue wsOutput.Cells(5, 52), _
        "ELSE " & W(&H2192) & " tb2.name"
End Sub

'@TestMethod("SetupWorkbook")
Public Sub SetupWorkbook_CreatesOutputSheet()
    Dim wsOutput As Worksheet

    SetupWorkbook

    ' Recreate the legacy state and verify that setup migrates the sheet in place.
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())
    wsOutput.Cells(200, 1).Value = "legacy marker"
    wsOutput.Name = LegacyOutputSheetName()
    SetupWorkbook

    If ThisWorkbook.Worksheets.Count <> 5 Then
        Fail "SetupWorkbook should create exactly five worksheets."
    End If
    AssertWorksheetNameAt 1, ReferenceSheetName()
    AssertWorksheetNameAt 2, SqlSheetName()
    AssertWorksheetNameAt 3, OutputSheetName()
    AssertWorksheetNameAt 4, OutputSheetTwoName()
    AssertWorksheetNameAt 5, TableListSheetName()
    AssertWorksheetDoesNotExist LegacyOutputSheetName()
    AssertWorksheetExists OutputSheetName()
    AssertWorksheetExists OutputSheetTwoName()
    AssertWorksheetExists TableListSheetName()
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())
    AssertCellValue wsOutput.Cells(200, 1), "legacy marker"
    wsOutput.Cells(200, 1).ClearContents
    wsOutput.Cells(200, 17).WrapText = True
    wsOutput.Cells(200, 17).ShrinkToFit = True

    ' 値のない過去セルに残った文字配置も初期化されることを確認
    SetupWorkbook
    AssertOutputCopyButton wsOutput
    AssertOutputSheetGridlinesHidden
    AssertOutputSheetFont
    AssertOutputTextFittingDisabled wsOutput
End Sub

'@TestMethod("SetupWorkbook")
Public Sub SetupWorkbook_CreatesOutputTwoLayout()
    Dim wsOutput As Worksheet
    Dim wsTableList As Worksheet

    SetupWorkbook
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetTwoName())
    Set wsTableList = ThisWorkbook.Worksheets(TableListSheetName())
    wsOutput.Range("A3:AV3").Font.Bold = True
    wsOutput.Range("BA3:CV3").Font.Bold = True
    wsTableList.Range("C2").NumberFormat = "General"
    SetupWorkbook

    AssertCellValue wsOutput.Range("A1"), InputInformationTitle()
    AssertCellValue wsOutput.Range("BA1"), OutputInformationTitle()
    AssertCellValue wsOutput.Range("A2"), ""
    AssertCellValue wsOutput.Range("BA2"), ""
    AssertCellValue wsOutput.Range("A3"), "No"
    AssertCellValue wsOutput.Range("C3"), TableIdHeaderText()
    AssertCellValue wsOutput.Range("S3"), TableNameHeaderText()
    AssertCellValue wsOutput.Range("AR3"), NumberHeaderText()
    AssertCellValue wsOutput.Range("BA3"), "No"
    AssertCellValue wsOutput.Range("BC3"), TableIdHeaderText()
    AssertCellValue wsOutput.Range("BS3"), TableNameHeaderText()
    AssertCellValue wsOutput.Range("CR3"), NumberHeaderText()
    AssertCellHasNoEdgeBorders wsOutput.Range("A1")
    AssertCellHasNoEdgeBorders wsOutput.Range("BA1")
    AssertCellHasNoEdgeBorders wsOutput.Range("A2"), False
    AssertCellHasNoEdgeBorders wsOutput.Range("BA2"), False
    AssertCellFont wsOutput.Range("A1"), OutputFontName(), 9
    AssertCellFont wsOutput.Range("CV20"), OutputFontName(), 9
    If CDbl(wsOutput.Rows(1).RowHeight) <> 13.5 Then
        Fail "Output-two row height should be 13.5."
    End If
    If Abs(CDbl(wsOutput.Columns(100).ColumnWidth) - 1.14) > 0.02 Then
        Fail "Output-two column width should be 1.14."
    End If
    If CBool(wsOutput.Range("CV20").WrapText) Then
        Fail "Output-two wrap text should be disabled."
    End If
    If CBool(wsOutput.Range("CV20").ShrinkToFit) Then
        Fail "Output-two shrink to fit should be disabled."
    End If
    AssertSheetGridlinesHidden wsOutput

    AssertMergedArea wsOutput.Range("A3"), "$A$3:$B$3"
    AssertMergedArea wsOutput.Range("BA3"), "$BA$3:$BB$3"
    AssertCellNotMerged wsOutput.Range("C3")
    AssertCellNotMerged wsOutput.Range("S3")
    AssertCellNotMerged wsOutput.Range("AR3")
    AssertCellNotMerged wsOutput.Range("BC3")
    AssertCellNotMerged wsOutput.Range("BS3")
    AssertCellNotMerged wsOutput.Range("CR3")

    AssertHorizontalAlignment wsOutput.Range("A3"), xlCenter
    AssertHorizontalAlignment wsOutput.Range("BA3"), xlCenter
    AssertHorizontalAlignment wsOutput.Range("C3"), xlLeft
    AssertHorizontalAlignment wsOutput.Range("S3"), xlLeft
    AssertHorizontalAlignment wsOutput.Range("AR3"), xlLeft
    AssertHorizontalAlignment wsOutput.Range("BC3"), xlLeft
    AssertHorizontalAlignment wsOutput.Range("BS3"), xlLeft
    AssertHorizontalAlignment wsOutput.Range("CR3"), xlLeft
    AssertCellNotBold wsOutput.Range("A3")
    AssertCellNotBold wsOutput.Range("C3")
    AssertCellNotBold wsOutput.Range("S3")
    AssertCellNotBold wsOutput.Range("AR3")
    AssertCellNotBold wsOutput.Range("BA3")
    AssertCellNotBold wsOutput.Range("BC3")
    AssertCellNotBold wsOutput.Range("BS3")
    AssertCellNotBold wsOutput.Range("CR3")

    AssertHeaderBlock wsOutput.Range("A3:B3"), False
    AssertHeaderBlock wsOutput.Range("C3:R3"), True
    AssertHeaderBlock wsOutput.Range("S3:AQ3"), True
    AssertHeaderBlock wsOutput.Range("AR3:AV3"), True
    AssertHeaderBlock wsOutput.Range("BA3:BB3"), False
    AssertHeaderBlock wsOutput.Range("BC3:BR3"), True
    AssertHeaderBlock wsOutput.Range("BS3:CQ3"), True
    AssertHeaderBlock wsOutput.Range("CR3:CV3"), True
    AssertBlankSeparatorRange wsOutput.Range("AW1:AZ4")

    AssertCellValue wsTableList.Range("A1"), TableIdHeaderText()
    AssertCellValue wsTableList.Range("B1"), TableNameHeaderText()
    AssertCellValue wsTableList.Range("C1"), NumberHeaderText()
    If CStr(wsTableList.Range("C2").NumberFormat) <> "@" Then
        Fail "Table-list number column should use text format."
    End If
    If CLng(wsTableList.Range("A1").Interior.Color) <> OUTPUT_FILL_COLOR Then
        Fail "Table-list header fill color is invalid."
    End If
    AssertRangeBorder wsTableList.Range("A1:C1"), xlEdgeLeft, xlContinuous
    AssertRangeBorder wsTableList.Range("A1:C1"), xlEdgeTop, xlContinuous
    AssertRangeBorder wsTableList.Range("A1:C1"), xlEdgeBottom, xlContinuous
    AssertRangeBorder wsTableList.Range("A1:C1"), xlEdgeRight, xlContinuous
    AssertRangeBorder wsTableList.Range("A1:C1"), xlInsideVertical, xlContinuous
    AssertOutputTwoCopyButton wsOutput
End Sub

'@TestMethod("SetupWorkbook")
' 変換定義A～D列の和名未取得表示がセル値の変更へ追従することを確認
Public Sub SetupWorkbook_TracksMissingNameFillColor()
    Dim wsRef As Worksheet
    Dim columnNumber As Long
    Dim missingName As String

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    missingName = "(" & W(&H548C, &H540D, &H672A, &H53D6, &H5F97) & ")"
    wsRef.Range("A2:D3").ClearContents

    For columnNumber = 1 To 4
        wsRef.Cells(2, columnNumber).Value = missingName
        wsRef.Cells(3, columnNumber).Value = "defined"
    Next columnNumber
    Application.Calculate

    For columnNumber = 1 To 4
        AssertCellDisplayFillColor _
            wsRef.Cells(2, columnNumber), RGB(252, 228, 214)
        AssertCellHasNoDisplayFill wsRef.Cells(3, columnNumber)
        wsRef.Cells(2, columnNumber).Value = "defined"
        wsRef.Cells(3, columnNumber).Value = missingName
    Next columnNumber
    Application.Calculate

    For columnNumber = 1 To 4
        AssertCellHasNoDisplayFill wsRef.Cells(2, columnNumber)
        AssertCellDisplayFillColor _
            wsRef.Cells(3, columnNumber), RGB(252, 228, 214)
    Next columnNumber
    wsRef.Range("A2:D3").ClearContents
End Sub

'@TestMethod("CopyOutput")
Public Sub CopyOutput_CopiesRenderedRange()
    Dim wsOutput As Worksheet

    SetupWorkbook
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())
    wsOutput.Cells(1, 1).Value = "copy target"
    wsOutput.Cells(2, 90).Value = "last column"
    Application.CutCopyMode = False

    CopyOutput False

    If Application.CutCopyMode <> xlCopy Then
        Fail "Output range was not copied."
    End If
    Application.CutCopyMode = False
End Sub

'@TestMethod("CopyOutputTwo")
Public Sub CopyOutputTwo_CopiesRenderedRange()
    Dim wsOutput As Worksheet
    Dim pasteSheet As Worksheet
    Dim previousDisplayAlerts As Boolean

    SetupWorkbook
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetTwoName())
    wsOutput.Cells(1, 1).Value = InputInformationTitle()
    wsOutput.Cells(4, 100).Value = "last column"
    Application.CutCopyMode = False

    CopyOutputTwo False

    If Application.CutCopyMode <> xlCopy Then
        Fail "Output-two range was not copied."
    End If
    Set pasteSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    pasteSheet.Range("A1").PasteSpecial xlPasteAll
    AssertCellValue pasteSheet.Cells(4, 100), "last column"
    Application.CutCopyMode = False
    previousDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    pasteSheet.Delete
    Application.DisplayAlerts = previousDisplayAlerts
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

'@TestMethod("AnalyzeQueries")
Public Sub AnalyzeQueries_WritesWithSubqueriesInsideOut()
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet

    ArrangeOutputOrderingFixtures
    AnalyzeQueries False

    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())
    AssertCellValue wsSql.Cells(2, COL_RESULT), ConvertFixtureSql(InputOutputNestedWithSql())
    If ExternalParserConfigured() Then
        AssertCellValue wsOutput.Cells(1, 1), SubqueryTitle("SQ1")
        AssertCellValue wsOutput.Cells(5, 1), ""
        AssertCellValue wsOutput.Cells(6, 1), SubqueryTitle("high_value_orders")
        AssertCellValue wsOutput.Cells(12, 1), ""
        AssertCellValue wsOutput.Cells(13, 1), SelectOutputTitle()
        AssertCellValue wsOutput.Cells(19, 1), ""
        AssertFormattedOutputLayout wsOutput
    Else
        AssertFallbackLines wsOutput, CStr(wsSql.Cells(2, COL_RESULT).Value), ExpectedParserNotFoundReason()
    End If
End Sub

'@TestMethod("AnalyzeQueries")
Public Sub AnalyzeQueries_PreservesLeadingApostropheInOutput()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    PutDefinition wsRef, 2, "-", "", "source_type", StatusText()
    wsSql.Cells(2, COL_SQL).Value = _
        "insert into user_summary(source_type)" & vbLf & _
        "select" & vbLf & _
        "    'BATCH'"

    AnalyzeQueries False

    AssertCellValue wsOutput.Cells(3, 17), "'BATCH'"
End Sub

'@TestMethod("AnalyzeQueries")
Public Sub AnalyzeQueries_WritesUnsupportedQueryAsIs()
    Dim wsOutput As Worksheet

    ArrangeUnsupportedOutputFixture
    AnalyzeQueries False

    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())
    AssertCellValue wsOutput.Cells(1, 1), "exec dbo.refresh_user_summary"
    AssertCellValue wsOutput.Cells(2, 1), "    @target_date = '2026-07-12'"
    AssertCellValue wsOutput.Cells(3, 1), ""
    If ExternalParserConfigured() Then
        AssertCellValue wsOutput.Cells(4, 1), ExpectedUnsupportedStatementReason()
    Else
        AssertCellValue wsOutput.Cells(4, 1), ExpectedParserNotFoundReason() & _
            ExpectedFallbackLocation(1, 2)
    End If
End Sub

'@TestMethod("AnalyzeQueries")
' 未修飾取得列を変換定義から一意に決まるSQL別名で修飾することを確認
Public Sub AnalyzeQueries_QualifiesUnqualifiedSelectColumns()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    PutDefinition wsRef, 2, "tb1", "Users", "name", "Name"
    PutDefinition wsRef, 3, "tb2", "Location", "address", "Address"
    wsSql.Cells(2, COL_SQL).Value = _
        "SELECT name, address FROM users tb1 LEFT JOIN location tb2 ON tb1.id = tb2.id"

    AnalyzeQueries False

    AssertCellValue wsOutput.Cells(3, 17), "tb1.name"
    AssertCellValue wsOutput.Cells(4, 17), "tb2.address"
End Sub

'@TestMethod("AnalyzeQueries")
' A列がハイフンの未修飾列をB列が一致する一意なSQL別名へ結び付けることを確認
Public Sub AnalyzeQueries_QualifiesStandaloneColumnThroughTableName()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim nameText As String
    Dim ageText As String

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    nameText = W(&H540D, &H524D)
    ageText = W(&H5E74, &H9F62)
    PutDefinition wsRef, 2, "tb1", UserTableText(), "age", ageText
    PutDefinition wsRef, 3, "-", UserTableText(), "name", nameText
    PutDefinition wsRef, 4, "tb2", "Location", "address", "Address"
    wsSql.Cells(2, COL_SQL).Value = _
        "SELECT name, age FROM users tb1 LEFT JOIN location tb2 ON tb1.id = tb2.id"

    AnalyzeQueries False

    AssertCellValue wsSql.Cells(2, COL_REPLACEMENT), "tb1." & nameText
    AssertCellValue wsOutput.Cells(3, 17), "tb1." & nameText
    AssertCellValue wsOutput.Cells(4, 17), "tb1.age"
End Sub

'@TestMethod("AnalyzeQueries")
' AST補完後の同じ変換内容を統合し、C列以降へ最終結果だけを書き込むことを確認
Public Sub AnalyzeQueries_WritesQualifiedReplacementValuesOnce()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim nameText As String

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    nameText = W(&H540D, &H524D)
    PutDefinition wsRef, 2, "tb1", UserTableText(), "name", nameText
    PutDefinition wsRef, 3, "-", UserTableText(), "name", nameText
    wsSql.Cells(2, COL_SQL).Value = "SELECT name, tb1.name FROM users tb1"

    AnalyzeQueries False

    AssertCellValue wsSql.Cells(2, COL_REPLACEMENT), "tb1." & nameText
    AssertCellValue wsSql.Cells(2, COL_REPLACEMENT + 1), ""
End Sub

'@TestMethod("AnalyzeQueries")
' parserの論理行番号を複数行セルと後続のSQL解析行へ正しく対応付けることを確認
Public Sub AnalyzeQueries_MapsQualifiedReplacementsToMultilineRows()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim nameText As String

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    nameText = W(&H540D, &H524D)
    PutDefinition wsRef, 2, "tb1", UserTableText(), "age", W(&H5E74, &H9F62)
    PutDefinition wsRef, 3, "tb2", UserTableText(), "code", W(&H30B3, &H30FC, &H30C9)
    PutDefinition wsRef, 4, "-", UserTableText(), "name", nameText
    wsSql.Cells(2, COL_SQL).Value = _
        "SELECT" & vbLf & _
        "    name" & vbLf & _
        "FROM users tb1;"
    wsSql.Cells(3, COL_SQL).Value = "SELECT name FROM archived_users tb2;"

    AnalyzeQueries False

    AssertCellValue wsSql.Cells(2, COL_REPLACEMENT), "tb1." & nameText
    AssertCellValue wsSql.Cells(3, COL_REPLACEMENT), "tb2." & nameText
End Sub

'@TestMethod("AnalyzeQueries")
' parserがフォールバックしてもVBA変換で得たC列内容を失わないことを確認
Public Sub AnalyzeQueries_PreservesReplacementValuesOnParserFallback()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim nameText As String

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    nameText = W(&H540D, &H524D)
    PutDefinition wsRef, 2, "tb1", UserTableText(), "name", nameText
    wsSql.Cells(2, COL_SQL).Value = "SELECT tb1.name FROM"

    AnalyzeQueries False

    AssertCellValue wsSql.Cells(2, COL_REPLACEMENT), "tb1." & nameText
    AssertCellValue wsOutput.Cells(1, 1), "SELECT tb1." & nameText & " FROM"
End Sub

'@TestMethod("AnalyzeQueries")
' 構文エラーの論理行に対応するSQL解析シートの入力セルへフォーカスすることを確認
Public Sub AnalyzeQueries_FocusesSyntaxFallbackSqlRow()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    wsSql.Cells(2, COL_SQL).Value = "SELECT 1;"
    wsSql.Cells(3, COL_SQL).Value = "SELECT id FROM users WHERE;"
    wsOutput.Activate

    AnalyzeQueries False

    If Not ActiveSheet Is wsSql Then
        Fail "Syntax fallback should activate the SQL analysis sheet."
    End If
    If ActiveCell.Address <> wsSql.Cells(3, COL_SQL).Address Then
        Fail "Syntax fallback should select the SQL input cell containing the error."
    End If
End Sub

'@TestMethod("AnalyzeQueries")
' 列指定なしINSERT SELECTを移送先テーブル側の和名で出力することを確認
Public Sub AnalyzeQueries_RendersInsertSelectWithoutColumnList()
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet
    Dim archiveIdText As String
    Dim userIdText As String

    If Not ExternalParserConfigured() Then Exit Sub

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    wsRef.Range("A2:D200").ClearContents
    wsSql.Range("A2:Z200").ClearContents
    wsOutput.Cells.ClearContents
    archiveIdText = W(&H30A2, &H30FC, &H30AB, &H30A4, &H30D6) & "ID"
    userIdText = UserTableText() & "ID"
    PutDefinition wsRef, 2, "user_archive", _
        W(&H30E6, &H30FC, &H30B6, &H30FC, &H5C65, &H6B74), "id", archiveIdText
    PutDefinition wsRef, 3, "tb1", UserTableText(), "id", userIdText
    wsSql.Cells(2, COL_SQL).Value = _
        "INSERT INTO user_archive SELECT tb1.id FROM users AS tb1;"

    AnalyzeQueries False

    AssertCellValue wsOutput.Cells(5, 1), DataTransferTitle()
    AssertCellValue wsOutput.Cells(8, 1), archiveIdText
    AssertCellValue wsOutput.Cells(8, 19), "tb1." & userIdText
End Sub

'@TestMethod("ClearData")
' クリア確認が解析結果を対象として案内することを確認
Public Sub ClearConfirmMessage_UsesAnalysisResultWording()
    Dim expected As String

    expected = W(&H89E3, &H6790, &H7D50, &H679C, &H3092, _
        &H30AF, &H30EA, &H30A2, &H3057, &H307E, &H3059, &H3002, _
        &H3088, &H308D, &H3057, &H3044, &H3067, &H3059, &H304B, &HFF1F)
    If ClearConfirmMessage() <> expected Then
        Fail "Clear confirmation message does not describe analysis results."
    End If
End Sub

'@TestMethod("ClearData")
Public Sub ClearData_ClearsOutputSheet()
    Dim activeSheetBeforeClear As Object
    Dim wsRef As Worksheet
    Dim wsSql As Worksheet
    Dim wsOutput As Worksheet

    SetupWorkbook
    Set wsRef = ThisWorkbook.Worksheets(ReferenceSheetName())
    Set wsSql = ThisWorkbook.Worksheets(SqlSheetName())
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())
    wsOutput.Cells(1, 1).Value = "output header"
    wsOutput.Cells(3, 2).Value = "output detail"
    wsOutput.Cells(4, 90).Value = "last column output"
    wsOutput.Cells(5, 1).Formula = "=""formula output"""
    wsOutput.Cells(200, 17).WrapText = True
    wsOutput.Cells(200, 17).ShrinkToFit = True

    ' クリア前の各シートの表示位置とアクティブシートを再現
    wsRef.Activate
    wsRef.Cells(80, 4).Select
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollColumn = 2
    wsSql.Activate
    wsSql.Cells(90, 20).Select
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollColumn = 10
    wsOutput.Activate
    wsOutput.Cells(100, 90).Select
    ActiveWindow.ScrollRow = 50
    ActiveWindow.ScrollColumn = 40
    wsSql.Activate
    Set activeSheetBeforeClear = ActiveSheet

    ClearData False

    AssertCellValue wsOutput.Cells(1, 1), ""
    AssertCellValue wsOutput.Cells(3, 2), ""
    AssertCellValue wsOutput.Cells(4, 90), ""
    AssertCellValue wsOutput.Cells(5, 1), ""
    AssertOutputSheetFont
    AssertOutputTextFittingDisabled wsOutput
    If Not ActiveSheet Is activeSheetBeforeClear Then
        Fail "Active sheet changed while clearing data."
    End If
    AssertSheetViewReset wsRef
    AssertSheetViewReset wsSql
    AssertSheetViewReset wsOutput
End Sub

'@TestMethod("ClearData")
Public Sub ClearData_InitializesOutputTwoAndPreservesTableList()
    Dim activeSheetBeforeClear As Object
    Dim wsOutput As Worksheet
    Dim wsTableList As Worksheet

    SetupWorkbook
    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetTwoName())
    Set wsTableList = ThisWorkbook.Worksheets(TableListSheetName())
    wsTableList.Range("A2:C200").ClearContents
    PutTableListRow wsTableList, 2, "users", UserTableText(), "1-1"
    wsOutput.Range("A4").Value = 99
    wsOutput.Range("C4").Value = "stale input"
    wsOutput.Range("BA4").Value = 99
    wsOutput.Range("BC4").Value = "stale output"

    wsTableList.Activate
    wsTableList.Cells(80, 3).Select
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollColumn = 2
    wsOutput.Activate
    wsOutput.Cells(100, 100).Select
    ActiveWindow.ScrollRow = 50
    ActiveWindow.ScrollColumn = 40
    Set activeSheetBeforeClear = ActiveSheet

    ClearData False

    AssertCellValue wsTableList.Range("A2"), "users"
    AssertCellValue wsTableList.Range("B2"), UserTableText()
    AssertCellValue wsTableList.Range("C2"), "1-1"
    AssertCellValue wsOutput.Range("A1"), InputInformationTitle()
    AssertCellValue wsOutput.Range("BA1"), OutputInformationTitle()
    AssertCellValue wsOutput.Range("A3"), "No"
    AssertCellValue wsOutput.Range("C3"), TableIdHeaderText()
    AssertCellValue wsOutput.Range("BA3"), "No"
    AssertCellValue wsOutput.Range("BC3"), TableIdHeaderText()
    AssertCellValue wsOutput.Range("A4"), ""
    AssertCellValue wsOutput.Range("C4"), ""
    AssertCellValue wsOutput.Range("BA4"), ""
    AssertCellValue wsOutput.Range("BC4"), ""
    AssertBlankSeparatorRange wsOutput.Range("AW1:AZ4")
    If Not ActiveSheet Is activeSheetBeforeClear Then
        Fail "Active sheet changed while clearing output two."
    End If
    AssertSheetViewReset wsOutput
    AssertSheetViewReset wsTableList
End Sub

' アウトプット順序テスト用データを作成
Private Sub ArrangeOutputOrderingFixtures()
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
    wsSql.Cells(2, COL_SQL).Value = InputOutputNestedWithSql()
End Sub

' 未対応クエリのアウトプットテスト用データを作成
Private Sub ArrangeUnsupportedOutputFixture()
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
    wsSql.Cells(2, COL_SQL).Value = InputUnsupportedSql()
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

' WITHとネストしたサブクエリを含むアウトプット順序確認用SQLを返す
Private Function InputOutputNestedWithSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, TS("with high_value_orders as (")
    AppendA5M2Line resultText, "    select"
    AppendA5M2Line resultText, "        orders.user_id"
    AppendA5M2Line resultText, "        , orders.amount"
    AppendA5M2Line resultText, "    from"
    AppendA5M2Line resultText, TS("        orders")
    AppendA5M2Line resultText, "    where"
    AppendA5M2Line resultText, "        orders.amount > 1000"
    AppendA5M2Line resultText, TS("        and exists (")
    AppendA5M2Line resultText, "            select"
    AppendA5M2Line resultText, TS("                1")
    AppendA5M2Line resultText, "            from"
    AppendA5M2Line resultText, TS("                order_items")
    AppendA5M2Line resultText, "            where"
    AppendA5M2Line resultText, "                order_items.order_id = orders.order_id"
    AppendA5M2Line resultText, "        )"
    AppendA5M2Line resultText, TS(")")
    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "    users.user_id"
    AppendA5M2Line resultText, "    , high_value_orders.amount"
    AppendA5M2Line resultText, "from"
    AppendA5M2Line resultText, TS("    users")
    AppendA5M2Line resultText, TS("    inner join high_value_orders")
    AppendA5M2Line resultText, "        on high_value_orders.user_id = users.user_id"

    InputOutputNestedWithSql = FinishA5M2Sql(resultText)
End Function

' ネスト内側のサブクエリ期待値用SQLを返す
Private Function InputOutputNestedInnerSubquerySql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, TS("                1")
    AppendA5M2Line resultText, "            from"
    AppendA5M2Line resultText, TS("                order_items")
    AppendA5M2Line resultText, "            where"
    AppendA5M2Line resultText, "                order_items.order_id = orders.order_id"

    InputOutputNestedInnerSubquerySql = FinishA5M2Sql(resultText)
End Function

' ネスト外側のサブクエリ期待値用SQLを返す
Private Function InputOutputNestedOuterSubquerySql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "select"
    AppendA5M2Line resultText, "        orders.user_id"
    AppendA5M2Line resultText, "        , orders.amount"
    AppendA5M2Line resultText, "    from"
    AppendA5M2Line resultText, TS("        orders")
    AppendA5M2Line resultText, "    where"
    AppendA5M2Line resultText, "        orders.amount > 1000"
    AppendA5M2Line resultText, TS("        and exists (")
    AppendA5M2Line resultText, "            select"
    AppendA5M2Line resultText, TS("                1")
    AppendA5M2Line resultText, "            from"
    AppendA5M2Line resultText, TS("                order_items")
    AppendA5M2Line resultText, "            where"
    AppendA5M2Line resultText, "                order_items.order_id = orders.order_id"
    AppendA5M2Line resultText, "        )"

    InputOutputNestedOuterSubquerySql = FinishA5M2Sql(resultText)
End Function

' 未対応クエリのアウトプット期待値用SQLを返す
Private Function InputUnsupportedSql() As String
    Dim resultText As String

    AppendA5M2Line resultText, "exec dbo.refresh_user_summary"
    AppendA5M2Line resultText, "    @target_date = '2026-07-12'"

    InputUnsupportedSql = FinishA5M2Sql(resultText)
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

' 内側サブクエリのアウトプット期待値を返す
Private Function ExpectedOutputNestedInnerSubquerySql() As String
    ExpectedOutputNestedInnerSubquerySql = TrimFixtureSql(ConvertFixtureSql(InputOutputNestedInnerSubquerySql()))
End Function

' 外側サブクエリのアウトプット期待値を返す
Private Function ExpectedOutputNestedOuterSubquerySql() As String
    ExpectedOutputNestedOuterSubquerySql = TrimFixtureSql(ConvertFixtureSql(InputOutputNestedOuterSubquerySql()))
End Function

' WITHを含むクエリ全体のアウトプット期待値を返す
Private Function ExpectedOutputNestedWholeSql() As String
    ExpectedOutputNestedWholeSql = TrimFixtureSql(ConvertFixtureSql(InputOutputNestedWithSql()))
End Function

' 未対応クエリのアウトプット期待値を返す
Private Function ExpectedOutputUnsupportedSql() As String
    ExpectedOutputUnsupportedSql = TrimFixtureSql(ConvertFixtureSql(InputUnsupportedSql()))
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

' アウトプット期待値の前後空白を除去
Private Function TrimFixtureSql(ByVal sourceText As String) As String
    Dim startIndex As Long
    Dim endIndex As Long

    startIndex = 1
    Do While startIndex <= Len(sourceText)
        If Not IsFixtureWhitespace(Mid$(sourceText, startIndex, 1)) Then Exit Do
        startIndex = startIndex + 1
    Loop

    endIndex = Len(sourceText)
    Do While endIndex >= startIndex
        If Not IsFixtureWhitespace(Mid$(sourceText, endIndex, 1)) Then Exit Do
        endIndex = endIndex - 1
    Loop

    If endIndex >= startIndex Then
        TrimFixtureSql = Mid$(sourceText, startIndex, endIndex - startIndex + 1)
    End If
End Function

' テストSQLの除去対象空白か判定
Private Function IsFixtureWhitespace(ByVal value As String) As Boolean
    IsFixtureWhitespace = (value = " " Or value = vbTab Or value = vbCr Or value = vbLf)
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

' 変換定義のテストデータを1行設定
Private Sub PutDefinition(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal tableId As String, ByVal tableName As String, ByVal fieldId As String, ByVal fieldName As String)
    ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, 4)).NumberFormat = "@"
    ws.Cells(rowNumber, 1).Value = tableId
    ws.Cells(rowNumber, 2).Value = tableName
    ws.Cells(rowNumber, 3).Value = fieldId
    ws.Cells(rowNumber, 4).Value = fieldName
End Sub

' テーブル一覧へ1行分のマスターを設定
Private Sub PutTableListRow(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal tableId As String, ByVal tableName As String, ByVal tableNumber As String)
    ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, 3)).NumberFormat = "@"
    ws.Cells(rowNumber, 1).Value = tableId
    ws.Cells(rowNumber, 2).Value = tableName
    ws.Cells(rowNumber, 3).Value = tableNumber
End Sub

' アウトプット②の1明細を検証
Private Sub AssertOutputTwoRow( _
    ByVal ws As Worksheet, _
    ByVal rowNumber As Long, _
    ByVal startColumn As Long, _
    ByVal expectedTableId As String, _
    ByVal expectedTableName As String, _
    ByVal expectedTableNumber As String)

    AssertCellValue ws.Cells(rowNumber, startColumn), CStr(rowNumber - 3)
    AssertCellValue ws.Cells(rowNumber, startColumn + 2), expectedTableId
    AssertCellValue ws.Cells(rowNumber, startColumn + 18), expectedTableName
    AssertCellValue ws.Cells(rowNumber, startColumn + 43), expectedTableNumber
End Sub

' SQL解析行の変換結果と変換内容を検証
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

' 指定位置のシート名を検証
Private Sub AssertWorksheetNameAt(ByVal sheetIndex As Long, ByVal expectedName As String)
    Dim actualName As String

    actualName = ThisWorkbook.Worksheets(sheetIndex).Name
    If actualName <> expectedName Then
        Fail "Worksheet " & CStr(sheetIndex) & " expected=[" & expectedName & _
            "] actual=[" & actualName & "]"
    End If
End Sub

' 指定シートが存在しないことを検証
Private Sub AssertWorksheetDoesNotExist(ByVal sheetName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Fail "Unexpected worksheet found: " & sheetName
    End If
End Sub

' 指定セルの結合範囲を検証
Private Sub AssertMergedArea(ByVal cell As Range, ByVal expectedAddress As String)
    If Not CBool(cell.MergeCells) Then
        Fail cell.Worksheet.Name & "!" & cell.Address(False, False) & _
            " should be merged."
    End If
    If cell.MergeArea.Address <> expectedAddress Then
        Fail cell.Worksheet.Name & "!" & cell.Address(False, False) & _
            " merge expected=[" & expectedAddress & "] actual=[" & _
            cell.MergeArea.Address & "]"
    End If
End Sub

' 指定セルが結合されていないことを検証
Private Sub AssertCellNotMerged(ByVal cell As Range)
    If CBool(cell.MergeCells) Then
        Fail cell.Worksheet.Name & "!" & cell.Address(False, False) & _
            " should not be merged."
    End If
End Sub

' 指定セルの横位置を検証
Private Sub AssertHorizontalAlignment(ByVal cell As Range, ByVal expectedAlignment As Long)
    If CLng(cell.HorizontalAlignment) <> expectedAlignment Then
        Fail cell.Worksheet.Name & "!" & cell.Address(False, False) & _
            " horizontal alignment expected=[" & CStr(expectedAlignment) & _
            "] actual=[" & CStr(cell.HorizontalAlignment) & "]"
    End If
End Sub

' ヘッダーの論理項目外枠、行境界、塗りつぶしを検証
Private Sub AssertHeaderBlock(ByVal blockRange As Range, ByVal assertNoInsideVertical As Boolean)
    AssertRangeBorder blockRange, xlEdgeLeft, xlContinuous
    AssertRangeBorder blockRange, xlEdgeTop, xlContinuous
    AssertRangeBorder blockRange, xlEdgeBottom, xlContinuous
    AssertRangeBorder blockRange, xlEdgeRight, xlContinuous
    If assertNoInsideVertical Then
        AssertRangeBorder blockRange, xlInsideVertical, xlNone
    End If
    If CLng(blockRange.Cells(1, 1).Interior.Color) <> OUTPUT_FILL_COLOR Then
        Fail blockRange.Worksheet.Name & "!" & blockRange.Address(False, False) & _
            " header fill color is invalid."
    End If
    If CLng(blockRange.Cells(1, blockRange.Columns.Count).Interior.Color) <> OUTPUT_FILL_COLOR Then
        Fail blockRange.Worksheet.Name & "!" & blockRange.Address(False, False) & _
            " header fill does not span the logical item."
    End If
End Sub

' データ行の論理項目外枠と行境界を検証
Private Sub AssertDataBlock(ByVal blockRange As Range, Optional ByVal assertNoInsideVertical As Boolean = True)
    AssertRangeBorder blockRange, xlEdgeLeft, xlContinuous
    AssertRangeBorder blockRange, xlEdgeTop, xlContinuous
    AssertRangeBorder blockRange, xlEdgeBottom, xlContinuous
    AssertRangeBorder blockRange, xlEdgeRight, xlContinuous
    If assertNoInsideVertical Then
        AssertRangeBorder blockRange, xlInsideVertical, xlNone
    End If
    AssertCellHasNoDisplayFill blockRange.Cells(1, 1)
End Sub

' 罫線種別を検証
Private Sub AssertRangeBorder(ByVal targetRange As Range, ByVal borderIndex As Long, ByVal expectedLineStyle As Long)
    Dim actualLineStyle As Variant

    actualLineStyle = targetRange.Borders(borderIndex).LineStyle
    If IsNull(actualLineStyle) Then
        Fail targetRange.Worksheet.Name & "!" & targetRange.Address(False, False) & _
            " has mixed border styles."
    End If
    If CLng(actualLineStyle) <> expectedLineStyle Then
        Fail targetRange.Worksheet.Name & "!" & targetRange.Address(False, False) & _
            " border expected=[" & CStr(expectedLineStyle) & "] actual=[" & _
            CStr(actualLineStyle) & "]"
    End If
End Sub

' 入力情報と出力情報の間の4列が空白であることを検証
Private Sub AssertBlankSeparatorRange(ByVal separatorRange As Range)
    Dim cell As Range

    If Application.WorksheetFunction.CountA(separatorRange) <> 0 Then
        Fail separatorRange.Worksheet.Name & "!" & separatorRange.Address(False, False) & _
            " separator columns should be empty."
    End If
    For Each cell In separatorRange.Cells
        If CLng(cell.DisplayFormat.Interior.Pattern) <> xlPatternNone Then
            Fail cell.Worksheet.Name & "!" & cell.Address(False, False) & _
                " separator cell should not have a fill."
        End If
    Next cell
    ' Check an interior separator cell so neighboring block outer borders do not interfere.
    AssertCellHasNoEdgeBorders separatorRange.Cells(1, 2)
    If separatorRange.Rows.Count >= 3 Then
        AssertCellHasNoEdgeBorders separatorRange.Cells(3, 2)
    End If
End Sub

' 指定シートが存在することを検証
Private Sub AssertWorksheetExists(ByVal sheetName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Fail "Worksheet not found: " & sheetName
    End If
End Sub

' アウトプットシートの目盛り線非表示を検証
Private Sub AssertOutputSheetGridlinesHidden()
    Dim wsOutput As Worksheet

    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())
    AssertSheetGridlinesHidden wsOutput
End Sub

' 指定シートの目盛り線非表示を検証
Private Sub AssertSheetGridlinesHidden(ByVal ws As Worksheet)
    Dim previousSheet As Object

    ' 目盛り線の表示状態は対象シートを表示して確認
    Set previousSheet = ActiveSheet
    ws.Activate
    If ActiveWindow.DisplayGridlines Then
        Fail ws.Name & " gridlines should be hidden."
    End If
    previousSheet.Activate
End Sub

' アウトプットシートのフォントを検証
Private Sub AssertOutputSheetFont()
    Dim wsOutput As Worksheet

    Set wsOutput = ThisWorkbook.Worksheets(OutputSheetName())

    AssertCellFont wsOutput.Cells(1, 1), OutputFontName(), 9
    AssertCellFont wsOutput.Cells(20, 5), OutputFontName(), 9
End Sub

' アウトプット範囲で折り返しと縮小表示が無効なことを検証
Private Sub AssertOutputTextFittingDisabled(ByVal wsOutput As Worksheet)
    If CBool(wsOutput.Cells(3, 17).WrapText) Then
        Fail "Output wrap text should be disabled."
    End If
    If CBool(wsOutput.Cells(3, 17).ShrinkToFit) Then
        Fail "Output shrink to fit should be disabled."
    End If
    If CBool(wsOutput.Cells(200, 17).WrapText) Then
        Fail "Unused output cells should not retain wrap text."
    End If
    If CBool(wsOutput.Cells(200, 17).ShrinkToFit) Then
        Fail "Unused output cells should not retain shrink to fit."
    End If
End Sub

' 指定シートの選択セルとスクロール位置がA1へ戻ることを検証
Private Sub AssertSheetViewReset(ByVal ws As Worksheet)
    Dim previousSheet As Object

    Set previousSheet = ActiveSheet
    ws.Activate
    If Selection.Address <> "$A$1" Then
        Fail ws.Name & " selection should be A1 after clearing."
    End If
    If ActiveWindow.ScrollRow <> 1 Or ActiveWindow.ScrollColumn <> 1 Then
        Fail ws.Name & " scroll position should be A1 after clearing."
    End If
    previousSheet.Activate
End Sub

' アウトプットシートのコピーボタンを検証
Private Sub AssertOutputCopyButton(ByVal wsOutput As Worksheet)
    Dim copyButton As Shape

    On Error Resume Next
    Set copyButton = wsOutput.Shapes("btnCopyOutput")
    On Error GoTo 0
    If copyButton Is Nothing Then
        Fail "Output copy button was not found."
    End If
    If Right$(CStr(copyButton.OnAction), Len("CopyOutput")) <> "CopyOutput" Then
        Fail "Output copy button action is invalid."
    End If
End Sub

' アウトプット②のコピーボタンを検証
Private Sub AssertOutputTwoCopyButton(ByVal wsOutput As Worksheet)
    Dim copyButton As Shape

    For Each copyButton In wsOutput.Shapes
        If Right$(CStr(copyButton.OnAction), Len("CopyOutputTwo")) = _
            "CopyOutputTwo" Then
            Exit Sub
        End If
    Next copyButton
    Fail "Output-two copy button was not found."
End Sub

' SQL解析シートの操作ボタンが表示可能な固定配置であることを検証
Private Sub AssertSqlActionButtons(ByVal wsSql As Worksheet)
    If ThisWorkbook.DisplayDrawingObjects <> xlDisplayShapes Then
        Fail "Workbook drawing objects should be displayed."
    End If

    AssertSqlActionButton _
        wsSql, "btnAnalyzeQueries", "AnalyzeQueries", wsSql.Columns("E").Left
    AssertSqlActionButton _
        wsSql, "btnClearData", "ClearData", wsSql.Columns("E").Left + 82
End Sub

' 指定したSQL解析シートボタンの表示、配置、操作を検証
Private Sub AssertSqlActionButton( _
    ByVal wsSql As Worksheet, _
    ByVal buttonName As String, _
    ByVal macroName As String, _
    ByVal expectedLeft As Double)

    Dim actionButton As Shape

    On Error Resume Next
    Set actionButton = wsSql.Shapes(buttonName)
    On Error GoTo 0
    If actionButton Is Nothing Then
        Fail "SQL action button was not found: " & buttonName
    End If
    If actionButton.Visible <> msoTrue Then
        Fail buttonName & " should be visible."
    End If
    If actionButton.Placement <> xlFreeFloating Then
        Fail buttonName & " should use free-floating placement."
    End If
    If Abs(CDbl(actionButton.Top) - (CDbl(wsSql.Rows(1).Top) + 2)) > 0.1 Then
        Fail buttonName & " top position is invalid."
    End If
    If Abs(CDbl(actionButton.Left) - expectedLeft) > 0.1 Then
        Fail buttonName & " left position is invalid."
    End If
    If Abs(CDbl(actionButton.Width) - 72) > 0.1 Or _
        Abs(CDbl(actionButton.Height) - 24) > 0.1 Then
        Fail buttonName & " size is invalid."
    End If
    If Right$(CStr(actionButton.OnAction), Len(macroName)) <> macroName Then
        Fail buttonName & " action is invalid."
    End If
End Sub

' フォールバックSQLの行と原因を検証
Private Sub AssertFallbackLines(ByVal wsOutput As Worksheet, ByVal queryText As String, ByVal expectedReason As String)
    Dim lines As Variant
    Dim normalizedText As String
    Dim lineIndex As Long
    Dim reasonRow As Long

    normalizedText = Replace(queryText, vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)
    Do While Len(normalizedText) > 0 And Left$(normalizedText, 1) = vbLf
        normalizedText = Mid$(normalizedText, 2)
    Loop
    Do While Len(normalizedText) > 0 And Right$(normalizedText, 1) = vbLf
        normalizedText = Left$(normalizedText, Len(normalizedText) - 1)
    Loop
    lines = Split(normalizedText, vbLf)
    For lineIndex = LBound(lines) To UBound(lines)
        AssertCellValue wsOutput.Cells(lineIndex - LBound(lines) + 1, 1), CStr(lines(lineIndex))
    Next lineIndex

    reasonRow = UBound(lines) - LBound(lines) + 3
    AssertCellValue wsOutput.Cells(reasonRow - 1, 1), ""
    AssertCellValue wsOutput.Cells(reasonRow, 1), expectedReason & _
        ExpectedFallbackLocation(1, UBound(lines) - LBound(lines) + 1)
End Sub

' parser出力の共通書式を確認
Private Sub AssertFormattedOutputLayout(ByVal wsOutput As Worksheet)
    If CLng(wsOutput.Cells(3, 1).Interior.Color) <> OUTPUT_FILL_COLOR Then
        Fail "Output label fill color is invalid."
    End If
    AssertCellHasNoEdgeBorders wsOutput.Cells(1, 1)
    AssertCellHasNoEdgeBorders wsOutput.Cells(1, 90)
    AssertCellHasNoEdgeBorders wsOutput.Cells(2, 1), False
    AssertCellHasNoEdgeBorders wsOutput.Cells(2, 90), False
    If wsOutput.Range("A3:CL3").Borders(xlEdgeTop).LineStyle <> xlContinuous Then
        Fail "Output table border is missing."
    End If
    If CDbl(wsOutput.Rows(1).RowHeight) <> 13.5 Then
        Fail "Output row height should be 13.5."
    End If
    If Abs(CDbl(wsOutput.Columns(1).ColumnWidth) - 1.14) > 0.02 Then
        Fail "Output column width should be 1.14."
    End If
    AssertOutputSheetFont
    AssertOutputSheetGridlinesHidden
End Sub

' 指定セルの辺に罫線がないことを検証
Private Sub AssertCellHasNoEdgeBorders( _
    ByVal cell As Range, _
    Optional ByVal includeBottom As Boolean = True)

    Dim borderIndex As Variant
    Dim borderIndexes As Variant

    If includeBottom Then
        borderIndexes = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
    Else
        ' 表本体の上辺は、隣接する参照テーブル行の下辺としても取得される
        borderIndexes = Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight)
    End If

    For Each borderIndex In borderIndexes
        If cell.Borders(CLng(borderIndex)).LineStyle <> xlNone Then
            Fail cell.Worksheet.Name & "!" & cell.Address(False, False) & _
                " should not have borders."
        End If
    Next borderIndex
End Sub

' 外部parserを使用するテスト実行か判定
Private Function ExternalParserConfigured() As Boolean
    ExternalParserConfigured = (Len(Environ$("SQL_ANALYSIS_FORMATTER_PARSER_EXE")) > 0)
End Function

' セルのフォント名とサイズを検証
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

' セルが太字でないことを検証
Private Sub AssertCellNotBold(ByVal cell As Range)
    If CBool(cell.Font.Bold) Then
        Fail cell.Worksheet.Name & "!" & cell.Address(False, False) & _
            " should not be bold."
    End If
End Sub

' セル値を検証
Private Sub AssertCellValue(ByVal cell As Range, ByVal expected As String)
    Dim actual As String

    actual = CStr(cell.Value)
    If actual <> expected Then
        Fail cell.Worksheet.Name & "!" & cell.Address(False, False) & _
            " expected=[" & expected & "] actual=[" & actual & "]"
    End If
End Sub

' 条件付き書式を反映したセルの塗りつぶし色を検証
Private Sub AssertCellDisplayFillColor(ByVal cell As Range, ByVal expectedColor As Long)
    If CLng(cell.DisplayFormat.Interior.Color) <> expectedColor Then
        Fail cell.Worksheet.Name & "!" & cell.Address(False, False) & _
            " fill expected=[" & CStr(expectedColor) & "] actual=[" & _
            CStr(cell.DisplayFormat.Interior.Color) & "]"
    End If
End Sub

' 条件付き書式を反映したセルに塗りつぶしがないことを検証
Private Sub AssertCellHasNoDisplayFill(ByVal cell As Range)
    If CLng(cell.DisplayFormat.Interior.Pattern) <> xlPatternNone Then
        Fail cell.Worksheet.Name & "!" & cell.Address(False, False) & _
            " fill pattern expected=[none] actual=[" & _
            CStr(cell.DisplayFormat.Interior.Pattern) & "]"
    End If
End Sub

' テスト失敗を例外として通知
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

' 変換定義シート名を返す
Private Function ReferenceSheetName() As String
    ReferenceSheetName = W(&H5909, &H63DB, &H5B9A, &H7FA9)
End Function

' SQL解析シート名を返す
Private Function SqlSheetName() As String
    SqlSheetName = "SQL" & W(&H89E3, &H6790)
End Function

' アウトプットシート名を返す
Private Function OutputSheetName() As String
    OutputSheetName = LegacyOutputSheetName() & W(&H2460)
End Function

' アウトプット②シート名を返す
Private Function OutputSheetTwoName() As String
    OutputSheetTwoName = LegacyOutputSheetName() & W(&H2461)
End Function

' テーブル一覧シート名を返す
Private Function TableListSheetName() As String
    TableListSheetName = W(&H30C6, &H30FC, &H30D6, &H30EB, &H4E00, &H89A7)
End Function

' 旧アウトプットシート名を返す
Private Function LegacyOutputSheetName() As String
    LegacyOutputSheetName = W(&H30A2, &H30A6, &H30C8, &H30D7, &H30C3, &H30C8)
End Function

' アウトプット用フォント名を返す
Private Function OutputFontName() As String
    OutputFontName = W(&HFF2D, &HFF33, &H20, &H30B4, &H30B7, &H30C3, &H30AF)
End Function

' アウトプット②入力情報タイトルを返す
Private Function InputInformationTitle() As String
    InputInformationTitle = W(&HFF1C, &H5165, &H529B, &H60C5, &H5831, &HFF1E)
End Function

' アウトプット②出力情報タイトルを返す
Private Function OutputInformationTitle() As String
    OutputInformationTitle = W(&HFF1C, &H51FA, &H529B, &H60C5, &H5831, &HFF1E)
End Function

' テーブルIDヘッダーを返す
Private Function TableIdHeaderText() As String
    TableIdHeaderText = W(&H30C6, &H30FC, &H30D6, &H30EB) & "ID"
End Function

' テーブル名称ヘッダーを返す
Private Function TableNameHeaderText() As String
    TableNameHeaderText = W(&H30C6, &H30FC, &H30D6, &H30EB, &H540D, &H79F0)
End Function

' 番号ヘッダーを返す
Private Function NumberHeaderText() As String
    NumberHeaderText = W(&H756A, &H53F7)
End Function

' サブクエリ表のタイトルを返す
Private Function SubqueryTitle(ByVal queryName As String) As String
    SubqueryTitle = W(&H30B5, &H30D6, &H30AF, &H30A8, &H30EA) & "[" & queryName & "]"
End Function

' SELECT系表のタイトルを返す
Private Function SelectOutputTitle() As String
    SelectOutputTitle = W(&HFF1C) & "DB" & W(&H5165, &H51FA, &H529B, &H9805, &H76EE, &H5B9A, &H7FA9, &HFF1E)
End Function

' データ移送表のタイトルを生成
Private Function DataTransferTitle() As String
    DataTransferTitle = W(&HFF1C, &H30C7, &H30FC, &H30BF, &H79FB, &H9001, &H8868, &HFF1E)
End Function

' 未対応ステートメントのフォールバック原因を返す
Private Function ExpectedUnsupportedStatementReason() As String
    ExpectedUnsupportedStatementReason = _
        W(&H30D5, &H30A9, &H30FC, &H30EB, &H30D0, &H30C3, &H30AF, &H539F, &H56E0) & ": " & _
        W(&H672A, &H5BFE, &H5FDC, &H306E, &H30B9, &H30C6, &H30FC, &H30C8, &H30E1, &H30F3, &H30C8) & ": EXECUTE" & _
        ExpectedFallbackLocation(1, 2)
End Function

' parser未配置時のフォールバック原因を返す
Private Function ExpectedParserNotFoundReason() As String
    ExpectedParserNotFoundReason = _
        W(&H30D5, &H30A9, &H30FC, &H30EB, &H30D0, &H30C3, &H30AF, &H539F, &H56E0) & ": parser EXE" & _
        W(&H304C, &H898B, &H3064, &H304B, &H308A, &H307E, &H305B, &H3093, &H3002)
End Function

' フォールバック対象クエリの行表示を返す
Private Function ExpectedFallbackLocation(ByVal startRow As Long, ByVal endRow As Long) As String
    Dim rowText As String

    rowText = CStr(startRow)
    If endRow <> startRow Then
        rowText = rowText & W(&HFF5E) & CStr(endRow)
    End If
    ExpectedFallbackLocation = W(&HFF08, &H5BFE, &H8C61, &H30AF, &H30A8, &H30EA, &H3A, &H20) & _
        OutputSheetName() & " " & rowText & W(&H884C, &H76EE, &HFF09)
End Function

' ユーザーテーブル和名を返す
Private Function UserTableText() As String
    UserTableText = W(&H30E6, &H30FC, &H30B6, &H30FC)
End Function

' 参照テーブル見出しを返す
Private Function ReferenceTablesText() As String
    ReferenceTablesText = W(&H53C2, &H7167, &H30C6, &H30FC, &H30D6, &H30EB)
End Function

' 注文テーブル和名を返す
Private Function OrderTableText() As String
    OrderTableText = W(&H6CE8, &H6587)
End Function

' 注文明細テーブル和名を返す
Private Function OrderItemTableText() As String
    OrderItemTableText = W(&H6CE8, &H6587, &H660E, &H7D30)
End Function

' 管理者テーブル和名を返す
Private Function ManagerTableText() As String
    ManagerTableText = W(&H7BA1, &H7406, &H8005)
End Function

' ユーザーID和名を返す
Private Function UserIdText() As String
    UserIdText = UserTableText() & "ID"
End Function

' 氏名和名を返す
Private Function FullNameText() As String
    FullNameText = W(&H6C0F, &H540D)
End Function

' メール和名を返す
Private Function MailText() As String
    MailText = W(&H30E1, &H30FC, &H30EB)
End Function

' 管理者ID和名を返す
Private Function ManagerIdText() As String
    ManagerIdText = ManagerTableText() & "ID"
End Function

' 注文ID和名を返す
Private Function OrderIdText() As String
    OrderIdText = OrderTableText() & "ID"
End Function

' 注文ユーザーID和名を返す
Private Function OrderUserIdText() As String
    OrderUserIdText = OrderTableText() & UserTableText() & "ID"
End Function

' 金額和名を返す
Private Function AmountText() As String
    AmountText = W(&H91D1, &H984D)
End Function

' 明細注文ID和名を返す
Private Function DetailOrderIdText() As String
    DetailOrderIdText = W(&H660E, &H7D30, &H6CE8, &H6587) & "ID"
End Function

' 商品ID和名を返す
Private Function ProductIdText() As String
    ProductIdText = W(&H5546, &H54C1) & "ID"
End Function

' 数量和名を返す
Private Function QuantityText() As String
    QuantityText = W(&H6570, &H91CF)
End Function

' 状態和名を返す
Private Function StatusText() As String
    StatusText = W(&H72B6, &H614B)
End Function

' 作成日時和名を返す
Private Function CreatedAtText() As String
    CreatedAtText = W(&H4F5C, &H6210, &H65E5, &H6642)
End Function

' 更新日時和名を返す
Private Function UpdatedAtText() As String
    UpdatedAtText = W(&H66F4, &H65B0, &H65E5, &H6642)
End Function
