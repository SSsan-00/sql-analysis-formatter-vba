Attribute VB_Name = "SqlAnalysisFormatterGoldenTests"
Option Explicit

'@TestModule
'@Folder("Tests")

Private Const NO_BORDER As Long = -4142
Private Const VALUE_TOLERANCE As Double = 0.01

' 期待値ブックと実出力の書式をExcel内部で一括比較
Public Function CompareOutputGoldenFormat( _
    ByVal caseId As String, _
    ByVal expectationWorkbookName As String, _
    ByVal expectationSheetName As String, _
    ByVal actualSheetName As String, _
    ByVal expectedRowCount As Long, _
    ByVal compareColumnWidths As Boolean) As String

    Dim expectedSheet As Worksheet
    Dim actualSheet As Worksheet
    Dim formatColumns As Variant
    Dim rowNumber As Long
    Dim columnNumber As Variant
    Dim mismatch As String

    On Error GoTo CompareError

    Set expectedSheet = Application.Workbooks(expectationWorkbookName).Worksheets(expectationSheetName)
    Set actualSheet = ThisWorkbook.Worksheets(actualSheetName)
    formatColumns = Array(1, 6, 7, 15, 17, 18, 19, 31, 32, 36, 37, 90)

    For rowNumber = 1 To expectedRowCount
        mismatch = CompareGoldenNear( _
            caseId, _
            "row-" & CStr(rowNumber) & "-height", _
            CDbl(expectedSheet.Rows(rowNumber).RowHeight), _
            CDbl(actualSheet.Rows(rowNumber).RowHeight))
        If Len(mismatch) > 0 Then GoTo MismatchFound

        For Each columnNumber In formatColumns
            mismatch = CompareGoldenCellFormat( _
                caseId, _
                expectedSheet.Cells(rowNumber, CLng(columnNumber)), _
                actualSheet.Cells(rowNumber, CLng(columnNumber)))
            If Len(mismatch) > 0 Then GoTo MismatchFound
        Next columnNumber
    Next rowNumber

    If compareColumnWidths Then
        For columnNumber = 1 To 90
            mismatch = CompareGoldenNear( _
                caseId, _
                "column-" & CStr(columnNumber) & "-width", _
                CDbl(expectedSheet.Columns(CLng(columnNumber)).ColumnWidth), _
                CDbl(actualSheet.Columns(CLng(columnNumber)).ColumnWidth))
            If Len(mismatch) > 0 Then GoTo MismatchFound
        Next columnNumber
    End If

    actualSheet.Activate
    If Application.ActiveWindow.DisplayGridlines Then
        CompareOutputGoldenFormat = caseId & " gridlines should be hidden"
    End If
    Exit Function

MismatchFound:
    CompareOutputGoldenFormat = mismatch
    Exit Function

CompareError:
    CompareOutputGoldenFormat = caseId & " format comparison error: " & Err.Description
End Function

' セルの塗り、フォント、折り返し、四辺の罫線を比較
Private Function CompareGoldenCellFormat( _
    ByVal caseId As String, _
    ByVal expectedCell As Range, _
    ByVal actualCell As Range) As String

    Dim address As String
    Dim borderIndex As Variant
    Dim expectedBorder As Border
    Dim actualBorder As Border
    Dim mismatch As String

    address = expectedCell.Address(False, False)
    mismatch = CompareGoldenEqual(caseId, address & " fill", expectedCell.Interior.Color, actualCell.Interior.Color)
    If Len(mismatch) > 0 Then GoTo MismatchFound

    mismatch = CompareGoldenEqual(caseId, address & " font", expectedCell.Font.Name, actualCell.Font.Name)
    If Len(mismatch) > 0 Then GoTo MismatchFound

    mismatch = CompareGoldenNear(caseId, address & " font-size", CDbl(expectedCell.Font.Size), CDbl(actualCell.Font.Size))
    If Len(mismatch) > 0 Then GoTo MismatchFound

    mismatch = CompareGoldenEqual(caseId, address & " wrap", expectedCell.WrapText, actualCell.WrapText)
    If Len(mismatch) > 0 Then GoTo MismatchFound

    For Each borderIndex In Array(7, 8, 9, 10)
        Set expectedBorder = expectedCell.Borders(CLng(borderIndex))
        Set actualBorder = actualCell.Borders(CLng(borderIndex))
        mismatch = CompareGoldenEqual( _
            caseId, _
            address & " border-" & CStr(borderIndex), _
            expectedBorder.LineStyle, _
            actualBorder.LineStyle)
        If Len(mismatch) > 0 Then GoTo MismatchFound

        If CLng(expectedBorder.LineStyle) <> NO_BORDER Then
            mismatch = CompareGoldenEqual( _
                caseId, _
                address & " border-" & CStr(borderIndex) & "-weight", _
                expectedBorder.Weight, _
                actualBorder.Weight)
            If Len(mismatch) > 0 Then GoTo MismatchFound

            mismatch = CompareGoldenEqual( _
                caseId, _
                address & " border-" & CStr(borderIndex) & "-color", _
                expectedBorder.Color, _
                actualBorder.Color)
            If Len(mismatch) > 0 Then GoTo MismatchFound
        End If
    Next borderIndex
    Exit Function

MismatchFound:
    CompareGoldenCellFormat = mismatch
End Function

' 文字列表現が異なる値の比較メッセージを生成
Private Function CompareGoldenEqual( _
    ByVal caseId As String, _
    ByVal location As String, _
    ByVal expectedValue As Variant, _
    ByVal actualValue As Variant) As String

    If StrComp(CStr(expectedValue), CStr(actualValue), vbBinaryCompare) <> 0 Then
        CompareGoldenEqual = caseId & " " & location & _
            " expected=[" & CStr(expectedValue) & "] actual=[" & CStr(actualValue) & "]"
    End If
End Function

' 許容差を超えた数値の比較メッセージを生成
Private Function CompareGoldenNear( _
    ByVal caseId As String, _
    ByVal location As String, _
    ByVal expectedValue As Double, _
    ByVal actualValue As Double) As String

    If Abs(expectedValue - actualValue) > VALUE_TOLERANCE Then
        CompareGoldenNear = caseId & " " & location & _
            " expected=[" & CStr(expectedValue) & "] actual=[" & CStr(actualValue) & "]"
    End If
End Function
