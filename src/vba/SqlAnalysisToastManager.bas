Attribute VB_Name = "SqlAnalysisToastManager"
Option Explicit

Private Const TOAST_FORM_NAME As String = "SqlAnalysisToast"
Private Const TOAST_LABEL_NAME As String = "MessageLabel"
Private Const TOAST_DURATION_SECONDS_VALUE As Long = 2
Private Const TOAST_WIDTH As Double = 264#
Private Const TOAST_HEIGHT As Double = 54#
Private Const TOAST_RIGHT_MARGIN As Double = 20#
Private Const TOAST_BOTTOM_MARGIN As Double = 70#
Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_EX_TOOLWINDOW As Long = &H80&
Private Const WS_EX_NOACTIVATE As Long = &H8000000
Private Const SW_SHOWNOACTIVATE As Long = 4
Private Const SWP_NOSIZE As Long = &H1&
Private Const SWP_NOMOVE As Long = &H2&
Private Const SWP_NOZORDER As Long = &H4&
Private Const SWP_NOACTIVATE As Long = &H10&
Private Const SWP_FRAMECHANGED As Long = &H20&
Private Const SWP_SHOWWINDOW As Long = &H40&

Private Type WindowRectangle
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal className As String, _
        ByVal windowName As String) As LongPtr
#If Win64 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
        ByVal windowHandle As LongPtr, _
        ByVal index As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
        ByVal windowHandle As LongPtr, _
        ByVal index As Long, _
        ByVal newValue As LongPtr) As LongPtr
#Else
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" ( _
        ByVal windowHandle As LongPtr, _
        ByVal index As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" ( _
        ByVal windowHandle As LongPtr, _
        ByVal index As Long, _
        ByVal newValue As LongPtr) As LongPtr
#End If
    Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
        ByVal windowHandle As LongPtr, _
        ByVal insertAfter As LongPtr, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal width As Long, _
        ByVal height As Long, _
        ByVal flags As Long) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" ( _
        ByVal windowHandle As LongPtr, _
        ByVal command As Long) As Long
    Private Declare PtrSafe Function GetWindowRect Lib "user32" ( _
        ByVal windowHandle As LongPtr, _
        ByRef rectangle As WindowRectangle) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal className As String, _
        ByVal windowName As String) As Long
    Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" ( _
        ByVal windowHandle As Long, _
        ByVal index As Long) As Long
    Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" ( _
        ByVal windowHandle As Long, _
        ByVal index As Long, _
        ByVal newValue As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" ( _
        ByVal windowHandle As Long, _
        ByVal insertAfter As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal width As Long, _
        ByVal height As Long, _
        ByVal flags As Long) As Long
    Private Declare Function ShowWindow Lib "user32" ( _
        ByVal windowHandle As Long, _
        ByVal command As Long) As Long
    Private Declare Function GetWindowRect Lib "user32" ( _
        ByVal windowHandle As Long, _
        ByRef rectangle As WindowRectangle) As Long
#End If

Private mToastForm As Object
Private mApplicationEvents As SqlAnalysisToastEvents
Private mToastScheduled As Boolean
Private mToastDismissAt As Date
Private mToastProcedure As String
Private mToastMessage As String

' Display a non-blocking completion notification for two seconds.
Public Sub ShowToast(ByVal message As String)
#If VBA7 Then
    Dim toastWindow As LongPtr
#Else
    Dim toastWindow As Long
#End If

    On Error GoTo ToastFailed
    If Len(message) = 0 Then Exit Sub

    EnsureApplicationEvents
    CancelScheduledToast

    If mToastForm Is Nothing Then
        Set mToastForm = VBA.UserForms.Add(TOAST_FORM_NAME)
        mToastForm.Caption = ToastWindowCaption()
    End If

    ConfigureToastForm mToastForm, message
    toastWindow = ConfigureNoActivateWindow(CStr(mToastForm.Caption))
    If toastWindow = 0 Then GoTo ToastFailed
    mToastForm.Show vbModeless
    PositionToastWindow toastWindow
    ShowWindow toastWindow, SW_SHOWNOACTIVATE

    mToastMessage = message
    ScheduleToastDismissal
    Exit Sub

ToastFailed:
    CancelScheduledToast
    HideToastForm
End Sub

' Immediately dismiss the current notification. Called by the form click handlers.
Public Sub DismissToast()
    On Error Resume Next
    CancelScheduledToast
    HideToastForm
    On Error GoTo 0
End Sub

' Application.OnTime callback. Old callbacks must not close a newer notification.
Public Sub ToastTimerElapsed()
    On Error GoTo TimerCleanUp
    If Not mToastScheduled Then Exit Sub
    If Now < mToastDismissAt Then Exit Sub

TimerCleanUp:
    mToastScheduled = False
    mToastDismissAt = 0
    mToastProcedure = vbNullString
    HideToastForm
End Sub

' Cancel the timer before this workbook is closed.
Public Sub ShutdownToast()
    DismissToast
    Set mApplicationEvents = Nothing
End Sub

' Exposed for deterministic VBA tests and documentation of the UX contract.
Public Function ToastDurationSeconds() As Long
    ToastDurationSeconds = TOAST_DURATION_SECONDS_VALUE
End Function

' Exposed for VBA/COM integration tests.
Public Function ToastIsVisible() As Boolean
    On Error GoTo NotVisible
    If mToastForm Is Nothing Then Exit Function
    ToastIsVisible = CBool(mToastForm.Visible)
    Exit Function

NotVisible:
    ToastIsVisible = False
End Function

' Exposed for VBA/COM integration tests.
Public Function CurrentToastMessage() As String
    CurrentToastMessage = mToastMessage
End Function

' Exposed for VBA tests that verify repeated notices restart the two-second timer.
Public Function CurrentToastDismissalTime() As Date
    CurrentToastDismissalTime = mToastDismissAt
End Function

' Exposed for VBA tests that guard the non-activating, borderless toast contract.
Public Function ToastWindowStyleIsValid() As Boolean
#If VBA7 Then
    Dim toastWindow As LongPtr
    Dim windowStyle As LongPtr
    Dim extendedStyle As LongPtr
#Else
    Dim toastWindow As Long
    Dim windowStyle As Long
    Dim extendedStyle As Long
#End If

    On Error GoTo InvalidStyle
    If mToastForm Is Nothing Then Exit Function
    toastWindow = FindWindow(vbNullString, CStr(mToastForm.Caption))
    If toastWindow = 0 Then Exit Function
    windowStyle = GetWindowLongPtr(toastWindow, GWL_STYLE)
    extendedStyle = GetWindowLongPtr(toastWindow, GWL_EXSTYLE)
    ToastWindowStyleIsValid = _
        ((windowStyle And (WS_CAPTION Or WS_THICKFRAME)) = 0) And _
        ((extendedStyle And WS_EX_NOACTIVATE) <> 0)
    Exit Function

InvalidStyle:
    ToastWindowStyleIsValid = False
End Function

Private Sub EnsureApplicationEvents()
    If mApplicationEvents Is Nothing Then
        Set mApplicationEvents = New SqlAnalysisToastEvents
        Set mApplicationEvents.ExcelApplication = Application
    End If
End Sub

Private Sub ConfigureToastForm(ByVal toastForm As Object, ByVal message As String)
    Dim messageLabel As Object

    toastForm.Width = TOAST_WIDTH
    toastForm.Height = TOAST_HEIGHT
    toastForm.BackColor = RGB(33, 115, 70)
    Set messageLabel = toastForm.Controls(TOAST_LABEL_NAME)
    messageLabel.Caption = message
    messageLabel.Left = 14
    messageLabel.Top = 8
    messageLabel.Width = toastForm.InsideWidth - 28
    messageLabel.Height = toastForm.InsideHeight - 16
    messageLabel.BackStyle = 0
    messageLabel.ForeColor = RGB(255, 255, 255)
    messageLabel.TextAlign = 2
    messageLabel.WordWrap = False

    On Error Resume Next
    messageLabel.Font.Name = "Yu Gothic UI"
    messageLabel.Font.Size = 10
    On Error GoTo 0
End Sub

#If VBA7 Then
Private Function ConfigureNoActivateWindow(ByVal caption As String) As LongPtr
    Dim toastWindow As LongPtr
    Dim windowStyle As LongPtr
    Dim extendedStyle As LongPtr

    toastWindow = FindWindow(vbNullString, caption)
    If toastWindow = 0 Then Exit Function
    windowStyle = GetWindowLongPtr(toastWindow, GWL_STYLE)
    SetWindowLongPtr toastWindow, GWL_STYLE, _
        windowStyle And Not (WS_CAPTION Or WS_THICKFRAME)
    extendedStyle = GetWindowLongPtr(toastWindow, GWL_EXSTYLE)
    SetWindowLongPtr toastWindow, GWL_EXSTYLE, _
        extendedStyle Or WS_EX_TOOLWINDOW Or WS_EX_NOACTIVATE
    SetWindowPos toastWindow, 0, 0, 0, 0, 0, _
        SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
    ConfigureNoActivateWindow = toastWindow
End Function
#Else
Private Function ConfigureNoActivateWindow(ByVal caption As String) As Long
    Dim toastWindow As Long
    Dim windowStyle As Long
    Dim extendedStyle As Long

    toastWindow = FindWindow(vbNullString, caption)
    If toastWindow = 0 Then Exit Function
    windowStyle = GetWindowLongPtr(toastWindow, GWL_STYLE)
    SetWindowLongPtr toastWindow, GWL_STYLE, _
        windowStyle And Not (WS_CAPTION Or WS_THICKFRAME)
    extendedStyle = GetWindowLongPtr(toastWindow, GWL_EXSTYLE)
    SetWindowLongPtr toastWindow, GWL_EXSTYLE, _
        extendedStyle Or WS_EX_TOOLWINDOW Or WS_EX_NOACTIVATE
    SetWindowPos toastWindow, 0, 0, 0, 0, 0, _
        SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
    ConfigureNoActivateWindow = toastWindow
End Function
#End If

#If VBA7 Then
Private Sub PositionToastWindow(ByVal toastWindow As LongPtr)
#Else
Private Sub PositionToastWindow(ByVal toastWindow As Long)
#End If
    Dim excelRectangle As WindowRectangle
    Dim toastRectangle As WindowRectangle
    Dim toastLeft As Long
    Dim toastTop As Long
    Dim toastWidth As Long
    Dim toastHeight As Long

    If GetWindowRect(Application.hWnd, excelRectangle) = 0 Then Exit Sub
    If GetWindowRect(toastWindow, toastRectangle) = 0 Then Exit Sub
    toastWidth = toastRectangle.Right - toastRectangle.Left
    toastHeight = toastRectangle.Bottom - toastRectangle.Top
    toastLeft = excelRectangle.Right - toastWidth - CLng(TOAST_RIGHT_MARGIN)
    toastTop = excelRectangle.Bottom - toastHeight - CLng(TOAST_BOTTOM_MARGIN)
    If toastLeft < excelRectangle.Left Then toastLeft = excelRectangle.Left
    If toastTop < excelRectangle.Top Then toastTop = excelRectangle.Top

    SetWindowPos toastWindow, 0, toastLeft, toastTop, 0, 0, _
        SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Private Sub ScheduleToastDismissal()
    mToastDismissAt = RoundedSecond(DateAdd("s", TOAST_DURATION_SECONDS_VALUE, Now))
    mToastProcedure = QualifiedProcedureName("SqlAnalysisToastManager.ToastTimerElapsed")
    Application.OnTime _
        EarliestTime:=mToastDismissAt, _
        Procedure:=mToastProcedure, _
        Schedule:=True
    mToastScheduled = True
End Sub

Private Sub CancelScheduledToast()
    Dim scheduledAt As Date
    Dim procedureName As String

    If Not mToastScheduled Then Exit Sub
    scheduledAt = mToastDismissAt
    procedureName = mToastProcedure
    mToastScheduled = False
    mToastDismissAt = 0
    mToastProcedure = vbNullString

    On Error Resume Next
    Application.OnTime _
        EarliestTime:=scheduledAt, _
        Procedure:=procedureName, _
        Schedule:=False
    On Error GoTo 0
End Sub

Private Sub HideToastForm()
    On Error Resume Next
    If Not mToastForm Is Nothing Then
        mToastForm.Hide
        Unload mToastForm
    End If
    Set mToastForm = Nothing
    mToastMessage = vbNullString
    On Error GoTo 0
End Sub

Private Function QualifiedProcedureName(ByVal procedureName As String) As String
    QualifiedProcedureName = "'" & Replace(ThisWorkbook.Name, "'", "''") & _
        "'!" & procedureName
End Function

Private Function RoundedSecond(ByVal value As Date) As Date
    RoundedSecond = CDate(Fix(CDbl(value) * 86400# + 0.5) / 86400#)
End Function

Private Function ToastWindowCaption() As String
    ToastWindowCaption = "SqlAnalysisToast_" & Hex$(CLng(Timer * 1000#))
End Function
