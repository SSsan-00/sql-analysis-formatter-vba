VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlAnalysisToastEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ExcelApplication As Application
Attribute ExcelApplication.VB_VarHelpID = -1

Private Sub ExcelApplication_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    If Wb Is ThisWorkbook Then
        SqlAnalysisToastManager.ShutdownToast
    End If
End Sub
