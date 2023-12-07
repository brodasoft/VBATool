VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cExcelAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const CLS_ID As Long = 0
Private Const CLS_VER As String = "0.27"
Private Const CLS_NAME As String = "cExcelAppEvents"
Private Const CLS_DATE As Date = #6/1/2020#
Private Const CLS_SKP As Boolean = True
Private Const CLS_CFG As String = "0"

'excel events handler

Option Explicit
'slawomirbroda@gmail.com

Public WithEvents cApp As Excel.Application
Attribute cApp.VB_VarHelpID = -1

'start
Private Sub Class_Initialize()
    Set cApp = Application
End Sub

Private Sub Class_Terminate()
    Set cApp = Nothing
End Sub

'events
Private Sub cApp_WorkbookBeforeSave(ByVal wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    'gEnron.RemovePeerReview wb
End Sub

