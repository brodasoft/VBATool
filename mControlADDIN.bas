Attribute VB_Name = "mControlADDIN"
Option Explicit
Public gEae As cExcelAppEvents

Sub Workbook_Initialize()
    Set gEae = New cExcelAppEvents
    'gvt.CheckDBVer
    gEae.cApp.OnKey "^%{BS}", "prcUnlockVBA"
    
End Sub

Sub Workbook_Terminate()
    gEae.cApp.OnKey "^%{BS}"
    
    Set gEae = Nothing
    Set gEnron = Nothing
    Set gVt = Nothing
    
    
End Sub

Private Sub clsMEnronExcel()
    gVt.IsModule = "cEnronExcel"
    gVt.AppCfgSet ReleseMode
    'gVt.AppVerMinor
    'gVt.CheckDBVer
    Debug.Print gVt.msgB
    
    Set gVt = Nothing
End Sub

Private Sub prcMVersionTracker()
    
    gVt.IsModule = "cVersionTracker"
    gVt.AppCfgSet ReleseMode
    'gVt.AppVerMinor
    'gVt.CheckDBVer
    Debug.Print gVt.msgB
    
    Set gVt = Nothing
End Sub

Private Sub prcMMakra()
    
    gVt.IsModule = ""
    'gVt.AppCfgSet ReleseMode
    'gVt.AppVerMinor
    gVt.SaveFile
    'gVt.CheckDBVer
    

    Debug.Print gVt.msgB
    
    Set gVt = Nothing
End Sub

Private Sub prcMExcelAppEvents()
    
    gVt.IsModule = "cExcelAppEvents"
    'gvt.AppCfgSet DebugMode
    'gVt.AppVerMinor
    'gvt.CheckDBVer
    Debug.Print gVt.msgB
    
    Set gVt = Nothing
End Sub

Public Sub prcUnlockVBA()
    Dim wb As Workbook
    If gEae Is Nothing Then Workbook_Initialize
    For Each wb In gEae.cApp.Workbooks
        If wb.MultiUserEditing = False And wb.vbproject.Protection = vbext_pp_locked Then              'workbook isn't shared, no access to VBA
            MultiPassUnlockProject wb.vbproject, Array("warsaw2015", getFolderPass(wb), getFolderPass(wb, True), "warsaw2014", "january88", "athens297", "boston76")
        End If
    Next wb
End Sub


