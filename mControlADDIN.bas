Attribute VB_Name = "mControlADDIN"
Option Explicit
Public gEae As cExcelAppEvents

Sub Workbook_Initialize()
    Set gEae = New cExcelAppEvents
    'gvt.CheckDBVer
    gEae.cApp.OnKey "^%{BS}", "prcUnlockVBA"
    
End Sub

Sub Workbook_Terminate()
    Set gEae = Nothing
    Set gEnron = Nothing
    Set gVt = Nothing
    
    gEae.cApp.OnKey "^%{BS}"
End Sub

Private Sub clsMEnronExcel()
    gVt.IsModule = "cEnronExcel"
    'gvt.AppCfgSet DebugMode
    gVt.AppVerMinor
    'gvt.CheckDBVer
    Debug.Print gVt.msgB
    
    Set gVt = Nothing
End Sub

Private Sub prcMVersionTracker()
    
    gVt.IsModule = "cVersionTracker"
    'gvt.AppCfgSet DebugMode
    gVt.AppVerMinor
    'gvt.CheckDBVer
    Debug.Print gVt.msgB
    
    Set gVt = Nothing
End Sub

Private Sub prcMMakra()
    'gvt.AppCfgSet DebugMode
    gVt.AppVerMinor
    gVt.SaveFile
    'gvt.CheckDBVer
    'gvt.AppNameSet "Makra"
    Debug.Print gVt.msgB
    
    Set gVt = Nothing
End Sub

Private Sub prcMExcelAppEvents()
    
    gVt.IsModule = "cExcelAppEvents"
    'gvt.AppCfgSet DebugMode
    'gvt.AppVerSet "0.23"
    gVt.AppVerMinor
    'gvt.CheckDBVer
    Debug.Print gVt.msgB
    
    Set gVt = Nothing
End Sub

Public Sub prcUnlockVBA()
    Dim wb As Workbook
    For Each wb In gEae.cApp.Workbooks
        If wb.MultiUserEditing = False And wb.vbproject.Protection = vbext_pp_locked Then              'workbook isn't shared, no access to VBA
            MultiPassUnlockProject wb.vbproject, Array("warsaw2015", getFolderPass(wb), getFolderPass(wb, True), "warsaw2014", "january88", "athens297", "boston76")
        End If
    Next wb
End Sub


