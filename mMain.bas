Attribute VB_Name = "mMain"
Option Explicit

Public gEnron As New cEnronExcel
Public gVt As New cVersionTracker

Sub Run()
    Dim wbTrg As Workbook
    Dim shTrg As Worksheet
    Dim shSrc As Worksheet
    
    Dim ln As Long
    Dim lstRow As Long
    
    gEnron.TRACE = True
        If gEnron.TRACE = False Then On Error GoTo err_h
    gEnron.OptTurnOn
    'main code
    
    
ende:                                            'clean
    gEnron.OptTurnOff
    Exit Sub
    
err_h:                                           'error
    If Err.Number = 999 Then
        MsgBox "Error:" & Err.Description, vbCritical, gVt.msgB
    Else
        MsgBox "Error:" & Err.Description, vbCritical, gVt.msgB
    End If
    Err.Clear
    Resume ende:
End Sub

