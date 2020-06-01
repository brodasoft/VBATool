VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEnronExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const CLS_ID As Long = 428
Private Const CLS_VER As String = "3.26"
Private Const CLS_NAME As String = "EnronExcel"
Private Const CLS_DATE As Date = #6/1/2020#
Private Const CLS_SKP As Boolean = True
Private Const CLS_CFG As String = "1"

'- excel speeding options
'- file open
'- toolbar
'- statusbar
'- remove PPV
'- timer

Option Explicit
'slawomirbroda@gmail.com

Private cApp As Excel.Application
Private cPObj As Excel.Workbook

Private cAppScreen As Boolean
Private cAppCalculcation As Long
Private cAppEvents As Boolean

Private cMsg As String
Private cStart As Long
Private cStep As Long
Private cCurStep As Long
Private cFinish As Long
Private cTime As Date
Private cTrace As Boolean

'ini
Private Sub Class_Initialize()
    Set cApp = Application
    Set cPObj = cApp.ThisWorkbook
End Sub

Private Sub Class_Terminate()
    Call StatBarEnd
    'clean global
    Set cPObj = Nothing
    Set cApp = Nothing
End Sub

'prop
Public Property Get msgB() As String
    On Error Resume Next
    msgB = " ([" & GetProp("AppID") & "] " & GetProp("AppName") & " ver. " & GetProp("AppVer") & " #" & GetProp("AppDate") & "#)" & vbLf
End Property

Public Property Get Timer(Optional STP As Boolean) As Long
    Timer = DateDiff("s", cTime, Now)
    If STP = True Then TimerEnd
End Property

Public Property Get TRACE() As Boolean
    TRACE = cTrace
End Property

Public Property Let TRACE(prop As Boolean)
    cTrace = prop
End Property

Private Function GetProp(prop As String) As String
    On Error Resume Next
    GetProp = cPObj.CustomDocumentProperties(prop).Value
End Function

'public
Sub OptTurnOn(Optional ByRef excApp As Excel.Application)
    
    If excApp Is Nothing Then Set excApp = cApp
        
    If cTrace = False Then
        cAppScreen = excApp.ScreenUpdating
        cAppCalculcation = excApp.Calculation
        cAppEvents = excApp.EnableEvents
        
        excApp.ScreenUpdating = False
        excApp.Calculation = xlCalculationManual
        excApp.EnableEvents = False
    Else
        Call OptTurnBack
    End If
End Sub

Sub OptTurnOff(Optional ByRef excApp As Excel.Application)
    If excApp Is Nothing Then Set excApp = cApp
    excApp.ScreenUpdating = cAppScreen
    excApp.Calculation = cAppCalculcation
    excApp.EnableEvents = cAppEvents
End Sub

Sub OptTurnBack(Optional ByRef excApp As Excel.Application)
    If excApp Is Nothing Then Set excApp = cApp
    excApp.ScreenUpdating = True
    excApp.Calculation = xlAutomatic
    excApp.EnableEvents = True
End Sub

Sub RemovePeerReview(Optional wb As Workbook)
    Dim ws As Worksheet
    
    On Error Resume Next
    If wb Is Nothing Then Set wb = cPObj
    wb.CustomDocumentProperties("MPR_PEERREVIEW").Delete
    
    For Each ws In wb.Worksheets
        If ws.Shapes.Count > 0 Then
            If ws.Shapes(1).name = "GLOBALPEERREVIEW" Then ws.Shapes(1).Delete
        End If
    Next ws
    On Error GoTo 0
End Sub

Sub CreateToolbar(tbName As String, vbaProc As Variant, bName As Variant, fid As Variant)
    Dim cbToolbar As CommandBar
    Dim ctButton As CommandBarButton
    Dim cnt As Long
    
    If UBound(vbaProc) <> UBound(fid) Then Exit Sub
        
    Call DeleteToolbar(tbName)
    
    Set cbToolbar = cApp.CommandBars.Add(tbName, msoBarRight, False, True)
        
    For cnt = LBound(vbaProc) To UBound(vbaProc)
        With cbToolbar
            Set ctButton = .Controls.Add(Type:=msoControlButton, id:=cbToolbar.Controls.Count + 1)
        End With
        
        With ctButton
            .Style = msoButtonIconAndWrapCaptionBelow
            If UBound(vbaProc) = 0 Then
                .Caption = msgB
            Else
                .Caption = bName(cnt)
            End If
            .FaceId = fid(cnt)
            .OnAction = vbaProc(cnt)
            .BeginGroup = True
        End With
    Next cnt
    
    With cbToolbar
        .Visible = True
        .Protection = msoBarNoChangeVisible
    End With
    
    Set cbToolbar = Nothing
    Set ctButton = Nothing
End Sub

Sub DeleteToolbar(tbName As String)
    Dim wnd As Window
    On Error Resume Next
    For Each wnd In cApp.Windows
        'wnd.Activate
        cApp.CommandBars(tbName).Delete
    Next wnd
    On Error GoTo 0
End Sub

Sub StatBarRun(Msg As String, START As Long, Finish As Long, Optional step As Long)
    cApp.StatusBar = False
    cMsg = Msg
    cStart = START
    If step = 0 Then
        cStep = Finish / 100
        If cStep = 0 Then cStep = 1
    Else
        cStep = step
    End If
    cFinish = Finish
    cCurStep = 0
End Sub

Sub StatBarNext(Optional ByVal Msg As String)
    Dim prc As Long
    Const sg1 As String = ","
    Const sg2 As String = "|"
    
    cCurStep = cCurStep + 1
    prc = CInt((cCurStep / cFinish) * 100)
    If cCurStep Mod cStep = 0 Then
        If Msg <> vbNullString Then Msg = "(" & Msg & "): "
        cApp.StatusBar = cMsg & " " & String(prc, sg1) & String(100 - prc, sg2) & " " & Msg & " " & cCurStep & " of " & cFinish
        DoEvents
    End If
End Sub

Sub StatBarEnd(Optional ByVal Msg As Boolean)
    Const sg1 As String = ","
    Const sg2 As String = "|"
    
    cApp.StatusBar = cMsg & " " & String(100, sg1) & String(0, sg2) & " " & cFinish & " of " & cFinish
    If Msg = True Then MsgBox cMsg & " done.", vbInformation, msgB
    cApp.StatusBar = False
End Sub

Sub StatBarMsg(ByVal Msg As String)
    cApp.StatusBar = Msg
End Sub

Sub TimerStart()
    cTime = Now()
End Sub

Sub TimerEnd()
    cTime = 0
End Sub

Function GetListOfFiles(fName As String, fType As String, Optional tTl As String, Optional isMulti As Boolean, Optional startPth As String, Optional isFld As Boolean) As Variant
    Dim tp As Long
    Dim arr() As String
    Dim i As Long
    
    If isFld = True Then
        tp = msoFileDialogFolderPicker
    Else
        tp = msoFileDialogFilePicker
    End If
    
    If tTl = vbNullString Then tTl = "Select file"
    If startPth = vbNullString Then startPth = cPObj.Path
    
    With cApp.FileDialog(tp)
        .InitialFileName = startPth
        .AllowMultiSelect = isMulti
        .Title = tTl
        If isFld = False Then
            .Filters.Clear
            .Filters.Add fName, fType            '"Excel file - *.xl*", "*.xl*"
        End If
        .Show
        If .SelectedItems.Count > 0 Then
            ReDim arr(1 To .SelectedItems.Count)
            For i = 1 To .SelectedItems.Count
                arr(i) = .SelectedItems(i)
            Next i
            GetListOfFiles = arr
        End If
    End With
End Function

Sub ExcelCleaner(Optional list As Variant)
    Dim COMad As COMAddIn
    Dim ad As AddIn
    Dim bar As CommandBar
    
    On Error Resume Next
    
    For Each COMad In cApp.COMAddIns
        If COMad.Connect Then
            COMad.Connect = False
        End If
    Next COMad
    
    For Each ad In cApp.AddIns
        If ad.Installed = True And ad.name <> cPObj.name Then
            ad.Installed = False
        End If
    Next ad
    
    For Each bar In cApp.CommandBars
        If bar.BuiltIn = False Then
            bar.Delete
        End If
    Next bar

End Sub

Function FileExtStr(FF As Byte) As String
    Select Case FF
    Case 51: FileExtStr = ".xlsx"
    Case 52: FileExtStr = ".xlsm"
    Case 55: FileExtStr = ".xlam"
    Case 56: FileExtStr = ".xls"
    Case 50: FileExtStr = ".xlsb"
    End Select
End Function

