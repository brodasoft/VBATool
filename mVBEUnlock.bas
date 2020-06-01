Attribute VB_Name = "mVBEUnlock"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As LongPtr, ByVal hwndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
    Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetDlgItem Lib "user32" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long) As LongPtr ' nIDDlgItem = int?
    Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare PtrSafe Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hWndLock As LongPtr) As Long
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal uIDEvent As LongPtr) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long ' nIDDlgItem = int?
    Private Declare Function GetDesktopWindow Lib "user32" () As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
    Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
    Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal uIDEvent As Long) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If



Private Const WM_CLOSE As Long = &H10
Private Const WM_GETTEXT As Long = &HD
Private Const EM_REPLACESEL As Long = &HC2
Private Const EM_SETSEL As Long = &HB1
Private Const BM_CLICK As Long = &HF5&
Private Const TCM_SETCURFOCUS As Long = &H1330&
Private Const IDPassword As Long = &H155E&
Private Const IDOK As Long = &H1&
Private Const IDOK2 As Long = &H2&
Private Const IDCNL As Long = &H2&

Private Const TIMEOUTSECOND As Long = 1

Private g_ProjectName    As String
Private g_Password       As String
Private g_Result         As Long
#If VBA7 Then
    Private g_hwndVBE        As LongPtr
    Private g_hwndPassword   As LongPtr
    Private g_hwndErr As LongPtr
#Else
    Private g_hwndVBE        As Long
    Private g_hwndPassword   As Long
    Private g_hwndErr         As Long
#End If
Sub Test_UnlockProject()
    
    Dim wb As Excel.Workbook
    
    'Set wb = Application.Workbooks.Open("c:\Enron\_Mercer\Projects\test.xlsm", , True)
    Set wb = Application.Workbooks("Szychtownice 2019.05.23_0.01.xlsm")
    'Application.Visible = True
    MultiPassUnlockProject wb.vbproject, Array("zz", getFolderPass(wb))

    'wb.Close False
    'Application.Quit
End Sub
Sub MultiPassUnlockProject(ByVal Project As Object, ByVal ProjPass As Variant)
    Dim Password As Variant
    Dim retVal As Long
    If Project.Protection <> 1 Then Exit Sub
    
    For Each Password In ProjPass
        retVal = UnlockProject(Project, Password)
            Select Case retVal
                Case 0
                    Debug.Print "The project was unlocked."
                    Exit For
                Case 2
                    Debug.Print "The project was already unlocked."
                    Exit For
                Case 3
                    Debug.Print "Wrong Password: " & Password
                Case Else
                    Debug.Print "Error or timeout"
            End Select
    Next Password
    
    Exit Sub

End Sub
Public Function UnlockProject(ByVal Project As Object, ByVal Password As String) As Long

    #If VBA7 Then
        Dim lRet As LongPtr
    #Else
        Dim lRet As Long
    #End If
    Dim timeout As Date

    On Error GoTo ErrorHandler
    UnlockProject = 1
        
        If Project.Protection <> 1 Then              'vbext_pp_locked
            UnlockProject = 2
            Exit Function
        End If

    g_ProjectName = Project.name
    g_Password = Password
    g_Result = 0
    
    'LockWindowUpdate GetDesktopWindow()
    
    Application.VBE.MainWindow.Visible = True
    g_hwndVBE = Application.VBE.MainWindow.hwnd

    lRet = SetTimer(0, 1, 100, AddressOf UnlockTimerProc)
        If lRet = 0 Then
            Debug.Print "Err" & "error setting timer"
            GoTo ErrorHandler
        End If

    Set Application.VBE.ActiveVBProject = Project
        If Not Application.VBE.ActiveVBProject Is Project Then GoTo ErrorHandler

    Application.VBE.CommandBars.FindControl(id:=2578).Execute

    timeout = Now() + TimeSerial(0, 0, TIMEOUTSECOND * 5)
        Do While g_Result = 0 And Now() < timeout
            DoEvents
        Loop
        
    Application.VBE.MainWindow.Visible = True
        
    If g_Result = 1 Then
        UnlockProject = 0
    ElseIf g_Result = 3 Then
        UnlockProject = 3
    End If
    
    'Debug.Print g_Result
    
ErrorHandler:
    'LockWindowUpdate 0

End Function

#If VBA7 Then
Private Function UnlockTimerProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long) As Long
#Else
Private Function UnlockTimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long) As Long
#End If

#If VBA7 Then
    Dim hWndPassword As LongPtr
    Dim hWndOK As LongPtr
    Dim lRet As LongPtr
#Else
    Dim hWndPassword As Long
    Dim hWndOK As Long
    Dim lRet As Long
#End If
Dim lRet2 As Long
Dim scaption As String
Dim timeout As Date
Dim timeout2 As Date

On Error GoTo ErrorHandler
KillTimer 0, idEvent

Select Case Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    ' For the japanese version
Case 1041
    scaption = ChrW(&H30D7) & ChrW(&H30ED) & ChrW(&H30B8) & _
               ChrW(&H30A7) & ChrW(&H30AF) & ChrW(&H30C8) & _
               ChrW(&H20) & ChrW(&H30D7) & ChrW(&H30ED) & _
               ChrW(&H30D1) & ChrW(&H30C6) & ChrW(&H30A3)
Case Else
    scaption = " Password"
End Select
scaption = g_ProjectName & scaption

timeout = Now() + TimeSerial(0, 0, TIMEOUTSECOND)
Do While Now() < timeout
    
    hWndPassword = 0
    hWndOK = 0
    g_hwndPassword = 0
    
        Do
            g_hwndPassword = FindWindowEx(0, g_hwndPassword, vbNullString, scaption)
            If g_hwndPassword = 0 Then Exit Do
            
        Loop Until GetParent(g_hwndPassword) = g_hwndVBE
    
        If g_hwndPassword = 0 Then GoTo Continue
    
    lRet2 = SendMessage(g_hwndPassword, TCM_SETCURFOCUS, 1, ByVal 0&)
    hWndPassword = GetDlgItem(g_hwndPassword, IDPassword)
    hWndOK = GetDlgItem(g_hwndPassword, IDOK)
    
        If (g_hwndPassword And hWndOK) = 0 Then GoTo Continue

    lRet = SetFocusAPI(hWndPassword)
    lRet2 = SendMessage(hWndPassword, EM_SETSEL, 0, ByVal -1&)
    lRet2 = SendMessage(hWndPassword, EM_REPLACESEL, 0, ByVal g_Password)

    lRet = SetTimer(0, 2, 100, AddressOf ClosePropertiesWindow)
    lRet = SetTimer(0, 3, 100, AddressOf CloseError)
    
    lRet = SetFocusAPI(hWndOK)
    lRet2 = SendMessage(hWndOK, BM_CLICK, 0, ByVal 0&)
  
    'g_Result = 1
    Exit Do
    
Continue:
    DoEvents
    Sleep 100
Loop
Exit Function

ErrorHandler:
Debug.Print "Err1" & Err.Number
    If hWndPassword <> 0 Then SendMessage hWndPassword, WM_CLOSE, 0, ByVal 0&
'LockWindowUpdate 0

End Function

#If VBA7 Then
Function ClosePropertiesWindow(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long) As Long
#Else
Function ClosePropertiesWindow(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long) As Long
#End If

#If VBA7 Then
    Dim hWndTmp As LongPtr
    Dim hWndOK As LongPtr
    Dim lRet As LongPtr
#Else
    Dim hWndTmp As Long
    Dim hWndOK As Long
    Dim lRet As Long
#End If
Dim lRet2 As Long
Dim timeout As Date
Dim scaption As String

On Error GoTo ErrorHandler

KillTimer 0, idEvent
    
scaption = g_ProjectName & " - Project Properties"
    
timeout = Now() + TimeSerial(0, 0, TIMEOUTSECOND)
Do While Now() < timeout

    hWndTmp = 0

    Do
        hWndTmp = FindWindowEx(0, hWndTmp, vbNullString, scaption)
        If hWndTmp = 0 Then Exit Do
    Loop Until GetParent(hWndTmp) = g_hwndVBE

    If hWndTmp = 0 Then GoTo Continue

    lRet2 = SendMessage(hWndTmp, TCM_SETCURFOCUS, 1, ByVal 0&)
    hWndOK = GetDlgItem(hWndTmp, IDOK)

    If (hWndTmp And hWndOK) = 0 Then GoTo Continue
        
    lRet = SetFocusAPI(hWndOK)
    lRet2 = SendMessage(hWndOK, BM_CLICK, 0, ByVal 0&)

    Exit Do

Continue:
    DoEvents
    Sleep 100
Loop
Exit Function

ErrorHandler:
Debug.Print "Err2" & Err.Number
'LockWindowUpdate 0

End Function

#If VBA7 Then
Function CloseError(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long) As Long
#Else
Function CloseError(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long) As Long
#End If

#If VBA7 Then
    Dim hWndOK As LongPtr
    Dim lRet As LongPtr
#Else

    Dim hWndOK As Long
    Dim lRet As Long
#End If
Dim lRet2 As Long
Dim timeout As Date
Dim scaption As String

On Error GoTo ErrorHandler

KillTimer 0, idEvent
    
scaption = "Project Locked"

g_Result = 1 ' zak³adamy, ¿e nie ma b³êdu

timeout = Now() + TimeSerial(0, 0, TIMEOUTSECOND)
Do While Now() < timeout

    g_hwndErr = 0

        Do
            g_hwndErr = FindWindowEx(0, g_hwndErr, vbNullString, scaption)
            'Debug.Print "WndErr: " & g_hwndErr & ",WndErrParr: " & GetParent(g_hwndErr) & ", ParrPASS: " & g_hwndPassword
            If g_hwndErr = 0 Then Exit Do
        Loop Until GetParent(g_hwndErr) = g_hwndPassword
    
        If g_hwndErr = 0 Then GoTo Continue
    hWndOK = GetDlgItem(g_hwndErr, IDOK2)
    
        If (g_hwndErr And hWndOK) = 0 Then GoTo Continue
    lRet = SetFocusAPI(hWndOK)
    lRet2 = SendMessage(hWndOK, BM_CLICK, 0, ByVal 0&)
    
    hWndOK = GetDlgItem(g_hwndPassword, IDCNL)
    lRet = SetFocusAPI(hWndOK)
    lRet2 = SendMessage(hWndOK, BM_CLICK, 0, ByVal 0&)
    'Debug.Print "click locked window"
    
    g_Result = 3        ' jest b³¹d
    Exit Do

Continue:
    DoEvents
    Sleep 100
Loop

Exit Function

ErrorHandler:
Debug.Print "Err3" & Err.Number
'LockWindowUpdate 0

End Function

Function getFolderPass(wb As Workbook, Optional up As Boolean) As String
    Dim fso As New FileSystemObject
    Dim fld As Folder: Set fld = fso.GetFolder(wb.Path)
    
    If up = False Then
        getFolderPass = fld.name
    Else
        getFolderPass = fld.ParentFolder.name
    End If
    
End Function





