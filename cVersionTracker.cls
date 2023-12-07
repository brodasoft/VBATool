VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cVersionTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const CLS_ID As Long = 401
Private Const CLS_VER As String = "0.44"
Private Const CLS_NAME As String = "VersionTracker"
Private Const CLS_DATE As Date = #6/1/2020#
Private Const CLS_SKP As Boolean = True
Private Const CLS_CFG As String = "1"
'----------------------------------------------------------------
'0. Class by default is working on project (file) level, to work module level us option from #3.
'   In module level version check if for all modules with const CLS_ID in row 1
'1. Copy [VersionTracker] class to your project
'2. Set global instance for class, or local in procedure
'   Public vt As New VersionTracker
'   Dim vt As New VersionTracker
'3. Select specific module using this option to perform any action on module level.
'    vt.IsModule = "Module1"
'4. Set macro ID from database - one time only
'    vt.AppIDSet 999
'5. To set first version
'    vt.AppVerMinor
'5.1 AppVersion by default if changed only in app itself, becouse default appcfg is set to DebugMode
' To copy your version change to datebase you need to change mode to ReleseMode
' vt.AppCfgSet ReleseMode
'6.Add code for auto check against database. For example to Thisworkbook open event.
'    Private Sub WorkbookOpen()
'        vt.CheckDBVer
'    End Sub
'7. Version change options. Version are changed also in database
'    vt.AppVerMinor             - change version by +0.01
'    vt.AppVerMajor             - change version by +1
'    vt.AppVerSet("2.12")       - set your own version
'    vt.SaveFile                - save file with appname, appdate & appver
'---------------------

Option Explicit
'slawomirbroda@gmail.com

Enum tCfgType
    DebugMode = 0
    ReleseMode = 1
End Enum

Enum tAppType
    AnyModule = 1
    ExcelFile = 2
    WordFile = 3
    AccessFile = 4
End Enum

Private cAppType As tAppType
Private cPropDic As Object
Private cModDic As Object
Private cApp As Object
Private cPObj As Object

Private cModuleName As String
Private cDbPath As String
Private cDbPass As String
Private cVerCheck As Boolean
Private cDomain As String

Private Const APP_ID As String = "AppID"
Private Const APP_VER As String = "AppVer"
Private Const APP_NAME As String = "AppName"
Private Const APP_DATE As String = "AppDate"
Private Const APP_SKP As String = "AppSkp"
Private Const APP_CFG As String = "AppCfg"
Private Const SEP As String = "|"

'ini
Private Sub Class_Initialize()
    Set cApp = Application
    If cApp.name = "Microsoft Excel" Then
        Set cPObj = cApp.ThisWorkbook
        cAppType = ExcelFile
    ElseIf cApp.name = "Microsoft Word" Then
        Set cPObj = SetWordObject
        cAppType = WordFile
    ElseIf cApp.name = "Microsoft Access" Then
        Set cPObj = cApp.DBEngine.Workspaces(0).Databases(0)
        cAppType = AccessFile
    End If
    
    If cApp Is Nothing Or cPObj Is Nothing Then
        MsgBox "FATAL OBJECT INITIALIZATION ERROR", vbCritical, "ERROR"
        Class_Terminate
    End If
    
    If VBATrusted Then
        Set cModDic = CreateObject("scripting.Dictionary")
        Call VBEModule                           'check all modules for ID
        
        Set cPropDic = CreateObject("scripting.Dictionary")
        cPropDic.Add APP_ID, "1" & SEP & "CLS_ID" & SEP & "Long" & SEP & vbNullString
        cPropDic.Add APP_VER, "2" & SEP & "CLS_VER" & SEP & "String" & SEP & """"
        cPropDic.Add APP_NAME, "3" & SEP & "CLS_NAME" & SEP & "String" & SEP & """"
        cPropDic.Add APP_DATE, "4" & SEP & "CLS_DATE" & SEP & "Date" & SEP & "#"
        cPropDic.Add APP_SKP, "5" & SEP & "CLS_SKP" & SEP & "Boolean" & SEP & vbNullString
        cPropDic.Add APP_CFG, "6" & SEP & "CLS_CFG" & SEP & "String" & SEP & """"
    End If
    'set domain
    cDomain = GetDomain()
    'get ini values
    cDbPath = GetINI("MACRODB_PATH")
    cDbPass = GetINI("MACRODB_PASS")
End Sub

Private Sub Class_Terminate()
    Set cPObj = Nothing
    Set cApp = Nothing
    Set cPropDic = Nothing
    Set cModDic = Nothing
End Sub

Property Get AppIDGet() As String
    On Error Resume Next
    AppIDGet = GetProp(APP_ID)
End Property

'prop
Property Get AppVerGet() As String
    On Error Resume Next
    AppVerGet = GetProp(APP_VER)
End Property

Property Get AppNameGet() As String
    On Error Resume Next
    AppNameGet = GetProp(APP_NAME)
End Property

Property Get AppDateGet() As String
    On Error Resume Next
    AppDateGet = GetProp(APP_DATE)
End Property

Property Get AppCfgGet() As tCfgType
    On Error Resume Next
    AppCfgGet = GetProp(APP_CFG)
End Property

Property Get msgB() As String
    msgB = IIf(AppCfgGet = 0, "DebugMode", "ReleseMode") & " => ([" & AppIDGet & "] " & AppNameGet & " ver. " & AppVerGet & " #" & AppDateGet & "#)" & vbLf
End Property

Property Get IsModule() As String
    If cAppType = AnyModule Then IsModule = cModuleName
End Property

Property Let IsModule(Value As String)
    If Value <> vbNullString Then
        cModuleName = Value
        cAppType = AnyModule
    End If
End Property

'public
Sub CheckDBVer()
    Dim tmp As Variant
    Dim pth As String
    Dim fDot As Long
    Dim fSls As Long
    Dim isMod As String
    Dim cnt As Long
    Dim cnt_s As Long
    Dim OldcModuleName As String
    Dim OldcAppType As tAppType
    
    If cVerCheck Then Exit Sub
    
    If cModDic.Count > 0 Then
        OldcAppType = cAppType
        OldcModuleName = cModuleName
        cnt_s = 1
    End If
        
    For cnt = cnt_s To cModDic.Count
            
        If cnt > 0 Then
            cAppType = AnyModule
            cModuleName = cModDic.Item(CStr(cnt))
            isMod = " [module or class]"
        End If
            
        tmp = FindDBVer
        
        If tmp <> vbNullString Then
            tmp = Split(tmp, SEP)
            If tmp(0) <> vbNullString Then
                If AppVerGet <> tmp(0) Then
                    If tmp(1) <> vbNullString Then
                        If MsgBox("There is different version " & AppNameGet & isMod & " (" & tmp(0) & ")" & vbLf & _
                                  "Using old version might cause errors in the macro results" & vbLf & _
                                  "Do you want to open folder with the new version?", vbQuestion + vbYesNo, msgB) = vbYes Then
                            fDot = InStrRev(tmp(1), ".")
                            fSls = InStrRev(tmp(1), "\")
                            If fDot > fSls Then
                                pth = Mid(tmp(1), 1, fSls)
                            Else
                                pth = tmp(1)
                            End If
                            Call Shell("explorer.exe" & " " & pth, vbNormalFocus)
                        End If
                    Else
                        MsgBox "There is different version (" & tmp(0) & ") of " & AppNameGet & isMod & _
                                                                       vbLf & "Pleas contact vbaprogrammers-warsaw@mercer.com team for more detalis.", vbInformation, msgB
                    End If
                End If
                        
                Debug.Print Join(tmp, SEP)
                Debug.Print msgB
            Else
                MsgBox "This application don't have version in database", vbCritical, msgB
            End If
        Else
            Debug.Print "NO_DB_DATA"
            Debug.Print msgB
        End If
                
    Next cnt
    
    cVerCheck = True
    
    If cModDic.Count > 0 Then
        cModuleName = OldcModuleName
        cAppType = OldcAppType
    End If
End Sub

Sub SaveFile(Optional NoVer As Boolean)
    Dim pth As String
    Dim name As String
    Dim ver As String
    Dim fForm As String
        
    name = gVt.AppNameGet
    ver = gVt.AppVerGet
    If ver = SEP Then
        gVt.AppIDSet "0"
        gVt.AppVerSet "0.01"
        gVt.AppCfgSet "0"
        ver = gVt.AppVerGet
    End If
    If name = SEP Or ver = vbNullString Or ver = SEP Then
        MsgBox "Missing AppName or AppVer", vbCritical, msgB
    Else
        If cAppType = ExcelFile Then
            fForm = FileExtStr(cPObj.FileFormat)
            If fForm <> vbNullString Then
                If cPObj.IsAddin = False And NoVer <> True Then
                    pth = cPObj.Path & "\" & name & " " & Format(gVt.AppDateGet, "YYYY.MM.DD") & "_" & ver & fForm
                    cApp.DisplayAlerts = False
                    cPObj.SaveAs pth, cPObj.FileFormat
                    cApp.DisplayAlerts = True
                Else
                    pth = cPObj.Path & "\" & name & "_" & ver & fForm
                    cPObj.SaveCopyAs pth
                    cPObj.Save
                End If
            End If
        ElseIf cAppType = WordFile Then
            fForm = FileExtStr(cPObj.SaveFormat)
            If fForm <> vbNullString Then
                pth = cPObj.Path & "\" & name & " " & Format(gVt.AppDateGet, "YYYY.MM.DD") & "_" & ver & fForm
                cApp.DisplayAlerts = False
                cPObj.SaveAs2 pth, cPObj.SaveFormat
                cApp.DisplayAlerts = True
            End If
        ElseIf cAppType = AccessFile Then
            MsgBox "Don't work for " & cAppType & " files yes", vbCritical, msgB
        End If
    End If
    Debug.Print msgB
End Sub

Sub AppVerMajor()
    Call AppVerPut(True)
End Sub

Sub AppVerMinor()
    Call AppVerPut(False)
End Sub

Sub AppVerSet(ver As String)
    Call AppVerPut(False, ver)
End Sub

Sub AppNameSet(name As String)
    Call SetProp(APP_NAME, name)
    Call SetProp(APP_DATE, Format(Date, "m\/d\/yyyy"))
End Sub

Sub AppIDSet(id As String)
    Call SetProp(APP_ID, id)
    Call SetProp(APP_DATE, Format(Date, "m\/d\/yyyy"))
End Sub

Sub AppCfgSet(cfg As tCfgType)
    Dim tCfg As String
    If cfg = ReleseMode Then
        tCfg = 1
    Else
        tCfg = 0
    End If
    Call SetProp(APP_CFG, tCfg)
    If AppIDGet > 0 And AppCfgGet = ReleseMode Then Call SetDBVer(AppVerGet)
End Sub

'private
Private Sub AppVerPut(verMajor As Boolean, Optional manualVer As String)
    Dim curVer As Variant
    Dim ver As String
    
    ver = AppVerGet
    If ver = SEP And manualVer = vbNullString Then
        ver = "0.01"
        Call SetProp(APP_VER, ver)
    Else
        If manualVer = vbNullString Then
            curVer = Split(ver, ".")
            If verMajor = True Then
                ver = curVer(0) + 1 & ".00"
            Else
                ver = curVer(1) + 1
                If Len(ver) = 1 Then ver = "0" & ver
                ver = curVer(0) & "." & ver
            End If
        Else
            ver = manualVer
        End If
        Call SetProp(APP_VER, ver)
    End If
        
    Call SetProp(APP_DATE, Format(Date, "m\/d\/yyyy"))
    If AppIDGet > 0 And AppCfgGet = ReleseMode Then Call SetDBVer(ver)
    Debug.Print msgB
End Sub

Private Sub SetProp(prop As String, val As String)
    If GetProp(prop) = SEP Then
        If cAppType = ExcelFile Or cAppType = WordFile Then
            cPObj.CustomDocumentProperties.Add _
        name:=prop, _
        LinkToContent:=False, _
        Type:=4, _
        Value:=val                               'msoPropertyTypeString
        ElseIf cAppType = AccessFile Then
            Dim prp As Object
            Set prp = cPObj.CreateProperty(prop, 10, val) 'DB_TEXT
            cPObj.Properties.Append prp
        ElseIf cAppType = AnyModule And VBATrusted Then
            Call VBEModule(cPropDic.Item(prop), cModuleName, val)
        End If
    Else
        If cAppType = ExcelFile Or cAppType = WordFile Then
            cPObj.CustomDocumentProperties(prop).Value = val
        ElseIf cAppType = AccessFile Then
            cPObj.Properties(prop) = val
        ElseIf cAppType = AnyModule And VBATrusted Then
            Call VBEModule(cPropDic.Item(prop), cModuleName, val)
        End If
    End If
End Sub

Private Function GetProp(prop As String) As String
    On Error GoTo err_h
    If cAppType = ExcelFile Or cAppType = WordFile Then
        GetProp = cPObj.CustomDocumentProperties(prop).Value
    ElseIf cAppType = AccessFile Then
        GetProp = cPObj.Properties(prop).Value
    ElseIf cAppType = AnyModule And VBATrusted Then
        GetProp = VBEModule(cPropDic.Item(prop), cModuleName)
    Else
        GetProp = SEP
    End If
    Exit Function
err_h:
    GetProp = SEP
End Function

Private Function ADO_CN(Optional wrt As Boolean) As Object
    Set ADO_CN = CreateObject("ADODB.Connection")
    With ADO_CN
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        If wrt Then
            .Mode = 3                            'adModeReadWrite
        Else
            .Mode = 1                            'adModeRead
        End If
        .CursorLocation = 3                      'adUseClient
        .Properties("Data Source") = cDbPath
        .Properties("Jet OLEDB:Database Password") = cDbPass
        .Open
    End With
End Function

Private Function FindDBVer() As String
    Dim cn As Object
    Dim rs As Object
    
    On Error GoTo err_h
    If AppIDGet = SEP Then Err.Raise 999
    If cDbPath = vbNullString Then Err.Raise 998
    Set cn = ADO_CN
    Set rs = CreateObject("ADODB.RecordSet")
    
    Set rs = cn.Execute("SELECT [AppVersion],[Localization],[Macro_name],[ID] FROM [tblMacroList] WHERE [ID]=" & CLng(AppIDGet))
    If rs.EOF = False Then
        Call AppNameSet(rs!Macro_name)
        FindDBVer = rs!AppVersion & SEP & rs!Localization & SEP & rs!id
    End If
ende:
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
        Set rs = Nothing
    End If
    If Not cn Is Nothing Then
        If cn.State <> 0 Then cn.Close
        Set cn = Nothing
    End If
    Exit Function
err_h:
    If Err.Number = 999 Then
        MsgBox "This project don't have AppID.", vbCritical, msgB
    ElseIf Err.Number = 998 Then
    
    Else
        MsgBox "Error:" & Err.Description, vbCritical, msgB
    End If
    Resume ende
End Function

Private Function SetDBVer(ver As String) As String
    Dim cn As Object
    
    On Error GoTo err_h
    If FindDBVer = vbNullString Then Err.Raise 999
    Set cn = ADO_CN(True)
    cn.Execute ("UPDATE [tblMacroList] SET [AppVersion]='" & ver & "' WHERE [ID]=" & CLng(AppIDGet))

ende:
    If Not cn Is Nothing Then
        If cn.State <> 0 Then cn.Close
        Set cn = Nothing
    End If
        
    Exit Function
err_h:
    If Err.Number = 999 Then
    
    Else
        MsgBox "You aren't connected to MMC network. Macro version can't changed.", vbCritical, msgB
    End If
    Resume ende
End Function

Private Function VBEModule(Optional propType As String, Optional ModuleName As String, Optional propValue As String) As String
    Dim vbproject As Object
    Dim vbComp As Object
    Dim vbMod As Object
    
    Dim projName As String
    Dim prop As Variant
    Dim ln As Long
    Dim cls As String
    Dim txt As String
    Dim tp As String
    Dim sp As String
    Dim sng As Long

    If ModuleName <> vbNullString And propType <> vbNullString Then
        prop = Split(propType, SEP)
        ln = prop(0)
        cls = prop(1)
        tp = prop(2)
        sp = prop(3)
    End If
    If cAppType = AccessFile Then
        projName = cPObj.name
    Else
        projName = cPObj.FullName
    End If
   
    For Each vbproject In cApp.VBE.VBProjects
        If projName = GetProjectPath(vbproject) And vbproject.Protection = 0 Then
            For Each vbComp In vbproject.vbcomponents
                Set vbMod = vbComp.CodeModule
                If vbComp.name = ModuleName Then 'vbext_pp_none
                    If propValue <> vbNullString Then
                        txt = "Private Const " & cls & " As " & tp & " = " & sp & propValue & sp
                            
                        If Mid(vbMod.Lines(ln, 1), 1, 14) & cls = "Private Const " & cls Then
                            vbMod.ReplaceLine ln, txt
                        Else
                            vbMod.InsertLines ln, txt
                        End If
                    Else
                        txt = vbMod.Lines(ln, 1)
                        sng = InStr(1, txt, "=")
                        If sng <> 0 Then
                            VBEModule = Replace(Trim(Mid(txt, sng + 1)), sp, vbNullString)
                        End If
                    End If
                    
        
                    Exit For
                Else
                    txt = "Private Const CLS_ID As Long"
                    If Mid(vbMod.Lines(1, 1), 1, 28) = txt Then
                        txt = "Private Const CLS_SKP As Boolean = True"
                        If Mid(vbMod.Lines(5, 1), 1, 39) <> txt Then
                            cModDic.Add CStr(cModDic.Count + 1), vbMod.name
                        End If
                    End If
                End If
            Next vbComp
            
            Exit For
        End If
    Next vbproject
        
    Set vbproject = Nothing
    Set vbMod = Nothing
    Set vbComp = Nothing

End Function

Private Function FileExtStr(FF As Byte) As String
    If cAppType = ExcelFile Then
        Select Case FF
        Case 51: FileExtStr = ".xlsx"
        Case 52: FileExtStr = ".xlsm"
        Case 55: FileExtStr = ".xlam"
        Case 56: FileExtStr = ".xls"
        Case 50: FileExtStr = ".xlsb"
        End Select
    ElseIf cAppType = WordFile Then
        Select Case FF
        Case 0: FileExtStr = ".doc"
        Case 12: FileExtStr = ".docx"
        Case 13: FileExtStr = ".docm"
        End Select
    ElseIf cAppType = AccessFile Then
        Select Case FF
        Case 0: FileExtStr = ".accdb"
        End Select
    End If
End Function

Private Function GetINI(id As String) As String
    Dim arr As Variant
    Dim ln As String
    Dim file As String
    On Error Resume Next
    
    Select Case cDomain
    Case Is = "EU": file = "c:\Enron\_Prv\_Projekt\_DB\vbateam.ini"
    Case Is = "MERCER": file = "\\teamsites-east.mercer.com\DavWWWRoot\sites\WarsawOpsCenter\Shared Documents\VBATeam\vbateam.ini"
    End Select
    Open file For Input As #1
    
    Do
        Line Input #1, ln
        arr = Split(ln, SEP)
        If arr(0) = id Then
            GetINI = arr(1)
            Exit Do
        End If
    Loop Until EOF(1)
    
    Close #1
End Function
Private Function GetDomain() As String
    GetDomain = Environ("USERDOMAIN")
End Function

Private Function VBATrusted() As Boolean
    On Error Resume Next
    VBATrusted = (cApp.VBE.VBProjects.Count) > 0
    Exit Function
End Function

Private Function FindParentObject(fileName As String) As Object
    Dim temp As Object
    Dim doc As Object
    For Each temp In cApp.Templates
        If temp.vbproject.fileName = fileName Then
            Set FindParentObject = temp
            GoTo ende:
        End If
    Next temp
    For Each doc In cApp.Documents
        If doc.vbproject.fileName = fileName Then
            Set FindParentObject = doc
            GoTo ende:
        End If
    Next doc
        
ende:
End Function

Private Function SetWordObject() As Object
    On Error GoTo ende
    Set SetWordObject = FindParentObject(cApp.VBE.ActiveVBProject.fileName)
    Exit Function
ende:
    Set SetWordObject = cApp.ActiveDocument
End Function

Private Function GetProjectPath(vbproject As Object) As String
    On Error Resume Next
    GetProjectPath = vbproject.fileName
End Function









