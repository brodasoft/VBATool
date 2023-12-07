Attribute VB_Name = "mVBEMemory"
Option Explicit

Private Const PAGE_EXECUTE_READWRITE As Variant = &H40

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
                    (Destination As Long, Source As Long, ByVal Length As Long)
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Long, _
                                                        ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) As Long
Private Declare Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As Long, _
                                                                              ByVal pTemplateName As Long, ByVal hWndParent As Long, _
                                                                              ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer
Public HookBytes(0 To 5) As Byte
Public OriginBytes(0 To 5) As Byte
Public pFunc As Long
Public Flag As Boolean

Public Sub VBEUnlockSTOP()
    If Flag Then
        MoveMemory ByVal pFunc, ByVal VarPtr(OriginBytes(0)), 6
        Debug.Print "VBEUnlock stoped"
    End If
End Sub

Public Function VBEUnlock() As Boolean
    Dim TmpBytes(0 To 5) As Byte
    Dim p As Long
    Dim OriginProtect As Long

    VBEUnlock = False

    pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")

    If VirtualProtect(ByVal pFunc, 6, PAGE_EXECUTE_READWRITE, OriginProtect) <> 0 Then

        MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal pFunc, 6
        If TmpBytes(0) <> &H68 Then

            MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal pFunc, 6

            p = GetPtr(AddressOf MyDialogBoxParam)
            
            HookBytes(0) = &H68
            MoveMemory ByVal VarPtr(HookBytes(1)), ByVal VarPtr(p), 4
            HookBytes(5) = &HC3

            MoveMemory ByVal pFunc, ByVal VarPtr(HookBytes(0)), 6
            Flag = True
            VBEUnlock = True
            
        End If
    End If
End Function

Private Function MyDialogBoxParam(ByVal hInstance As Long, _
                                  ByVal pTemplateName As Long, ByVal hWndParent As Long, _
                                  ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer
    On Error GoTo err_h
    
    If pTemplateName = 4070 Then
        MyDialogBoxParam = 1
    Else
        Call VBEUnlockSTOP
        MyDialogBoxParam = DialogBoxParam(hInstance, pTemplateName, hWndParent, lpDialogFunc, dwInitParam)
        Call VBEUnlock
    End If
    
err_h:

End Function

Private Function GetPtr(ByVal Value As Long) As Long
    GetPtr = Value
End Function


