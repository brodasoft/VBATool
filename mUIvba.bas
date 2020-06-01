Attribute VB_Name = "mUIvba"
Option Explicit

Sub tworzenie_listy()
    Dim a As String
    Dim adr As String
    Dim kierunek As String
    Dim info  As String
    Dim znacznik As String
    Dim Text  As String
    Dim il As Long


    a = Selection.Cells(1, 1).Address
    If Mid$(a, 3, 1) = "$" Then
        adr = Mid$(a, 2, 1) & Mid$(a, 4, 5)

        adr = Mid$(a, 2, 2) & Mid$(a, 5, 5)
    End If
    If MsgBox("Sortowaæ?", vbYesNo + vbQuestion, "Sort") = vbYes Then
        Selection.Sort Key1:=ActiveSheet.Range(adr), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    End If
jez:
    kierunek = InputBox("Podaj kierunek ³aczenia" & Chr$(13) & "v - pionowo" & Chr$(13) & "h - poziomo", "Orientacja", "v")
    If kierunek <> "v" And kierunek <> "h" Then
        info = MsgBox("B³¹dne oznaczenie kierunku" & Chr$(13) & "Chcesz kontynuowaæ ???", vbYesNo, "B³¹d")
        If info = vbYes Then
            GoTo jez
        Else: Exit Sub
        End If
    End If
    znacznik = InputBox("Podaj znacznik rozdzielaj¹cy" & Chr$(13) & "Puste - brak znacznika", "Znacznik", ";")

    Text = Empty

    For il = 1 To Selection.Count
        If kierunek = "v" Then

            If Text = Empty Then
                Text = Text & Selection.Cells(0 + il, 1)
            Else
                Text = Text & znacznik & Selection.Cells(0 + il, 1)
            End If

        Else
            If Text = Empty Then
                Text = Text & Selection.Cells(1, 0 + il)
            Else
                Text = Text & znacznik & Selection.Cells(1, 0 + il)
            End If

        End If

    Next il

    If kierunek = "v" Then
        Selection.Cells(0 + il, 1) = Text
    Else
        Selection.Cells(1, 0 + il) = Text

    End If



End Sub

Sub auto_filtr()
    On Error Resume Next
    Selection.AutoFilter
End Sub

Public Sub zmiana_liczba_tekst()
    Dim pyt As Long
    Dim c As Long
    Dim ca As Long
    Dim r As Long
    Dim ra As Long
    Dim l As Long
    Dim k As Long
    Dim b As String
    Dim z As String

    Application.Calculation = xlCalculationManual
    pyt = MsgBox("Tak - zamieñ na liczbê mno¿¹c przez 1" & Chr$(13) & "Nie - zamieñ na tekst przez dodanie '" & Chr$(13) & "Anuluj - usuñ '", vbYesNoCancel, "TRYB (sbroda converter)")

    b = Selection.AddressLocal
    z = Selection.Address

    If pyt = vbYes Then
        c = Selection.Column
        ca = Selection.Columns.Count
        r = Selection.Row
        ra = Selection.Rows.Count
    
        For l = r To r + ra - 1
            For k = c To c + ca - 1
                If ActiveSheet.Cells(l, k) <> vbNullString And IsNumeric(ActiveSheet.Cells(l, k)) Then ActiveSheet.Cells(l, k) = CDbl(ActiveSheet.Cells(l, k)) * 1
        
            Next k
    
        Next l

    End If

    If pyt = vbNo Then
        c = Selection.Column
        ca = Selection.Columns.Count
        r = Selection.Row
        ra = Selection.Rows.Count
    
        For l = r To r + ra - 1
            For k = c To c + ca - 1
                If ActiveSheet.Cells(l, k) <> vbNullString Then ActiveSheet.Cells(l, k) = "'" & ActiveSheet.Cells(l, k)
        
            Next k
    
        Next l

    End If

    If pyt = vbCancel Then
        c = Selection.Column
        ca = Selection.Columns.Count
        r = Selection.Row
        ra = Selection.Rows.Count
    
        For l = r To r + ra - 1
            For k = c To c + ca - 1
                If ActiveSheet.Cells(l, k) <> vbNullString Then ActiveSheet.Cells(l, k) = ActiveSheet.Cells(l, k)
        
            Next k
    
        Next l

    End If
    
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub znajdz_pustr()
    Selection.SpecialCells(xlCellTypeBlanks).Select


End Sub

Sub znajdz_bledne()
    On Error Resume Next
    Selection.SpecialCells(xlCellTypeFormulas, 16).Select


End Sub

Sub wklej_formuly()

   
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
                           SkipBlanks:=False, Transpose:=False
   

End Sub


