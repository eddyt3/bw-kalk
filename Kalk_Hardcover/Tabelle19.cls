VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Private Sub CheckBox1_Click()
'manuelle V Bg. A
    If Worksheets("Falzen").CheckBox1.Value = True Then
        Worksheets("SFalzen").Range("N4") = Application.InputBox("Bitte Geschwindigkeit für" & vbCrLf & "Bogen A in Bogen/Stunde eingeben:", "Eingabe", , Type:=1)
    Else: Worksheets("SFalzen").Range("N4").FormulaLocal = "=WENN(Falzen!A4<>5;SVERWEIS(B106/B104;Falzen!M4:O10;3);F7)"
    End If
End Sub
Private Sub CheckBox2_Click()
'manuelle V Bg. B
    If Worksheets("Falzen").CheckBox2.Value = True Then
        Worksheets("SFalzen").Range("N5") = Application.InputBox("Bitte Geschwindigkeit für" & vbCrLf & "Bogen B in Bogen/Stunde eingeben:", "Eingabe", , Type:=1)
    Else: Worksheets("SFalzen").Range("N5").FormulaLocal = "=WENN(Falzen!A24<>5;SVERWEIS(C106/C104;Falzen!M4:O10;3);F7)"
    End If
End Sub
Private Sub CheckBox3_Click()
'manuelle V Bg. C
    If Worksheets("Falzen").CheckBox3.Value = True Then
        Worksheets("SFalzen").Range("N6") = Application.InputBox("Bitte Geschwindigkeit für" & vbCrLf & "Bogen C in Bogen/Stunde eingeben:", "Eingabe", , Type:=1)
    Else: Worksheets("SFalzen").Range("N6").FormulaLocal = "=WENN(Falzen!A43<>5;SVERWEIS(D106/D104;Falzen!M4:O10;3);F7)"
    End If
End Sub
Private Sub CheckBox4_Click()
'manuelle V Bg. D
    If Worksheets("Falzen").CheckBox4.Value = True Then
        Worksheets("SFalzen").Range("N7") = Application.InputBox("Bitte Geschwindigkeit für" & vbCrLf & "Bogen D in Bogen/Stunde eingeben:", "Eingabe", , Type:=1)
    Else: Worksheets("SFalzen").Range("N7").FormulaLocal = "=WENN(Falzen!A62<>5;SVERWEIS(E106/E104;Falzen!M4:O10;3);F7)"
    End If
End Sub
Private Sub ComboBox5_DropButtonClick()
    'Falznutzen
    'Hilfskräfte
    Worksheets("Falzen").Range("D7").FormulaLocal = "=SVERWEIS(SFalzen!J4;SFalzen!G11:L13;6)"
    Worksheets("Falzen").Range("D16").FormulaLocal = "=SVERWEIS(SFalzen!J4;SFalzen!G11:L13;6)"
    Worksheets("Falzen").Range("D8").FormulaLocal = "=SVERWEIS(SFalzen!J4;SFalzen!G11:L13;5)"
    Worksheets("Falzen").Range("D17").FormulaLocal = "=SVERWEIS(SFalzen!J4;SFalzen!G11:L13;5)"
End Sub
Private Sub ComboBox6_DropButtonClick()
    'Falznutzen
    'Hilfskräfte
    Worksheets("Falzen").Range("D26").FormulaLocal = "=SVERWEIS(SFalzen!J5;SFalzen!G11:L13;6)"
    Worksheets("Falzen").Range("D35").FormulaLocal = "=SVERWEIS(SFalzen!J5;SFalzen!G11:L13;6)"
    Worksheets("Falzen").Range("D27").FormulaLocal = "=SVERWEIS(SFalzen!J5;SFalzen!G11:L13;5)"
    Worksheets("Falzen").Range("D36").FormulaLocal = "=SVERWEIS(SFalzen!J5;SFalzen!G11:L13;5)"
End Sub
Private Sub ComboBox7_DropButtonClick()
    'Falznutzen
    'Hilfskräfte
    Worksheets("Falzen").Range("D45").FormulaLocal = "=SVERWEIS(SFalzen!J6;SFalzen!G11:L13;6)"
    Worksheets("Falzen").Range("D54").FormulaLocal = "=SVERWEIS(SFalzen!J6;SFalzen!G11:L13;6)"
    Worksheets("Falzen").Range("D46").FormulaLocal = "=SVERWEIS(SFalzen!J6;SFalzen!G11:L13;5)"
    Worksheets("Falzen").Range("D55").FormulaLocal = "=SVERWEIS(SFalzen!J6;SFalzen!G11:L13;5)"
End Sub
Private Sub ComboBox8_DropButtonClick()
    'Falznutzen
    'Hilfskräfte
    Worksheets("Falzen").Range("D64").FormulaLocal = "=SVERWEIS(SFalzen!J7;SFalzen!G11:L13;6)"
    Worksheets("Falzen").Range("D73").FormulaLocal = "=SVERWEIS(SFalzen!J7;SFalzen!G11:L13;6)"
    Worksheets("Falzen").Range("D65").FormulaLocal = "=SVERWEIS(SFalzen!J7;SFalzen!G11:L13;5)"
    Worksheets("Falzen").Range("D74").FormulaLocal = "=SVERWEIS(SFalzen!J7;SFalzen!G11:L13;5)"
End Sub
Sub Falzbogen()
'unused
    ' Loeschen d. vorhandenen Eingaben u. Ausblendung, wenn d. Bg. nicht ausgewaehlt wurde
    '
    If Worksheets("Eingabe").Range("A17") < 1 Then
        Worksheets("Falzen").Rows("23:41").Hidden = True
    Else: Worksheets("Falzen").Rows("23:41").Hidden = False
    End If
    If Worksheets("Eingabe").Range("A26") < 1 Then
        Worksheets("Falzen").Rows("42:60").Hidden = True
    Else: Worksheets("Falzen").Rows("42:60").Hidden = False
    End If
    If Worksheets("Eingabe").Range("A35") < 1 Then
        Worksheets("Falzen").Rows("61:79").Hidden = True
    Else: Worksheets("Falzen").Rows("61:79").Hidden = False
    End If
End Sub
Sub Falzen_Hardcover()
'20190116
'Nur für Hardcover!
    Dim intFalzen As Boolean
    intFalzen = ThisWorkbook.Worksheets("Eingabe").CheckBox5.Value
    If intFalzen = True Then
        If Worksheets("SEingabe").Range("D3") = 1 Then 'Bg. A Inhalt
            ThisWorkbook.Worksheets("Falzen").ComboBox1.ListIndex = 0
        End If
        If Worksheets("SEingabe").Range("D4") = 1 Then 'Bg. A Umschlag nicht Falzen
            ThisWorkbook.Worksheets("Falzen").ComboBox1.ListIndex = 6
        End If
        If Worksheets("SEingabe").Range("D5") = 0 Then 'Bg. A nicht Falzen
            ThisWorkbook.Worksheets("Falzen").ComboBox1.ListIndex = 6
        End If
        If Worksheets("Eingabe").CheckBox1.Value = True Then 'Bg. B ja/nein
            If Worksheets("SEingabe").Range("e3") = 1 Then
                ThisWorkbook.Worksheets("Falzen").ComboBox2.ListIndex = 0
            End If
            If Worksheets("SEingabe").Range("e4") = 1 Then
                ThisWorkbook.Worksheets("Falzen").ComboBox2.ListIndex = 6
            End If
            If Worksheets("SEingabe").Range("e5") = 0 Then
                ThisWorkbook.Worksheets("Falzen").ComboBox2.ListIndex = 6
            End If
        Else: ThisWorkbook.Worksheets("Falzen").ComboBox2.ListIndex = 6
        End If
        If Worksheets("Eingabe").CheckBox2.Value = True Then 'Bg. C ja/nein
            If Worksheets("SEingabe").Range("f3") = 1 Then
                ThisWorkbook.Worksheets("Falzen").ComboBox3.ListIndex = 0
            End If
            If Worksheets("SEingabe").Range("f4") = 1 Then
                ThisWorkbook.Worksheets("Falzen").ComboBox3.ListIndex = 6
            End If
            If Worksheets("SEingabe").Range("f5") = 0 Then
                ThisWorkbook.Worksheets("Falzen").ComboBox3.ListIndex = 6
            End If
        Else: ThisWorkbook.Worksheets("Falzen").ComboBox3.ListIndex = 6
        End If
        If Worksheets("Eingabe").CheckBox3.Value = True Then 'Bg. D ja/nein
            If Worksheets("SEingabe").Range("g3") = 1 Then
                ThisWorkbook.Worksheets("Falzen").ComboBox4.ListIndex = 0
            End If
            If Worksheets("SEingabe").Range("g4") = 1 Then
                ThisWorkbook.Worksheets("Falzen").ComboBox4.ListIndex = 6
            End If
            If Worksheets("SEingabe").Range("g5") = 0 Then
                ThisWorkbook.Worksheets("Falzen").ComboBox4.ListIndex = 6
            End If
        Else: ThisWorkbook.Worksheets("Falzen").ComboBox4.ListIndex = 6
        End If
    Else: ThisWorkbook.Worksheets("Falzen").ComboBox1.ListIndex = 6
    End If
End Sub
