VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Private Sub CheckBox1_Click()
    If CheckBox1.Value = True Then
        Worksheets("SFalzen").Range("N4") = Application.InputBox("Bitte Geschwindigkeit für" & vbCrLf & "Bogen A in Bogen/Stunde eingeben:", "Eingabe", , Type:=1)
        Else:
        Worksheets("SFalzen").Range("N4").FormulaLocal = "=WENN(Falzen!A4<>5;SVERWEIS(B106/B104;Falzen!M4:O10;3);F7)"
    End If
End Sub
Private Sub CheckBox2_Click()
    If CheckBox2.Value = True Then
        Worksheets("SFalzen").Range("N5") = Application.InputBox("Bitte Geschwindigkeit für" & vbCrLf & "Bogen B in Bogen/Stunde eingeben:", "Eingabe", , Type:=1)
    Else: Worksheets("SFalzen").Range("N5").FormulaLocal = "=WENN(Falzen!A24<>5;SVERWEIS(C106/C104;Falzen!M4:O10;3);F7)"
    End If
End Sub
Private Sub CheckBox3_Click()
    If CheckBox3.Value = True Then
        Worksheets("SFalzen").Range("N6") = Application.InputBox("Bitte Geschwindigkeit für" & vbCrLf & "Bogen C in Bogen/Stunde eingeben:", "Eingabe", , Type:=1)
    Else: Worksheets("SFalzen").Range("N6").FormulaLocal = "=WENN(Falzen!A43<>5;SVERWEIS(D106/D104;Falzen!M4:O10;3);F7)"
    End If
End Sub
Private Sub CheckBox4_Click()
    If CheckBox4.Value = True Then
        Worksheets("SFalzen").Range("N7") = Application.InputBox("Bitte Geschwindigkeit für" & vbCrLf & "Bogen D in Bogen/Stunde eingeben:", "Eingabe", , Type:=1)
    Else: Worksheets("SFalzen").Range("N7").FormulaLocal = "=WENN(Falzen!A62<>5;SVERWEIS(E106/E104;Falzen!M4:O10;3);F7)"
    End If
End Sub
Private Sub ComboBox5_DropButtonClick()
'Falznutzen
    'Hilfskräfte
    Worksheets("Falzen").Range("D7").FormulaLocal = "=SVERWEIS(SFalzen!J4;SFalzen!G11:L13;6)"
    Worksheets("Falzen").Range("D16").FormulaLocal = "=SVERWEIS(SFalzen!J4;SFalzen!G11:L13;6)"
    'Zuschlag
    Worksheets("Falzen").Range("D8").FormulaLocal = "=SVERWEIS(SFalzen!J4;SFalzen!G11:L13;5)"
    Worksheets("Falzen").Range("D17").FormulaLocal = "=SVERWEIS(SFalzen!J4;SFalzen!G11:L13;5)"
End Sub
Private Sub ComboBox6_DropButtonClick()
'Falznutzen
    'Hilfskräfte
    Worksheets("Falzen").Range("D26").FormulaLocal = "=SVERWEIS(SFalzen!J5;SFalzen!G11:L13;6)"
    Worksheets("Falzen").Range("D35").FormulaLocal = "=SVERWEIS(SFalzen!J5;SFalzen!G11:L13;6)"
    'Zuschlag
    Worksheets("Falzen").Range("D27").FormulaLocal = "=SVERWEIS(SFalzen!J5;SFalzen!G11:L13;5)"
    Worksheets("Falzen").Range("D36").FormulaLocal = "=SVERWEIS(SFalzen!J5;SFalzen!G11:L13;5)"
End Sub
Private Sub ComboBox7_DropButtonClick()
'Falznutzen
    'Hilfskräfte
    Worksheets("Falzen").Range("D45").FormulaLocal = "=SVERWEIS(SFalzen!J6;SFalzen!G11:L13;6)"
    Worksheets("Falzen").Range("D54").FormulaLocal = "=SVERWEIS(SFalzen!J6;SFalzen!G11:L13;6)"
    'Zuschlag
    Worksheets("Falzen").Range("D46").FormulaLocal = "=SVERWEIS(SFalzen!J6;SFalzen!G11:L13;5)"
    Worksheets("Falzen").Range("D55").FormulaLocal = "=SVERWEIS(SFalzen!J6;SFalzen!G11:L13;5)"
End Sub
Private Sub ComboBox8_DropButtonClick()
'Falznutzen
    'Hilfskräfte
    Worksheets("Falzen").Range("D64").FormulaLocal = "=SVERWEIS(SFalzen!J7;SFalzen!G11:L13;6)"
    Worksheets("Falzen").Range("D73").FormulaLocal = "=SVERWEIS(SFalzen!J7;SFalzen!G11:L13;6)"
    'Zuschlag
    Worksheets("Falzen").Range("D65").FormulaLocal = "=SVERWEIS(SFalzen!J7;SFalzen!G11:L13;5)"
    Worksheets("Falzen").Range("D74").FormulaLocal = "=SVERWEIS(SFalzen!J7;SFalzen!G11:L13;5)"
End Sub
'Sub Falzen()
''Falzberechnung aktivieren/deaktivieren
''deaktiviert da bei Änderungen auf d. Eingabeseite immer wieder die 1. Falzmaschine ausgewählt wird
'    'Bogen A
'    'Umschlag falzbar wenn mehr als 4 Seiten
'    If Worksheets("SFalzen").Range("B110").Value = 1 Then
'        If Worksheets("SEingabe").Range("D5").Value = 0 Or Worksheets("SFalzen").Range("B103").Value > 4 Then
'            Worksheets("Falzen").ComboBox1.ListFillRange = "SFalzen!B3:C8" 'Auswahl Falzmaschine
'            Worksheets("Falzen").ComboBox1.ListIndex = 0 'Ersten Wert d. Auswahl autom. anzeigen
'            Worksheets("Falzen").ComboBox1.LinkedCell = "SFalzen!I4" 'Ausgabe Falzmaschine
'            Worksheets("SFalzen").Range("I4").Value = 1 'Standardwert
'            Else
'            Worksheets("Falzen").ComboBox1.ListFillRange = "SFalzen!B8:C8" 'Auswahl Falzmaschine
'            Worksheets("Falzen").ComboBox1.ListIndex = 0 'Ersten Wert d. Auswahl autom. anzeigen
'            Worksheets("Falzen").ComboBox1.LinkedCell = "SFalzen!I4" 'Ausgabe Falzmaschine
'            Worksheets("SFalzen").Range("I4").Value = 6 'nicht Falzen
'        End If
'    End If
'    'Bogen B
'    'Umschlag falzbar wenn mehr als 4 Seiten
'    If Worksheets("SFalzen").Range("C110").Value = 1 Then
'       If Worksheets("SEingabe").Range("E5").Value = 0 Or Worksheets("SFalzen").Range("C103").Value > 4 Then
'            Worksheets("Falzen").ComboBox2.ListFillRange = "SFalzen!B3:C8" 'Auswahl Falzmaschine
'            Worksheets("Falzen").ComboBox2.ListIndex = 0 'Ersten Wert d. Auswahl autom. anzeigen
'            Worksheets("Falzen").ComboBox2.LinkedCell = "SFalzen!I5" 'Ausgabe Falzmaschine
'            Worksheets("SFalzen").Range("I5").Value = 1 'Standardwert
'            Else
'            Worksheets("Falzen").ComboBox2.ListFillRange = "SFalzen!B8:C8" 'Auswahl Falzmaschine
'            Worksheets("Falzen").ComboBox2.ListIndex = 0 'Ersten Wert d. Auswahl autom. anzeigen
'            Worksheets("Falzen").ComboBox2.LinkedCell = "SFalzen!I5" 'Ausgabe Falzmaschine
'            Worksheets("SFalzen").Range("I5").Value = 6 'nicht Falzen
'        End If
'     End If
'   'Bogen C
'   'Umschlag falzbar wenn mehr als 4 Seiten
'    If Worksheets("SFalzen").Range("D110").Value = 1 Then
'        If Worksheets("SEingabe").Range("F5").Value = 0 Or Worksheets("SFalzen").Range("D103").Value > 4 Then
'            Worksheets("Falzen").ComboBox3.ListFillRange = "SFalzen!B3:C8" 'Auswahl Falzmaschine
'            Worksheets("Falzen").ComboBox3.ListIndex = 0 'Ersten Wert d. Auswahl autom. anzeigen
'            Worksheets("Falzen").ComboBox3.LinkedCell = "SFalzen!I6" 'Ausgabe Falzmaschine
'            Worksheets("SFalzen").Range("I6").Value = 1 'Standardwert
'            Else
'            Worksheets("Falzen").ComboBox3.ListFillRange = "SFalzen!B8:C8" 'Auswahl Falzmaschine
'            Worksheets("Falzen").ComboBox3.ListIndex = 0 'Ersten Wert d. Auswahl autom. anzeigen
'            Worksheets("Falzen").ComboBox3.LinkedCell = "SFalzen!I6" 'Ausgabe Falzmaschine
'            Worksheets("SFalzen").Range("I6").Value = 6 'nicht Falzen
'        End If
'    End If
'    'Bogen D
'    'Umschlag falzbar wenn mehr als 4 Seiten
'    If Worksheets("SFalzen").Range("E110").Value = 1 Then
'        If Worksheets("SEingabe").Range("G5").Value = 0 Or Worksheets("SFalzen").Range("E103").Value > 4 Then
'            Worksheets("Falzen").ComboBox4.ListFillRange = "SFalzen!B3:C8" 'Auswahl Falzmaschine
'            Worksheets("Falzen").ComboBox4.ListIndex = 0 'Ersten Wert d. Auswahl autom. anzeigen
'            Worksheets("Falzen").ComboBox4.LinkedCell = "SFalzen!I7" 'Ausgabe Falzmaschine
'            Worksheets("SFalzen").Range("I7").Value = 1 'Standardwert
'            Else
'            Worksheets("Falzen").ComboBox4.ListFillRange = "SFalzen!B8:C8" 'Auswahl Falzmaschine
'            Worksheets("Falzen").ComboBox4.ListIndex = 0 'Ersten Wert d. Auswahl autom. anzeigen
'            Worksheets("Falzen").ComboBox4.LinkedCell = "SFalzen!I7" 'Ausgabe Falzmaschine
'            Worksheets("SFalzen").Range("I7").Value = 6 'nicht Falzen
'        End If
'    End If
'End Sub
Sub Falzbogen()
'unused
' Loeschen d. vorhandenen Eingaben u. Ausblendung, wenn d. Bg. nicht ausgewaehlt wurde
'
If Worksheets("Eingabe").Range("A18") < 1 Then
    Worksheets("Falzen").Rows("23:41").Hidden = True
    Else: Worksheets("Falzen").Rows("23:41").Hidden = False
End If
If Worksheets("Eingabe").Range("A27") < 1 Then
    Worksheets("Falzen").Rows("42:60").Hidden = True
    Else: Worksheets("Falzen").Rows("42:60").Hidden = False
End If
If Worksheets("Eingabe").Range("A36") < 1 Then
    Worksheets("Falzen").Rows("61:79").Hidden = True
    Else: Worksheets("Falzen").Rows("61:79").Hidden = False
End If
End Sub
