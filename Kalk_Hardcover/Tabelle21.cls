VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


'
'Copyright by E.Dargel ed@dissenter.de
'
Private Sub Worksheet_Activate()
    On Error Resume Next
    Call produkt
    Call Kommentar
    Run Worksheets("SVerpacken").draw_all
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    Call Kommentar
'    Call lkw
End Sub
Private Sub ToggleButton1_Click()
    '20140623
    'On Error Resume Next
    Dim intGewicht, intPaletten As Integer
    intGewicht = Worksheets("SVerpacken").Range("K53")
    intPaletten = Worksheets("SVerpacken").Range("K52")
    Worksheets("Verpacken").ToggleButton1.Caption = "Soll: " & IIf(ToggleButton1.Value = True, intPaletten & " Paletten", (intGewicht / 1000) & "t Frachtgewicht")
    'Worksheets("Verpacken").ToggleButton1.BackColor = RGB(255, 153, 0)
    Call lkw
End Sub
Private Sub CheckBox1_Click()
    'Paletten ja/nein
    On Error Resume Next
    Call Kommentar
    Run Worksheets("SVerpacken").draw_all
End Sub
Private Sub CheckBox4_Click()
    'Eingabe manuelle Lagen ja/nein
    '20140623
    On Error Resume Next
    If Worksheets("Verpacken").CheckBox4.Value = True Then
        Worksheets("Verpacken").Range("D31").Interior.ColorIndex = 6
        Worksheets("Verpacken").Range("D31").Interior.Pattern = xlSolid
        Worksheets("Verpacken").Range("D31").Locked = False
        'Grafik löschen
        Dim shpShape As Shape
            For Each shpShape In Worksheets("Verpacken").Shapes
                If shpShape.Type = msoAutoShape Then
                    shpShape.Delete
                End If
            Next shpShape
        Else
        Worksheets("Verpacken").Range("D31").Value = 0
        Worksheets("Verpacken").Range("D31").Interior.ColorIndex = 0
        Worksheets("Verpacken").Range("D31").Interior.Pattern = xlSolid
        Worksheets("Verpacken").Range("D31").Locked = True
        Run Worksheets("SVerpacken").VE_einschweissen
        Run Worksheets("SVerpacken").Palettenlage_erstellen
        Run Worksheets("SVerpacken").draws_delete
        Run Worksheets("SVerpacken").VE_erstellen
    End If
    MsgBox "Hinweis: Format A4 = 12 VE/Lage."
    Call Kommentar
End Sub
Private Sub ComboBox1_DropButtonClick()
    'Palettenauswahl
    '20140623
    On Error Resume Next
    Dim intRow As Integer
    intRow = Worksheets("Material").Range("D78")
    Worksheets("Verpacken").ComboBox1.ListRows = intRow
    Worksheets("Verpacken").ComboBox1.ListFillRange = "Material!C79:C" & (79 + intRow - 1)
    Call Kommentar
    Call Palettencheck
    Run Worksheets("SVerpacken").draw_all
End Sub
Private Sub ComboBox2_Change()
    'Verpackungsart
    On Error Resume Next
    Call Verpackungseinheit
    Run Worksheets("SVerpacken").draw_all
    Call Kommentar
End Sub
Private Sub ComboBox3_Change()
    'Palettendeckelauswahl
    On Error Resume Next
    Call Kommentar
End Sub
Private Sub ComboBox7_Change()
    'Auswahl der Einschweißmaschine
    '20090824
    'Standard V eintragen
    On Error Resume Next
    If Worksheets("SVerpacken").Range("B71") > 1 Then
        Worksheets("Verpacken").Range("D65").FormulaLocal = "=SVERWEIS(SVerpacken!B71;SVerpacken!A72:I75;9)"
        'MsgBox "Achtung, die Maschinengeschwindigkeit wurde zurückgesetzt!"
        Run Worksheets("SVerpacken").draw_all
    End If
End Sub
Private Sub ComboBox8_Change()
    'Auswahl Kantenschutz
    Call Kommentar
End Sub
Private Sub ComboBox9_Change()
    'Auswahl Umreifung
    Call Kommentar
End Sub
Private Sub ComboBox13_Change()
    'Banderole zeichnen
    On Error Resume Next
    Run Worksheets("SVerpacken").draw_all
End Sub
Private Sub ComboBox14_Change()
    'Banderole zeichnen
    On Error Resume Next
    Run Worksheets("SVerpacken").draw_all
End Sub
Private Sub ComboBox16_Change()
    'VE Lage Einschweißen
    '20090820
    On Error Resume Next
    Run Worksheets("SVerpacken").draw_all
End Sub
Private Sub ComboBox17_Change()
    'VE Lage Einschweißen
    '20090820
    On Error Resume Next
    Run Worksheets("SVerpacken").draw_all
End Sub
Private Sub CommandButton1_Click()
    'Reset Ex./VE auf Standard
    '20140711
    If Worksheets("Material").Range("C121") < 14 Then
        Worksheets("Verpacken").Range("D6").FormulaLocal = "=WENN(ODER(Material!C121=14;Material!C121=13;Material!C121=12);1;SVerpacken!G27)"
    End If
    Call Kommentar
End Sub
Private Sub CommandButton2_Click()
    'Materialseite einblenden
    On Error Resume Next
    Worksheets("Material").Visible = xlSheetVisible
    Worksheets("Material").Activate
    Worksheets("Material").Range("B74").Select
End Sub
Sub Verpackungseinheit()
    'Festlegen Verpackungseinheit u. Einschweißen
    '20141223
    'On Error Resume Next
    Worksheets("Verpacken").Range("D6").ClearComments
    Worksheets("Verpacken").Range("D6").Locked = False
    Worksheets("Verpacken").Range("D6").FormulaHidden = False
    Worksheets("Verpacken").Range("D6").Interior.ColorIndex = 6
    Worksheets("Verpacken").Range("D6").Interior.Pattern = xlSolid
    'Standard
    If Worksheets("Material").Range("C121") < 14 Then
        If Worksheets("Verpacken").Range("D6").HasFormula = False Then
       ' Prüfen ob Expl. durch Benutzer verändert wurden
        MsgBox "Achtung!" & Chr(10) & "Die Anzahl der Exemplare/VE wurde manuell eingetragen." _
        & Chr(10) & "Für die Automatische Berechnung bitte Reset drücken."
        End If
    End If
    'Einschweißen
    If Worksheets("Material").Range("C121") = 14 Then
        'Ja mit Vorauswahl
        With Worksheets("Verpacken").ComboBox7
            .ListIndex = 1
        End With
   End If
    'Abstapeln v. Einzelexemplaren ohne Einschweißen
    If Worksheets("Material").Range("C121") = 15 Then
        With Worksheets("Verpacken").ComboBox7
            .ListIndex = 0
        End With
        Worksheets("Verpacken").Range("D6") = 1
        Worksheets("Verpacken").Range("D6").Interior.ColorIndex = 0
        Worksheets("Verpacken").Range("D6").Locked = True
        Worksheets("Verpacken").Range("D6").FormulaHidden = False
        Worksheets("Verpacken").Range("D6").AddComment
        Worksheets("Verpacken").Range("D6").Comment.Visible = False
        Worksheets("Verpacken").Range("D6").Comment.Text Text:="Einfaches Abstapeln ohne Folie od. Karton." & Chr(10) & _
        "Bei Bedarf Zwischenlagen festlegen."
    End If
    Call Kommentar
End Sub
Sub lkw()
    'Berechnung der benötigten Lkws
    '20140623
    'On Error Resume Next
    Dim intGewicht, intOptPaletten, intPaletten As Integer
    Dim strActiveCell As String
    intGewicht = Worksheets("SVerpacken").Range("K53")
    intGewichtIst = Worksheets("SVerpacken").Range("L54")
    intOptPaletten = Worksheets("SVerpacken").Range("G52")
    intPaletten = Worksheets("SVerpacken").Range("K52")
    strActiveCell = ActiveCell.Address
    If Worksheets("Verpacken").ToggleButton1.Value = True Then
        Worksheets("SVerpacken").Range("K55").Value = 1
        If Worksheets("SVerpacken").Range("K54") > 0 Then
            MsgBox "Achtung! Das Frachtgewicht von " & intGewicht / 1000 & " t wurde um " & intGewichtIst & "t überschritten." & vbCrLf & vbCrLf & _
            "Es können maximal nur " & intOptPaletten & " Paletten je Lkw transportiert werden."
            Worksheets("Verpacken").Range("C40:D40,F40:G40").Interior.ColorIndex = 3
            Worksheets("Verpacken").Range("C40:D40,F40:G40").Interior.Pattern = xlSolid
            Worksheets("Verpacken").Range("C40:D40,F40:G40").Locked = False
            Else
            Worksheets("Verpacken").Range("C40:D40,F40:G40").Interior.ColorIndex = 0
            Worksheets("Verpacken").Range("C40:D40,F40:G40").Interior.Pattern = xlSolid
            Worksheets("Verpacken").Range("C40:D40,F40:G40").Locked = False
        End If
        Else
        Worksheets("Verpacken").Range("C40:D40,F40:G40").Interior.ColorIndex = 0
        Worksheets("Verpacken").Range("C40:D40,F40:G40").Interior.Pattern = xlSolid
        Worksheets("Verpacken").Range("C40:D40,F40:G40").Locked = False
        Worksheets("SVerpacken").Range("K55").Value = 0
    End If
End Sub
Sub Palettencheck()
    'Palettenwerte setzen
    'modified 20140623
    'On Error Resume Next
    Dim intAuswahlPl As Integer
    intAuswahlPl = Worksheets("SVerpacken").Range("H7")
    If intAuswahlPl = 8 Then
        Worksheets("Verpacken").Range("D23").FormulaLocal = "=(SVERWEIS(SVerpacken!H7;Material!B79:J93;9)+SVERWEIS(SVerpacken!H7;Material!B79:J93;5))/10"
        Worksheets("Verpacken").Range("F23") = "(Palettenvorgabe)"
        Worksheets("Verpacken").Range("D23").Interior.ColorIndex = 0
        Worksheets("Verpacken").Range("D23").Interior.Pattern = xlSolid
    Else
        Worksheets("Verpacken").Range("D23") = 120
        Worksheets("Verpacken").Range("F23") = ""
        Worksheets("Verpacken").Range("D23").Interior.ColorIndex = 6
        Worksheets("Verpacken").Range("D23").Interior.Pattern = xlSolid
    End If
    Call Kommentar
End Sub
Sub Kommentar()
    'Anzeigen d. Palettendaten
    '201401002
    'On Error Resume Next
    Dim VPalette, VExemplare As String
    VPalette = Worksheets("SVerpacken").Range("I39")
    VExemplare = Worksheets("SVerpacken").Range("G45")
    Worksheets("Verpacken").Label2.Caption = "Volle Einzelpalette:" & Chr(10) & VPalette & " kg" & Chr(10) & VExemplare & " Ex."
End Sub

