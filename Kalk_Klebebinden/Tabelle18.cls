VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle18"
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
    Call Worksheets("SVerpacken").draw_all
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    Call Kommentar
'    Call lkw
End Sub
Private Sub ToggleButton1_Click()
    On Error Resume Next
    Dim intGewicht, intPaletten As Integer
    intGewicht = Worksheets("SVerpacken").Range("K53")
    intPaletten = Worksheets("SVerpacken").Range("K52")
    With ToggleButton1
    .Caption = "Soll: " & IIf(.Value = True, intPaletten & " Paletten", (intGewicht / 1000) & "t Frachtgewicht")
    '.BackColor = RGB(255, 153, 0)
    End With
    Call lkw
End Sub
Private Sub CheckBox1_Click()
    'Paletten ja/nein
    On Error Resume Next
    Call Kommentar
    Call Worksheets("SVerpacken").draw_all
End Sub
Private Sub CheckBox4_Click()
    'Eingabe manuelle Lagen ja/nein
    On Error Resume Next
    Worksheets("Verpacken").Unprotect "bw"
    If CheckBox4.Value = True Then
        ActiveSheet.Range("D31").Select
        With Selection.Interior
            .ColorIndex = 6
            .Pattern = xlSolid
        End With
        Selection.Locked = False
        'Grafik löschen
        Dim shpShape As Shape
            For Each shpShape In ActiveSheet.Shapes
                If shpShape.Type = msoAutoShape Then
                    shpShape.Delete
                End If
            Next shpShape
        Else
        ActiveSheet.Range("D31").Select
        Selection.Value = 0
        With Selection.Interior
            .ColorIndex = 0
            .Pattern = xlSolid
        End With
        Selection.Locked = True
        Worksheets("Verpacken").Protect "bw"
        Call Worksheets("SVerpacken").VE_einschweissen
        Call Worksheets("SVerpacken").Palettenlage_erstellen
        Call Worksheets("SVerpacken").draws_delete
        Call Worksheets("SVerpacken").VE_erstellen
    End If
    MsgBox "Hinweis: Format A4 = 12 VE/Lage."
    Call Kommentar
End Sub
Private Sub ComboBox1_DropButtonClick()
    'Palettenauswahl
    On Error Resume Next
    Dim intRow As Integer
    intRow = Worksheets("Material").Range("D78")
    With ComboBox1
        .ListRows = intRow
        .ListFillRange = "Material!C79:C" & (79 + intRow - 1)
    End With
    Call Kommentar
    Call Palettencheck
    Call Worksheets("SVerpacken").draw_all
End Sub
Private Sub ComboBox2_Change()
    'Verpackungsart
    On Error Resume Next
    Call Verpackungseinheit
    Call Worksheets("SVerpacken").draw_all
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
        Worksheets("Verpacken").Unprotect "bw"
        Worksheets("Verpacken").Range("D65").FormulaLocal = "=SVERWEIS(SVerpacken!B71;SVerpacken!A72:I75;9)"
        'MsgBox "Achtung, die Maschinengeschwindigkeit wurde zurückgesetzt!"
        Call Worksheets("SVerpacken").draw_all
        Worksheets("Verpacken").Protect "bw"
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
    Call Worksheets("SVerpacken").draw_all
End Sub
Private Sub ComboBox14_Change()
    'Banderole zeichnen
    On Error Resume Next
    Call Worksheets("SVerpacken").draw_all
End Sub
Private Sub ComboBox16_Change()
    'VE Lage Einschweißen
    '20090820
    On Error Resume Next
    Call Worksheets("SVerpacken").draw_all
End Sub
Private Sub ComboBox17_Change()
    'VE Lage Einschweißen
    '20090820
    On Error Resume Next
    Call Worksheets("SVerpacken").draw_all
End Sub
Private Sub CommandButton1_Click()
    'Reset Ex./VE auf Standard
    '20140319
    On Error Resume Next
    If Worksheets("Material").Range("C121") < 14 Then
        Worksheets("Verpacken").Range("D6").FormulaLocal = "=WENN(ODER(Material!C121=14;Material!C121=13;Material!C121=12);1;ABRUNDEN(SVerpacken!D8/(SVerpacken!D122*10);0))"
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
    '20140319
    On Error Resume Next
    Worksheets("Verpacken").Unprotect "bw"
    Worksheets("Verpacken").Activate
    ActiveSheet.Range("D6").Select
    Selection.ClearComments
    Selection.Locked = False
    Selection.FormulaHidden = False
    With Selection.Interior
        .ColorIndex = 6
        .Pattern = xlSolid
    End With
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
        With ComboBox7
            .BoundColumn = 2
            .ColumnCount = 2
            .ListStyle = fmListStyleOption
            .ListFillRange = "SVerpacken!B73:C74"
        End With
        Worksheets("SVerpacken").Range("B71") = 2
       Else
        'ohne Einschweißen
        With ComboBox7
            .BoundColumn = 2
            .ColumnCount = 2
            .ListStyle = fmListStyleOption
            .ListFillRange = "SVerpacken!B72:C74"
        End With
        Worksheets("SVerpacken").Range("B71") = 1
   End If
    'Abstapeln v. Einzelexemplaren
    If Worksheets("Material").Range("C121") = 15 Then
        With ComboBox7
            .BoundColumn = 2
            .ColumnCount = 2
            .ListStyle = fmListStyleOption
            .ListFillRange = "SVerpacken!B72:C74"
        End With
        Worksheets("SVerpacken").Range("B71") = 1
        Worksheets("Verpacken").Range("D6") = 1
        ActiveSheet.Range("D6").Select
        Range("D6").Activate
        With Selection.Interior
            .ColorIndex = 0
        End With
        Selection.Locked = True
        Selection.FormulaHidden = False
        Range("D6").AddComment
        Range("D6").Comment.Visible = False
        Range("D6").Comment.Text Text:="Einfaches Abstapeln ohne Folie od. Karton." & Chr(10) & _
        "Bei Bedarf Zwischenlagen festlegen."
    End If
    ActiveSheet.Range("D6").Select
    Worksheets("Verpacken").Protect "bw"
    Call Kommentar
End Sub
Sub lkw()
    'Berechnung der benötigten Lkws
    On Error Resume Next
    Dim intGewicht, intOptPaletten, intPaletten As Integer
    Dim strActiveCell As String
    intGewicht = Worksheets("SVerpacken").Range("K53")
    intGewichtIst = Worksheets("SVerpacken").Range("L54")
    intOptPaletten = Worksheets("SVerpacken").Range("G52")
    intPaletten = Worksheets("SVerpacken").Range("K52")
    strActiveCell = ActiveCell.Address
    Worksheets("Verpacken").Unprotect "bw"
    If ToggleButton1.Value = True Then
        Worksheets("SVerpacken").Range("K55").Value = 1
        If Worksheets("SVerpacken").Range("K54") > 0 Then
            MsgBox "Achtung! Das Frachtgewicht von " & intGewicht / 1000 & " t wurde um " & intGewichtIst & "t überschritten." & vbCrLf & vbCrLf & _
            "Es können maximal nur " & intOptPaletten & " Paletten je Lkw transportiert werden."
            Worksheets("Verpacken").Activate
            Worksheets("Verpacken").Range("C40:D40,F40:G40").Select
            With Selection.Interior
                .ColorIndex = 3
                .Pattern = xlSolid
            End With
            Selection.Locked = False
            Else
            Worksheets("Verpacken").Range("C40:D40,F40:G40").Select
            With Selection.Interior
                .ColorIndex = 0
                .Pattern = xlSolid
            End With
            Selection.Locked = False
        End If
        Else
        Worksheets("Verpacken").Activate
        Worksheets("Verpacken").Range("C40:D40,F40:G40").Select
        With Selection.Interior
            .ColorIndex = 0
            .Pattern = xlSolid
        End With
        Selection.Locked = False
        Worksheets("SVerpacken").Range("K55").Value = 0
    End If
    ActiveSheet.Range("D6").Select
    'Worksheets("Verpacken").Range(strActiveCell).Activate
End Sub
Sub Palettencheck()
    'Palettenwerte setzen
    'modified 20080724
    On Error Resume Next
    Dim intAuswahlPl As Integer
    Dim intRow, intCol As Integer
    intRow = ActiveCell.Row
    intCol = ActiveCell.Column
    intAuswahlPl = Worksheets("SVerpacken").Range("H7")
    If intAuswahlPl = 8 Then
        Worksheets("Verpacken").Range("D23").FormulaLocal = "=(SVERWEIS(SVerpacken!H7;Material!B79:J93;9)+SVERWEIS(SVerpacken!H7;Material!B79:J93;5))/10"
        Worksheets("Verpacken").Range("F23") = "(Palettenvorgabe)"
        Worksheets("Verpacken").Range("D23").Select
        With Selection.Interior
            .ColorIndex = 0
            .Pattern = xlSolid
        End With
    Else
        Worksheets("Verpacken").Range("D23") = 120
        Worksheets("Verpacken").Range("F23") = ""
        Worksheets("Verpacken").Range("D23").Select
        With Selection.Interior
            .ColorIndex = 6
            .Pattern = xlSolid
        End With
    End If
    Cells(intRow, intCol).Select
    Call Kommentar
End Sub
Sub Kommentar()
    'Anzeigen d. Palettendaten
    '20140401
    On Error Resume Next
    Dim VPalette, VExemplare As String
    'Worksheets("Verpacken").Unprotect "bw"
    VPalette = Worksheets("SVerpacken").Range("I39")
    VExemplare = Worksheets("SVerpacken").Range("I45")
    Worksheets("Verpacken").Label2.Caption = "Gewicht volle Einzelpalette:" & Chr(10) & Chr(10) & VPalette & " kg"
End Sub


