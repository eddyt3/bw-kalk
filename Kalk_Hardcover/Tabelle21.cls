VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    Call produkt
    Call Worksheets("Verpacken").Kommentar
    Run Worksheets("SVerpacken").draw_all
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    Call Worksheets("Verpacken").Kommentar
End Sub
Private Sub ToggleButton1_Click()
    Dim intGewicht, intPaletten As Integer
    intGewicht = Worksheets("SVerpacken").Range("K53")
    intPaletten = Worksheets("SVerpacken").Range("K52")
    Worksheets("Verpacken").ToggleButton1.Caption = "Soll: " & IIf(ToggleButton1.Value = True, intPaletten & " Paletten", (intGewicht / 1000) & "t Frachtgewicht")
    Call Worksheets("Verpacken").lkw
End Sub
Private Sub CheckBox1_Click()
    Call Worksheets("Verpacken").Kommentar
    Run Worksheets("SVerpacken").draw_all
End Sub
Private Sub CheckBox4_Click()
    If Worksheets("Verpacken").CheckBox4.Value = True Then
        Worksheets("Verpacken").Range("D31").Interior.ColorIndex = 6
        Worksheets("Verpacken").Range("D31").Interior.Pattern = xlSolid
        Worksheets("Verpacken").Range("D31").Locked = False
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
    Call Worksheets("Verpacken").Kommentar
End Sub
Private Sub ComboBox1_DropButtonClick()
    Dim intRow As Integer
    intRow = Worksheets("Material").Range("D78")
    Worksheets("Verpacken").ComboBox1.ListRows = intRow
    Worksheets("Verpacken").ComboBox1.ListFillRange = "Material!C79:C" & (79 + intRow - 1)
    Call Worksheets("Verpacken").Kommentar
    Call Worksheets("Verpacken").Palettencheck
    Run Worksheets("SVerpacken").draw_all
End Sub
Private Sub ComboBox2_Change()
'Verpackungsart
    Call Worksheets("Verpacken").Verpackungseinheit
    Run Worksheets("SVerpacken").draw_all
    Call Worksheets("Verpacken").Kommentar
    If Worksheets("SVerpacken").Range("G25") = 0 Then
        MsgBox "Achtung!" & vbLf & vbLf & "Verpackungsmaterial < Produkt!"
        Worksheets("Verpacken").ComboBox2.BackColor = &HFF
        Else: Worksheets("Verpacken").ComboBox2.BackColor = vbYellow
    End If
    If ComboBox2.ListIndex = 14 Then
        ComboBox6.ListIndex = 4 'Zwischenlagen aller x mm, wenn ohne Verpackung
    'Else: ComboBox6.ListIndex = 1 'Jede 2. Lage
    End If
End Sub
Private Sub ComboBox3_Change()
    Call Worksheets("Verpacken").Kommentar
End Sub
Private Sub ComboBox7_Change()
    If Worksheets("SVerpacken").Range("B71") > 1 Then
        Worksheets("Verpacken").Range("D65").FormulaLocal = "=SVERWEIS(SVerpacken!B71;SVerpacken!A72:I75;9)"
        Run Worksheets("SVerpacken").draw_all
    End If
End Sub
Private Sub ComboBox8_Change()
    Call Worksheets("Verpacken").Kommentar
End Sub
Private Sub ComboBox9_Change()
    Call Worksheets("Verpacken").Kommentar
End Sub
Private Sub ComboBox13_Change()
    Run Worksheets("SVerpacken").draw_all
End Sub
Private Sub ComboBox14_Change()
    Run Worksheets("SVerpacken").draw_all
End Sub
Private Sub ComboBox16_Change()
    Run Worksheets("SVerpacken").draw_all
End Sub
Private Sub ComboBox17_Change()
    Run Worksheets("SVerpacken").draw_all
End Sub
Private Sub CommandButton1_Click()
    If Worksheets("Material").Range("C121") < 14 Then
        Worksheets("Verpacken").Range("D6").FormulaLocal = "=WENN(ODER(Material!C121=14;Material!C121=13;Material!C121=12);1;WENN(SVerpacken!B121<SVerpacken!G27;SVerpacken!B121;SVerpacken!G27))"
    End If
    Call Worksheets("Verpacken").Kommentar
End Sub
Private Sub CommandButton2_Click()
    Worksheets("Material").Visible = xlSheetVisible
    Worksheets("Material").Activate
    Worksheets("Material").Range("B74").Select
End Sub
Sub Verpackungseinheit()
    Worksheets("Verpacken").Range("D6").ClearComments
    Worksheets("Verpacken").Range("D6").Locked = False
    Worksheets("Verpacken").Range("D6").FormulaHidden = False
    Worksheets("Verpacken").Range("D6").Interior.ColorIndex = 6
    Worksheets("Verpacken").Range("D6").Interior.Pattern = xlSolid
    If Worksheets("Material").Range("C121") < 14 Then
        If Worksheets("Verpacken").Range("D6").HasFormula = False Then
            MsgBox "Achtung!" & Chr(10) & "Die Anzahl der Exemplare/VE wurde manuell eingetragen." & Chr(10) & "Für die Automatische Berechnung bitte Reset drücken."
        End If
    End If
    If Worksheets("Material").Range("C121") = 14 Then
        With Worksheets("Verpacken").ComboBox7
            .ListIndex = 2
        End With
    End If
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
        Worksheets("Verpacken").Range("D6").Comment.Text Text:="Einfaches Abstapeln ohne Folie od. Karton." & Chr(10) & "Bei Bedarf Zwischenlagen festlegen."
    End If
    Call Worksheets("Verpacken").Kommentar
End Sub
Sub lkw()
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
            MsgBox "Achtung! Das Frachtgewicht von " & intGewicht / 1000 & " t wurde um " & intGewichtIst & "t überschritten." & vbCrLf & vbCrLf & "Es können maximal nur " & intOptPaletten & " Paletten je Lkw transportiert werden."
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
    Dim intAuswahlPl As Integer
    intAuswahlPl = Worksheets("SVerpacken").Range("H7")
    Worksheets("Verpacken").Range("D23").FormulaLocal = "=(SVERWEIS(SVerpacken!H7;Material!B79:J93;9)+SVERWEIS(SVerpacken!H7;Material!B79:J93;5))/10"
    If intAuswahlPl = 8 Or intAuswahlPl = 9 Then
        Worksheets("Verpacken").Range("F23") = "(Palettenvorgabe)"
        Worksheets("Verpacken").Range("D23").Interior.ColorIndex = 0
        Worksheets("Verpacken").Range("D23").Interior.Pattern = xlSolid
    Else
        Worksheets("Verpacken").Range("F23") = ""
        Worksheets("Verpacken").Range("D23").Interior.ColorIndex = 6
        Worksheets("Verpacken").Range("D23").Interior.Pattern = xlSolid
    End If
    Call Worksheets("Verpacken").Kommentar
End Sub
Sub Kommentar()
    Dim VPalette, VExemplare As String
    VPalette = Worksheets("SVerpacken").Range("I39")
    VExemplare = Worksheets("SVerpacken").Range("G45")
    Worksheets("Verpacken").Label2.Caption = "Volle Einzelpalette:" & Chr(10) & VPalette & " kg" & Chr(10) & VExemplare & " Ex."
End Sub
