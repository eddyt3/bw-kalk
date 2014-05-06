VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Private Sub CheckBox3_Click()
'Einblenden Blatt Zusammentragen
    On Error Resume Next
    Application.ScreenUpdating = False
    If CheckBox3.Value = True Then
        Worksheets("Zusammentragen").Visible = True
    Else
        Worksheets("Zusammentragen").Visible = False
    End If
    Application.ScreenUpdating = True
End Sub
Private Sub ComboBox3_DropButtonClick()
' Auswahl Scharnierstoff, Fälzel, Krepp
' 20140506
    Call Stoff
End Sub
Private Sub ComboBox5_Change()
' Auswahl Leimart
' 20140422
    Call Leimart
End Sub
Private Sub CommandButton1_Click()
    'Reset V auf Standard
    '20140325
    On Error Resume Next
    Worksheets("Binden").Range("D8").FormulaLocal = "=Basis_Kolbus!H17"
    Worksheets("Binden").Range("F8").FormulaLocal = "=Basis_Kolbus!I17"
    Worksheets("Binden").Range("G8").FormulaLocal = "=Basis_Kolbus!J17"
End Sub
Sub Leimart()
' Auswahl Leimart
' 20140422
    ' Auswahl Leimart PUR
    On Error Resume Next
    If ComboBox5.ListIndex = 0 Then
        Worksheets("Binden").Range("D3") = ComboBox5.ListIndex + 1
        Worksheets("Binden").ComboBox1.ListFillRange = "Material!D9:D13" 'Auswahl PUR Kleber
        Worksheets("Binden").ComboBox1.ListIndex = 0 'Ersten Wert d. Auswahl autom. anzeigen
        Worksheets("Binden").ComboBox1.LinkedCell = "Material!B8" 'Ausgabe Leimname
        Worksheets("Binden").ComboBox4.ListFillRange = "Material!D38:D43" 'Auswahl PUR Reiniger
        Worksheets("Binden").ComboBox4.ListIndex = 0
        Worksheets("Binden").ComboBox4.LinkedCell = "Material!B37" 'Ausgabe PUR Reiniger
    Else
        Worksheets("Binden").ComboBox1.ListFillRange = "Material!D14:D18" 'Auswahl HM Kleber
        Worksheets("Binden").ComboBox1.ListIndex = 0
        Worksheets("Binden").ComboBox1.LinkedCell = "Material!B8" 'Ausgabe Leimname
        Worksheets("Binden").ComboBox4.ListFillRange = "Material!D43" 'kein Reiniger
        Worksheets("Binden").ComboBox4.ListIndex = 0
        Worksheets("Binden").ComboBox4.LinkedCell = "Material!B37" 'kein Reiniger
    End If
    ' Auswahl Leimart Hotmelt
    If ComboBox5.ListIndex = 1 Then
        Worksheets("Binden").Range("D3") = ComboBox5.ListIndex + 1
        Worksheets("Binden").ComboBox1.ListFillRange = "Material!D14:D18"
        Worksheets("Binden").ComboBox1.ListIndex = 0
        Worksheets("Binden").ComboBox1.LinkedCell = "Material!B8"
        Worksheets("Binden").ComboBox4.ListFillRange = "Material!D43" 'kein Reiniger
        Worksheets("Binden").ComboBox4.ListIndex = 0
        Worksheets("Binden").ComboBox4.LinkedCell = "Material!B37" 'kein Reiniger
    Else
        Worksheets("Binden").ComboBox1.ListFillRange = "Material!D9:D13"
        Worksheets("Binden").ComboBox1.ListIndex = 0
        Worksheets("Binden").ComboBox1.LinkedCell = "Material!B8"
        Worksheets("Binden").ComboBox4.ListFillRange = "Material!D38:D43" 'Auswahl PUR Reiniger
        Worksheets("Binden").ComboBox4.ListIndex = 0
        Worksheets("Binden").ComboBox4.LinkedCell = "Material!B37" 'Ausgabe PUR Reiniger
    End If
    Worksheets("Binden").ComboBox2.ListFillRange = "Material!D26:D36"
    Worksheets("Binden").ComboBox2.ListIndex = 0
    Worksheets("Binden").ComboBox2.LinkedCell = "Material!B25"
End Sub
Sub Produktart()
' Broschüre, Buchblock, Ableimen
' 20140423
    Dim intArt As Integer
    intArt = Worksheets("SBinden").Range("B34")
    Select Case intArt
    Case 1
        'Broschüre
        'Hotmelt (Standardleim), mit Seitenleim, ohne Fälzel, ohne Vorsatz, mit Fadenheften
        Worksheets("Binden").ComboBox5.ListIndex = 1 'Hotmelt auswählen
        Worksheets("Binden").Range("D3") = ComboBox5.ListIndex + 1
        Worksheets("Binden").ComboBox1.ListFillRange = "Material!D14:D18" 'Leimsorte
        Worksheets("Binden").ComboBox1.LinkedCell = "Material!B8"
        Worksheets("Binden").ComboBox1.ListIndex = 0 'Topmelt
        Worksheets("Binden").ComboBox2.ListFillRange = "Material!D26:D36" 'Seitenleim
        Worksheets("Binden").ComboBox2.ListIndex = 0 'mit Seitenleim
        Worksheets("Binden").ComboBox2.LinkedCell = "Material!B25"
        Worksheets("Binden").ComboBox3.ListFillRange = "Material!D52:D55" 'Rückenstoff
        Worksheets("Binden").ComboBox3.LinkedCell = "Material!B51"
        Worksheets("Binden").ComboBox3.ListIndex = 0 'Rücken ohne Stoff
        Worksheets("Binden").ComboBox4.ListFillRange = "Material!D43" 'kein Reiniger
        Worksheets("Binden").ComboBox4.LinkedCell = "Material!B37"
        Worksheets("Binden").ComboBox4.ListIndex = 0 'kein Reiniger
    Case 2
        'Buchblock
        'Hotmelt (Standardleim), mit Seitenleim, mit Fälzel, mit Vorsatz und Fadenheften
        Worksheets("Binden").ComboBox5.ListIndex = 1 'Hotmelt auswählen
        Worksheets("Binden").Range("D3") = ComboBox5.ListIndex + 1
        Worksheets("Binden").ComboBox1.ListFillRange = "Material!D14:D18" 'Leimsorte
        Worksheets("Binden").ComboBox1.LinkedCell = "Material!B8"
        Worksheets("Binden").ComboBox1.ListIndex = 0 'Topmelt
        Worksheets("Binden").ComboBox2.ListFillRange = "Material!D26:D36" 'Seitenleim
        Worksheets("Binden").ComboBox2.ListIndex = 0 'mit Seitenleim
        Worksheets("Binden").ComboBox2.LinkedCell = "Material!B25"
        Worksheets("Binden").ComboBox3.ListFillRange = "Material!D52:D55" 'Rückenstoff
        Worksheets("Binden").ComboBox3.LinkedCell = "Material!B51"
        Worksheets("Binden").ComboBox3.ListIndex = 3 'Rücken mit Krepp
        Worksheets("Binden").ComboBox4.ListFillRange = "Material!D43" 'kein Reiniger
        Worksheets("Binden").ComboBox4.LinkedCell = "Material!B37"
        Worksheets("Binden").ComboBox4.ListIndex = 0 'kein Reiniger
    Case 3
        'Ableimen
        'Hotmelt (Katalogleim), 1/4 Leimeinsatz, ohne Seitenleim, ohne Fälzel
        Worksheets("Binden").ComboBox5.ListIndex = 1 'Hotmelt auswählen
        Worksheets("Binden").Range("D3") = ComboBox5.ListIndex + 1
        Worksheets("Binden").ComboBox1.ListFillRange = "Material!D14:D18" 'Leimsorte
        Worksheets("Binden").ComboBox1.LinkedCell = "Material!B8"
        Worksheets("Binden").ComboBox1.ListIndex = 1 'Katalogleim
        Worksheets("Binden").ComboBox2.ListFillRange = "Material!D26:D36" 'Seitenleim
        Worksheets("Binden").ComboBox2.ListIndex = 10 'ohne Seitenleim
        Worksheets("Binden").ComboBox2.LinkedCell = "Material!B25"
        Worksheets("Binden").ComboBox3.ListFillRange = "Material!D52:D55" 'Rückenstoff
        Worksheets("Binden").ComboBox3.LinkedCell = "Material!B51"
        Worksheets("Binden").ComboBox3.ListIndex = 0 'Rücken ohne Stoff
        Worksheets("Binden").ComboBox4.ListFillRange = "Material!D43" 'kein Reiniger
        Worksheets("Binden").ComboBox4.LinkedCell = "Material!B37"
        Worksheets("Binden").ComboBox4.ListIndex = 0 'kein Reiniger
    End Select
    Call Stoff
End Sub
Sub Stoff()
' Auswahl Scharnierstoff, Fälzel, Krepp
' 20140506
    On Error Resume Next
    If Worksheets("Material").Range("C51") > 1 Then
        Worksheets("Binden").Unprotect "bw"
        Worksheets("Binden").Rows("20:28").Hidden = False
        Worksheets("Binden").Rows("27").Hidden = True
    Else
        Worksheets("Binden").Unprotect "bw"
        Worksheets("Binden").Rows("20:28").Hidden = True
    End If
    Worksheets("Binden").Protect "bw"
End Sub
