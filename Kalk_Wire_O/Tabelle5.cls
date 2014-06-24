VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Private Sub ComboBox1_DropButtonClick()
    'Auswahl Maschine, Einrichtezeit
    Worksheets("Binden").Range("D3").FormulaLocal = "=SVERWEIS(SBinden!B4;SBinden!A5:F12;5)"
    'Geschwindigkeitsauswahl
    Worksheets("Binden").Range("D4").FormulaLocal = "=SVERWEIS(SBinden!B4;SBinden!A5:G12;6)"
    Call Rillecart
    Call Inline
End Sub
Sub Rillecart()
    'Rillecart-Maschinenbedingungen P1765 PB796 grün (1MFhr / 2HK)
    '20140508 modified
    Application.ScreenUpdating = False
    Dim intV As Integer
        intV = Worksheets("SBinden").Range("F8").Value
    'Produktstärke > 0.32, Auswahl Rillecart
    If Worksheets("SBinden").Range("D122").Value > 0.32 And Worksheets("SBinden").Range("B4").Value = 4 Then
        Answer = MsgBox("Hinweis: Produktstärke > 3,2 mm, ausgewählte Maschine: Rillecart (P1765)." & vbCrLf & vbCrLf & "Für die Rillecart muss geteilt gestanzt werden." _
        & vbCrLf & "(Die Geschwindigkeit wurde halbiert; 1 zusätzl. HK wurde berücksichtigt.)", vbOKOnly, "Rillecart.")
        'keine Mehrfachberechnung
        If intV = 600 Then
            Worksheets("SBinden").Range("F8").Value = intV / 2
            Worksheets("Binden").Range("D8").Value = 1
        End If
    Else
        Worksheets("SBinden").Range("F8").Value = 600
        Worksheets("Binden").Range("D8").Value = 0
    End If
    Application.ScreenUpdating = True
End Sub
Sub Inline()
    'Inline-Maschinenbedingungen P1766 PB796 grau + ZTM TB216 (1MFhr / 2HK)
    '20140507
    Dim intMaschine, intVorsammeln, intBindeB, intBindeL As Integer
    Dim sngBlock As Single
    intMaschine = Worksheets("SBinden").Range("X13").Value
    intVorsammeln = Worksheets("Zusammentragen").Range("D14").Value
    intBindeB = Worksheets("SBinden").Range("B122").Value
    intBindeL = Worksheets("SBinden").Range("C122").Value
    sngBlock = Worksheets("SBinden").Range("D122").Value
    Application.ScreenUpdating = False
    Worksheets("Eingabe").Unprotect "bw"
    If intMaschine = 1 Then 'Inlinefertigung
        If intVorsammeln = 0 Then
            'Zusammentr., Stanzen, 3Seitenbesch. deaktivieren
            Worksheets("Eingabe").CheckBox2.Value = False
            Worksheets("Eingabe").CheckBox2.Enabled = False
            Worksheets("Eingabe").CheckBox3.Value = False
            Worksheets("Eingabe").CheckBox3.Enabled = False
            Worksheets("Eingabe").CheckBox4.Value = False
            Worksheets("Eingabe").CheckBox4.Enabled = False
        Else
            MsgBox "Achtung!" & vbCrLf & vbCrLf & "Eine Inlinefertigung ist nicht möglich, da Vorgesammelt wird." _
             & vbCrLf & vbCrLf & "Bitte ggf. Zusammentragen auf der Eingabeseite aktivieren."
            Worksheets("Eingabe").CheckBox2.Enabled = True
            Worksheets("Eingabe").CheckBox3.Enabled = True
            Worksheets("Eingabe").CheckBox4.Enabled = True
        End If
    Else
        Worksheets("Eingabe").CheckBox2.Enabled = True
        Worksheets("Eingabe").CheckBox3.Enabled = True
        Worksheets("Eingabe").CheckBox4.Enabled = True
    End If
    'V + HK festlegen
    'V größenabhängig
    If intBindeB <= 42 And intBindeL <= 42 Then
        If sngBlock <= 0.32 Then
            Worksheets("SBinden").Range("F9").Value = 750
        Else: Worksheets("SBinden").Range("F9").Value = 450
        End If
    Else
        If sngBlock <= 0.32 Then
            Worksheets("SBinden").Range("F9").Value = 600
        Else: Worksheets("SBinden").Range("F9").Value = 450
        End If
    End If
    'V geteilt gestanzt
    If sngBlock > 0.32 And Worksheets("SBinden").Range("B4").Value = 5 Then
        Answer = MsgBox("Hinweis: Produktstärke > 3,2 mm, ausgewählte Maschine: Inline (P1766)." & vbCrLf & vbCrLf & "Für die Inline muss geteilt gestanzt werden." _
        & vbCrLf & "(Die Geschwindigkeit wurde angepasst; 1 zusätzl. HK wurde berücksichtigt.)", vbOKOnly, "Inline.")
        Worksheets("Binden").Range("D8").Value = 1
    Else: Worksheets("Binden").Range("D8").Value = 0
    End If
    Worksheets("Eingabe").Protect "bw"
    Application.ScreenUpdating = True
End Sub
