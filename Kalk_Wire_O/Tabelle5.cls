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
    'Geschwindigkeitsauswahl bei Einzel-, Doppelnutzen
    Worksheets("Binden").Range("D4").FormulaLocal = "=WENN(SBinden!M20=1;SVERWEIS(SBinden!B4;SBinden!A5:G12;6);SVERWEIS(SBinden!B4;SBinden!A5:G12;7))"
    Call Rillecart
    Call Inline750
End Sub
Sub Rillecart()
    'Rillecart-Maschinenbedingungen
    '20100614 modified
    Application.ScreenUpdating = False
    Dim intV, intNutzen As Integer
    intNutzen = Worksheets("SBinden").Range("M20").Value
    Select Case intNutzen
    Case Is = 1
        intV = Worksheets("SBinden").Range("F9").Value
    Case Is = 2
        intV = Worksheets("SBinden").Range("G9").Value
    End Select
    'Produktstärke > 0.3, Auswahl Rillecart
    If Worksheets("SBinden").Range("D122").Value > 0.3 And Worksheets("SBinden").Range("B4").Value = 5 Then
        Answer = MsgBox("Hinweis: Produktstärke > 3mm, ausgewählte Maschine: Rillecart." & vbCrLf & vbCrLf & "Für die Rillecart muss geteilt gestanzt werden." _
        & vbCrLf & "(Die Geschwindigkeit wurde halbiert; 1 zusätzl. HK wurde berücksichtigt.)", vbOKOnly, "Rillecart.")
        'keine Mehrfachberechnung
        If Worksheets("SBinden").Range("F9").Value = 750 Then
            Worksheets("SBinden").Range("F9").Value = intV / 2
            Worksheets("Binden").Range("D8").Value = 1
        End If
        Else
        Select Case intNutzen
        'Einzelnutzen
        Case Is = 1
            Worksheets("SBinden").Range("F9").Value = 750
            Worksheets("Binden").Range("D8").Value = 0
        'Doppelnutzen
        Case Is = 2
            Worksheets("SBinden").Range("G9").Value = 0
            Worksheets("Binden").Range("D8").Value = 0
        End Select
    End If
    Application.ScreenUpdating = True
End Sub
Sub Inline750()
    'Inline750-Maschinenbedingungen
    '20101021
    Dim intMaschine, intVorsammeln As Integer
    intMaschine = Worksheets("SBinden").Range("B4").Value
    intVorsammeln = Worksheets("Zusammentragen").Range("D14").Value
    If intMaschine = 6 Then
        If intVorsammeln = 0 Then
            ComboBox2.Visible = True
            'Zusammentr., Stanzen, 3Seitenbesch. deaktivieren
            Worksheets("Eingabe").CheckBox2.Value = False
            Worksheets("Eingabe").CheckBox3.Value = False
            Worksheets("Eingabe").CheckBox4.Value = False
        Else
            MsgBox "Achtung!" & vbCrLf & vbCrLf & "Eine Inlinefertigung ist nicht möglich, da Vorgesammelt wird." _
             & vbCrLf & vbCrLf & "Bitte ggf. Zusammentragen auf der Eingabeseite aktivieren."
'            Worksheets("Eingabe").CheckBox2.Value = True
'            Worksheets("Eingabe").CheckBox3.Value = True
'            Worksheets("Eingabe").CheckBox4.Value = True
        End If
    Else
        ComboBox2.Visible = False
    End If
End Sub
