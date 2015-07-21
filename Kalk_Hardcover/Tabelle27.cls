VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Private Sub ComboBox2_DropButtonClick()
     Run Worksheets("Zusammentragen").Zusammentragen
     Run Worksheets("Zusammentragen").GeschwindigkeitZus
End Sub
Private Sub Worksheet_Activate()
    Run Worksheets("Zusammentragen").Zusammentragen
End Sub
Sub Zusammentragen()
    '
    'Ermittlung der Stationen und der Anzahl Vorsammeln
    '20150701 modifiziert
    Dim a, intM, BgAnzahl, AnzahlBgZus, AnzahlBgVors, AnzahlVors, intMa, intStatMa, intBg As Integer
    intStatMa = Worksheets("SZusammentragen").Range("E11") 'max Stationen
    intBg = Worksheets("SZusammentragen").Range("B8") 'Bogensumme
    intMa = Worksheets("SZusammentragen").Range("C11") 'Maschine
    If intBg > intStatMa Then
        a = 1
        Do
            BgAnzahl = intBg / a
            a = a + 1
        Loop Until BgAnzahl <= intStatMa
        If BgAnzahl - Int(BgAnzahl) > 0 Then
            intAnzahlBgVors = Int(BgAnzahl) + 1
            intAnzahlBgZus = intBg - intAnzahlBgVors * (a - 2)
            Else
            intAnzahlBgZus = BgAnzahl
            intAnzahlBgVors = BgAnzahl
            End If
        intAnzahlVors = a - 2
        Worksheets("Zusammentragen").Range("D3") = intAnzahlBgZus
        Worksheets("Zusammentragen").Range("D14") = intAnzahlBgVors
        Worksheets("Zusammentragen").Range("D15") = intAnzahlVors
    Else
        intAnzahlBgZus = intBg
        intAnzahlBgVors = 0
        intAnzahlVors = 0
        Worksheets("Zusammentragen").Range("D3") = intAnzahlBgZus
        Worksheets("Zusammentragen").Range("D14") = intAnzahlBgVors
        Worksheets("Zusammentragen").Range("D15") = intAnzahlVors
    End If
    'Formatwechselcheckbox nur bei Maschine TB216
    If intMa = 1 Then
        Dim intBreite As Integer
        Dim intWechsel As Boolean
        intBreite = Worksheets("SZusammentragen").Range("B37")
        intWechsel = Worksheets("SZusammentragen").Range("B38")
        'ab 35 cm Breite immer Formatwechsel berechnen
        If intWechsel Then
            Worksheets("Zusammentragen").CheckBox1.Visible = True
            Worksheets("Zusammentragen").CheckBox1.Value = True
        Else
            Worksheets("Zusammentragen").CheckBox1.Visible = True
            Worksheets("Zusammentragen").CheckBox1.Value = False
        End If
    Else
        Worksheets("Zusammentragen").CheckBox1.Visible = False
        Worksheets("Zusammentragen").CheckBox1.Value = False
    End If
    'Prüfung händisch Zusammenlegen (maschinenabhängig)
    If Worksheets("SZusammentragen").Range("N16") = 1 And Worksheets("Zusammentragen").CheckBox2.Value = False Then
        Worksheets("Zusammentragen").CheckBox2.Value = True
    End If
        If intAnzahlVors = 0 Then
        Worksheets("Zusammentragen").CheckBox2.Value = False
    End If
    Application.ScreenUpdating = True
End Sub

Sub GeschwindigkeitZus()
    '20140515 Festlegen der Geschwindigkeiten
    'Kolbus V auflagenabhängig
    If Worksheets("SZusammentragen").Range("C11") = 4 Then
        Worksheets("Zusammentragen").Range("D4").FormulaLocal = "=Basis_Kolbus!H17"
        Worksheets("Zusammentragen").Range("F4").FormulaLocal = "=Basis_Kolbus!I17"
        Worksheets("Zusammentragen").Range("G4").FormulaLocal = "=Basis_Kolbus!J17"
    Else 'Standard Werte
        Worksheets("Zusammentragen").Range("D4").FormulaLocal = "=SZusammentragen!M11"
        Worksheets("Zusammentragen").Range("F4").FormulaLocal = "=D4"
        Worksheets("Zusammentragen").Range("G4").FormulaLocal = "=D4"
    End If
End Sub
