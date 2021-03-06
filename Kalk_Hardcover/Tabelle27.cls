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
End Sub
Private Sub Worksheet_Activate()
    Run Worksheets("Zusammentragen").Zusammentragen
End Sub
Sub Zusammentragen()
    'Nur f�r Hardcover g�ltig!
    'Ermittlung der Stationen und der Anzahl Vorsammeln
    '20190604 modifiziert
    Dim a, intM, BgAnzahl, AnzahlBgZus, AnzahlBgVors, AnzahlVors, intMa, intStatMa, intBg As Integer
    intStatMa = Worksheets("SZusammentragen").Range("E11") 'max Stationen
    intFadenh = Worksheets("SEingabe").Range("B34") 'Fadenheften ja/nein
    intMa = Worksheets("SZusammentragen").Range("C11") 'Maschine
    If intFadenh = 1 Then
        intBg = Worksheets("SZusammentragen").Range("D8") 'Bogensumme > 4 Seiter
    Else
        intBg = Worksheets("SZusammentragen").Range("B8") 'alle Inhaltsbg.
    End If
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
        intBreite = Worksheets("SZusammentragen").Range("B38")
        intWechsel = Worksheets("SZusammentragen").Range("B39")
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
    'Pr�fung h�ndisch Zusammenlegen (maschinenabh�ngig)
    If Worksheets("SZusammentragen").Range("O17") = 1 And Worksheets("Zusammentragen").CheckBox2.Value = False Then
        Worksheets("Zusammentragen").CheckBox2.Value = True
    End If
        If intAnzahlVors = 0 Then
        Worksheets("Zusammentragen").CheckBox2.Value = False
        Worksheets("Zusammentragen").Range("D21") = "" 'HK
    End If
    'HK Vorsammeln Uniplex
    If intMa = 5 And intAnzahlVors > 0 Then
        If Worksheets("Zusammentragen").Range("D21") = "" Then 'nur eintragen wenn kein anderer Wert drin stand
            Worksheets("Zusammentragen").Range("D21") = 1
        'Else: Worksheets("Zusammentragen").Range("D21") = ""
        End If
    End If
    Run Worksheets("Zusammentragen").GeschwindigkeitZus
    Application.ScreenUpdating = True
End Sub
Sub GeschwindigkeitZus()
    '20140515 Festlegen der Geschwindigkeiten
    'Kolbus V auflagenabh�ngig
    If Worksheets("SZusammentragen").Range("C11") = 4 Then
        Worksheets("Zusammentragen").Range("D4").FormulaLocal = "=Basis_Kolbus!H20"
        Worksheets("Zusammentragen").Range("F4").FormulaLocal = "=Basis_Kolbus!I20"
        Worksheets("Zusammentragen").Range("G4").FormulaLocal = "=Basis_Kolbus!J20"
    Else 'Standard Werte
        Worksheets("Zusammentragen").Range("D4").FormulaLocal = "=SZusammentragen!N11"
        Worksheets("Zusammentragen").Range("F4").FormulaLocal = "=D4"
        Worksheets("Zusammentragen").Range("G4").FormulaLocal = "=D4"
    End If
End Sub

