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
'Formatwechsel bei TB216
    Run Worksheets("Zusammentragen").Zusammentragen
    Run Worksheets("Zusammentragen").Formatfehler
End Sub
Private Sub ComboBox2_DropButtonClick()
     Run Worksheets("Zusammentragen").Zusammentragen
     Run Worksheets("Zusammentragen").Formatfehler
End Sub
Private Sub Worksheet_Activate()
    Run Worksheets("Zusammentragen").Zusammentragen
    Run Worksheets("Zusammentragen").Formatfehler
End Sub
Sub Zusammentragen()
    'Ermittlung der Stationen und der Anzahl Vorsammeln
    '20140415 modifiziert
    Application.ScreenUpdating = False
    Dim a, intM, BgAnzahl, AnzahlBgZus, AnzahlBgVors, AnzahlVors, intMa, intStatMa, intBg As Integer
    intStatMa = Worksheets("SZusammentragen").Range("E11") 'max Stationen
    intBg = Worksheets("SZusammentragen").Range("B8") 'Bogensumme
    intMa = Worksheets("Zusammentragen").Range("A4") 'Maschine
    'Bei Formatwechsel TB216 nur die H�lfte der Maschinen ber�cksichtigen
        If intMa = 1 Then
            Dim intWechsel As Boolean
            intWechsel = Worksheets("SZusammentragen").Range("B38")
            'ab (35 cm Breite u. H�he) oder (Streifenkalender = Breite < 35cm, H�he > 35cm, Bindung schmale Seite) immer Formatwechsel berechnen
            If intWechsel Then
                Worksheets("Zusammentragen").CheckBox1.Visible = True
                Worksheets("Zusammentragen").CheckBox1.Value = True
                intStatMa = intStatMa / 2
            Else
                Worksheets("Zusammentragen").CheckBox1.Visible = False
                Worksheets("Zusammentragen").CheckBox1.Value = False
            End If
        Else
            Worksheets("Zusammentragen").CheckBox1.Visible = False
            Worksheets("Zusammentragen").CheckBox1.Value = False
        End If
    'Zusammentragen
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
    Application.ScreenUpdating = True
End Sub
Sub Formatfehler()
    'Pr�fung der Min./Max. Formate der Maschinen und Ausgabe einer Fehlermeldung
    Dim Z, F, Fb, Fh As Integer
    Z = Worksheets("SZusammentragen").Range("B2") 'Zusammentragen ja/nein
    F = Worksheets("SZusammentragen").Range("Q11") 'Fehlerwert
    Fb = Worksheets("SZusammentragen").Range("O11") 'Fehlerwert Breite
    Fh = Worksheets("SZusammentragen").Range("P11") 'Fehlerwert H�he
    Worksheets("Eingabe").Range("J60") = ""
    If Z > 0 Then
        If F = 1 Or F = 2 Then
            Worksheets("Eingabe").Range("J60") = "Fehler"
            If Fb = 1 Then
               MsgBox "Fehler Zusammentragen!" & vbCrLf & vbCrLf & "Die Produktbreite �ber-/unterschreitet das Maschinenma�."
            End If
            If Fh = 2 Then
               MsgBox "Fehler Zusammentragen!" & vbCrLf & vbCrLf & "Die Produkth�he �ber-/unterschreitet das Maschinenma�."
            End If
        End If
        If F = 3 Then
            MsgBox "Fehler Zusammentragen!" & vbCrLf & vbCrLf & "Die Produktbreite und -h�he �ber-/unterschreiten die Maschinenma�e."
        End If
    End If
End Sub
