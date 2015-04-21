VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Standertwerte_Sammelheften()
    '20140623
    Application.ScreenUpdating = False
    Worksheets("SSammelheften").Range("I2:I22").Copy Destination:=Worksheets("Sammelheften").Range("K3")
    Application.ScreenUpdating = True
End Sub
Sub Zusammentragen()
'Ermittlung der Stationen und der Anzahl Vorsammeln
Dim a, BgAnzahl, AnzahlBgZus, AnzahlBgVors, AnzahlVors, intStatMa, intBg As Integer
Dim intSeitenMax, intSeitenBgMin, intSeitenBgMax As Integer

intSeitenMax = Worksheets("SSammelheften").Range("J28") 'max Produktseiten
intStatMa = Worksheets("SSammelheften").Range("J29") 'max Stationen
intSeitenBgMin = Worksheets("SSammelheften").Range("J30") 'Mindestseiten pro Bogen
intSeitenBgMax = Worksheets("SSammelheften").Range("J31") 'Maximalseiten pro Bogen
intBg = Worksheets("SSammelheften").Range("J32") 'Bogensumme Inhalt
intSeiten = Worksheets("SSammelheften").Range("J33") 'Produktseiten
'Anzahl Vorsammeln ermitteln
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
    Worksheets("Sammelheften").Range("D3") = intAnzahlBgZus
    Worksheets("Sammelheften").Range("D20") = intAnzahlBgVors
    Worksheets("Sammelheften").Range("D21") = intAnzahlVors
    Else
    intAnzahlBgZus = intBg
    intAnzahlBgVors = 0
    intAnzahlVors = 0
    Worksheets("Sammelheften").Range("D3") = intAnzahlBgZus
    Worksheets("Sammelheften").Range("D20") = intAnzahlBgVors
    Worksheets("Sammelheften").Range("D21") = intAnzahlVors
End If
If Worksheets("SEingabe").Range("G5") > 1 Then
    MsgBox ("Sie haben " & Worksheets("SEingabe").Range("G5") & " Umschläge!")
End If

End Sub

