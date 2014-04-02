VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Option Explicit
Sub Zusammentragen()
Attribute Zusammentragen.VB_ProcData.VB_Invoke_Func = " \n14"
'Ermittlung der Stationen und der Anzahl Vorsammeln
'Manuell od. autom. Zusammentragen
If Worksheets("SZusammentragen").Range("C4") = 1 Then
    'automatisches Zusammentragen
    Dim a, BgAnzahl, AnzahlBgZus, AnzahlBgVors, AnzahlVors, intMa, intStatMa(3), intBg As Integer
    Dim intAnzahlBgVors, intAnzahlBgZus, intAnzahlVors As Integer
    
    intStatMa(1) = Worksheets("Steuerung").Range("K135") 'max Stationen
    intStatMa(2) = Worksheets("Steuerung").Range("K141")
    intStatMa(3) = Worksheets("Steuerung").Range("K147")
    intBg = Worksheets("Steuerung").Range("N139")
    intMa = Worksheets("Steuerung").Range("S138")
    Worksheets("Zusammentragen").Range("D7") = 0
    
        If intBg > intStatMa(intMa) Then
            a = 1
            Do
            BgAnzahl = intBg / a
            a = a + 1
            Loop Until BgAnzahl <= intStatMa(intMa)
            If BgAnzahl - Int(BgAnzahl) > 0 Then
                intAnzahlBgVors = Int(BgAnzahl) + 1
                intAnzahlBgZus = intBg - intAnzahlBgVors * (a - 2)
                Else
                intAnzahlBgZus = BgAnzahl
                intAnzahlBgVors = BgAnzahl
                End If
            intAnzahlVors = a - 2
            Worksheets("Zusammentragen").Range("D3") = intAnzahlBgZus
            Worksheets("Zusammentragen").Range("D15") = intAnzahlBgVors
            Worksheets("Zusammentragen").Range("D16") = intAnzahlVors
            Else
            intAnzahlBgZus = intBg
            intAnzahlBgVors = 0
            intAnzahlVors = 0
            Worksheets("Zusammentragen").Range("D3") = intAnzahlBgZus
            Worksheets("Zusammentragen").Range("D15") = intAnzahlBgVors
            Worksheets("Zusammentragen").Range("D16") = intAnzahlVors
        End If
    'manuelles Zusammentragen
    Else
    Worksheets("Zusammentragen").Range("D3") = 1
    Worksheets("Zusammentragen").Range("D7") = 0 ' zusätzl. HK?
    Worksheets("Zusammentragen").Range("D15") = 0
    Worksheets("Zusammentragen").Range("D16") = 0
    End If
    
End Sub
