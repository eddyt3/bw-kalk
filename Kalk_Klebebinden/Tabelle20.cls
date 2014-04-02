VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Sub Formatinput()
    '
    'Eingabe Produktformat
    '
    Dim sngPL, sngPB As Single 'Produkt Länge, Breite
    Do
        sngPL = Application.InputBox("Bitte Produktbreite in cm eingeben:", "Eingabe", , Type:=1)
        If sngPL > 37000 Then
            MsgBox "Bitte nur Komma anstatt einem Punkt verwenden"
        End If
    Loop Until sngPL < 37000
    Worksheets("SEingabe").Range("B16").FormulaLocal = "=RUNDEN(" & sngPL & ";1)"
    sngPL = Worksheets("SEingabe").Range("B16")
    Do
        sngPB = Application.InputBox("Bitte Produkthöhe in cm eingeben:", "Eingabe", , Type:=1)
        If sngPB > 37000 Then
            MsgBox "Bitte nur Komma anstatt einem Punkt verwenden"
        End If
    Loop Until sngPB < 37000
    Worksheets("SEingabe").Range("B17").FormulaLocal = "=RUNDEN(" & sngPB & ";1)"
    sngPB = Worksheets("SEingabe").Range("B17")
    If Worksheets("SEingabe").Range("C18") = 0 Then
        Answer = MsgBox("Ist Ihre Eingabe richtig?" & vbLf & vbLf & sngPL & " cm x " & sngPB & " cm", vbYesNo + 256 + vbQuestion, "Nachfrage")
        Worksheets("Eingabe").CommandButton3.Caption = sngPL & " x " & sngPB & "cm"
        Worksheets("Produktionsdaten").Unprotect "bw"
        Worksheets("Produktionsdaten").Range("C7") = sngPL & " x " & sngPB & "cm"
        Worksheets("Produktionsdaten").Protect "bw"
        FFormat = ""
        Else
        MsgBox ("Achtung! Bitte nur Zahlen eingeben!")
        Worksheets("Eingabe").CommandButton3.Caption = "Fehler!"
        FFormat = "Produktformat: Falscheingabe!"
    End If
End Sub

