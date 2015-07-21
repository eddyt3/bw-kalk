VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'

Private Sub CheckBox2_Click()
    'Auswahl Eckenrunden
    '20090824
    If Worksheets("Bohren").CheckBox2.Value = True Then
        Worksheets("SBohren").Range("F5") = Application.InputBox("Bitte Geschwindigkeit für" & vbCrLf & "Bohrungen/Stunde eingeben:", "Eingabe", , Type:=1)
    Else: Worksheets("SBohren").Range("F5").FormulaLocal = "=SVERWEIS(B12;C13:F15;4)"
    End If
End Sub
