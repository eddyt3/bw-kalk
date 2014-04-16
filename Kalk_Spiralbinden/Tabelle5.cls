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
    'Geschwindigkeitsauswahl Einzelnutzen
    Worksheets("Binden").Range("D4").FormulaLocal = "=SVERWEIS(SBinden!B4;SBinden!A5:G12;6)"
End Sub
