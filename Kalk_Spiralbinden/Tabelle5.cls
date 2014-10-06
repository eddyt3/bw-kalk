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
Private Sub Worksheet_Activate()
    Application.ScreenUpdating = False
    Run Worksheets("Binden").Spirale
    Application.ScreenUpdating = True
End Sub
Private Sub ComboBox1_DropButtonClick()
    'Auswahl Maschine, Einrichtezeit
    Worksheets("Binden").Range("D3").FormulaLocal = "=SVERWEIS(SBinden!B4;SBinden!A5:F12;5)"
    'Geschwindigkeitsauswahl Einzelnutzen
    Worksheets("Binden").Range("D4").FormulaLocal = "=SVERWEIS(SBinden!B4;SBinden!A5:G12;6)"
End Sub
Sub Spirale()
    '20140509
    'Spiralprüfung
    '
    'On Error Resume Next
    Dim intLProdukt, intMaxLProdukt As Integer
    intLBindeseite = Worksheets("SBinden").Range("A64")
    intMaxLProdukt = Worksheets("Binden").Range("K6")
    If intLBindeseite > intMaxLProdukt Then
        MsgBox ("Achtung!" & vbCrLf & vbCrLf & "Die max. Bindelänge (" & intMaxLProdukt & "mm) wurde überschritten." _
        & vbCrLf & vbCrLf & "Bindelänge Ist: " & intLBindeseite & "mm.")
        Worksheets("Eingabe").Range("J66") = "Fehler"
    Else: Worksheets("Eingabe").Range("J66") = ""
    End If
End Sub


