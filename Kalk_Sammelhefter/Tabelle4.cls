VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'
'Copyright by E.Dargel ed@dissenter.de
'
Private Sub Worksheet_Change(ByVal Target As Range)
    
    Worksheets("Sammelheften").EnableOutlining = True
End Sub
Private Sub Worksheet_Activate()
    Run Worksheets("SSammelheften").Zusammentragen
End Sub
Private Sub CommandButton1_Click()
    'Reset V auf Standard
    '20150112
    If Worksheets("Sammelheften").Range("D4").HasFormula = False Then
        Worksheets("Sammelheften").Range("D4").FormulaLocal = "=SSammelheften!E51"
    End If
End Sub
Private Sub ComboBox2_Change()
    'Berechnung der Klammern und Ösen
    '20150415
    Dim intTyp, intAnz As Integer
    intTyp = Sheets("SSammelheften").[B8]
    intAnz = Application.WorksheetFunction.VLookup(Sheets("SSammelheften").[B8], Sheets("SSammelheften").[B9:C11], 2, False)
    Select Case intTyp
    Case 1
        'Klammern
        Worksheets("Eingabe").Range("C15") = 2
        Worksheets("Eingabe").Range("B15") = "Klammern/Ex."
    Case 2
        '2 Ösen
        Worksheets("Eingabe").Range("C15") = intAnz
        Worksheets("Eingabe").Range("B15") = "Ösen/Ex."
    Case 3
        '4 Ösen
        Worksheets("Eingabe").Range("C15") = intAnz
        Worksheets("Eingabe").Range("B15") = "Ösen/Ex."
    End Select
End Sub

