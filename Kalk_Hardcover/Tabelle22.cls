VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Thickness()
    'Eingabe alternative Pappstärke Buchdecke
    Application.ScreenUpdating = False
    ' Auswahl alternative Materialstärke
        If Worksheets("Eingabe").ComboBox3.ListIndex = 3 Then
            Worksheets("Eingabe").TextBox3.Visible = True
        Else
            Worksheets("Eingabe").TextBox3.Visible = False
        End If
    Application.ScreenUpdating = True
End Sub
Sub Beschriftung()
    'Beschriftung der Skizze
    '20151127
    Dim intRücken, intPappeB, intPappeL, intBezugB, intBezugL, intDeckeB, intDeckeL, intFalz As String
    intRücken = Worksheets("SBuchdecke").Range("K3")
    intPappeB = Worksheets("SBuchdecke").Range("K4")
    intPappeL = Worksheets("SBuchdecke").Range("M4")
    intBezugB = Worksheets("SBuchdecke").Range("K5")
    intBezugL = Worksheets("SBuchdecke").Range("M5")
    intDeckeB = Worksheets("SBuchdecke").Range("K6")
    intDeckeL = Worksheets("SBuchdecke").Range("M6")
    intFalz = Worksheets("SBuchdecke").Range("K7")
    Worksheets("Buchdecke").Label1.Caption = intRücken & Chr(10) & "mm"
    Worksheets("Buchdecke").Label2.Caption = intPappeB & "mm"
    Worksheets("Buchdecke").Label3.Caption = intPappeL & "mm"
    Worksheets("Buchdecke").Label4.Caption = intBezugB & "mm"
    Worksheets("Buchdecke").Label5.Caption = intBezugL & Chr(10) & "mm"
    Worksheets("Buchdecke").Label6.Caption = intDeckeB & "mm"
    Worksheets("Buchdecke").Label7.Caption = intDeckeL & Chr(10) & "mm"
    Worksheets("Buchdecke").Label8.Caption = intFalz & Chr(10) & "mm"
    Worksheets("Buchdecke").Label9.Caption = intFalz & Chr(10) & "mm"
End Sub
