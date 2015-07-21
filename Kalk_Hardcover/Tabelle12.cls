VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Standardwerte
'Cello Kunde geliefert

Private Sub Worksheet_Activate()
    Run Worksheets("SBuchdecke").Beschriftung
    Run Worksheets("SBuchdecke").HardcoverVisual
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    Run Worksheets("SBuchdecke").Beschriftung
End Sub
Private Sub ComboBox1_Change()
    Run ThisWorkbook.Sheets("Buchdecke").Cello_alternativ
End Sub
Private Sub CommandButton1_Click()
    'Materialseite einblenden
    Worksheets("Material").Visible = xlSheetVisible
    Worksheets("Material").Activate
    Worksheets("Material").Range("H217").Select
End Sub
Sub Cello_alternativ()
    ' Auswahl Cello Alternativkosten
    If Worksheets("Buchdecke").ComboBox1.ListIndex = 4 Then
        Worksheets("Buchdecke").Rows("13:13").Hidden = False
        Worksheets("Buchdecke").TextBox21.Visible = True
        Worksheets("Eingabe").TextBox21.Visible = True
    Else
        Worksheets("Buchdecke").Rows("13:13").Hidden = True
        Worksheets("Buchdecke").Range("D13:G13").Value = 0
        Worksheets("Buchdecke").TextBox21.Visible = False
        Worksheets("Eingabe").TextBox21.Visible = False
    End If
End Sub
