VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ComboBox1_Change()
    Run Worksheets("Buchdecke").Cello_alternativ
End Sub

Private Sub CommandButton1_Click()
    'Materialseite einblenden
    On Error Resume Next
    Worksheets("Material").Visible = xlSheetVisible
    Worksheets("Material").Activate
    Worksheets("Material").Range("H217").Select
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Run Worksheets("SBuchdecke").Beschriftung
End Sub
Sub Cello_alternativ()
    ' Auswahl Cello Alternativkosten
    If Worksheets("SBuchdecke").Range("B3").Value = 5 Then
        Worksheets("Buchdecke").Rows("13:13").Hidden = False
        Else
        Worksheets("Buchdecke").Rows("13:13").Hidden = True
    End If
End Sub

