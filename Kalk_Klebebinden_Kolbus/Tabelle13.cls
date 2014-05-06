VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Private Sub ComboBox1_DropButtonClick()
    Run Worksheets("SZusammentragen").Zusammentragen
End Sub

Private Sub Worksheet_Activate()
    Run Worksheets("SZusammentragen").Zusammentragen
End Sub
