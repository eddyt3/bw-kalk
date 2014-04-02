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
    On Error Resume Next
    Worksheets("Sammelheften").EnableOutlining = True
End Sub
Private Sub Worksheet_Activate()
    Run Worksheets("SSammelheften").Zusammentragen
End Sub

