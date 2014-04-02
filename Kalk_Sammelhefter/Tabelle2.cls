VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Private Sub Worksheet_Change(ByVal Target As Range)
On Error Resume Next
Worksheets("Schneiden").EnableOutlining = True
End Sub
Private Sub CheckBox1_Click()
 Run Worksheets("SSchneiden").SchneidenAnzeige
 Run Worksheets("SSchneiden").Schneidbogen
End Sub
Private Sub CheckBox2_Click()
 Run Worksheets("SSchneiden").SchneidenAnzeige
 Run Worksheets("SSchneiden").Schneidbogen
End Sub
Private Sub CheckBox3_Click()
 Run Worksheets("SSchneiden").SchneidenAnzeige
 Run Worksheets("SSchneiden").Schneidbogen
End Sub
Private Sub CheckBox4_Click()
 Run Worksheets("SSchneiden").SchneidenAnzeige
 Run Worksheets("SSchneiden").Schneidbogen
End Sub

