VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'

Private Sub CommandButton1_Click()
'Materialseite ausblenden
    Worksheets("Material").Visible = False
    Worksheets("Verpacken").Activate
    Worksheets("Verpacken").Range("D3").Select
End Sub

