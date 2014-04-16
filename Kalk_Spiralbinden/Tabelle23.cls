VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Sub Standardwerte_Stanzen()
'20111209
    Application.ScreenUpdating = False
    Worksheets("SStanzen").Select
    Worksheets("SStanzen").Range("B4:G9").Copy Destination:=Worksheets("Stanzen").Range("K5")
    Worksheets("Stanzen").Select
    Worksheets("Stanzen").Range("D6").Select
    Application.ScreenUpdating = True
End Sub
