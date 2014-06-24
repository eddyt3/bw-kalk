VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Sub Standardwerte_3Seitenbeschnitt()
    '20140623
    Application.ScreenUpdating = False
    Worksheets("S3Seitenbeschnitt").Range("C4:C18").Copy Destination:=Worksheets("3Seitenbeschnitt").Range("K3")
    Application.ScreenUpdating = True
End Sub
