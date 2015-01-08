VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Private Sub CommandButton1_Click()
'Versionsstand aktualisieren
'14.01.2009
    Call version
End Sub
Private Sub CommandButton2_Click()
'Dokumenteigenschaften auflisten
'14.01.2009
    Call Dokumenteigenschaften_Ist
End Sub
Private Sub CommandButton3_Click()
'Dokumenteigenschaften setzen
'14.01.2009
    Call Dokumenteigenschaften_Soll
End Sub

