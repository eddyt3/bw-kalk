VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Private Sub Workbook_BeforePrint(Cancel As Boolean)
    'Druckbereich festlegen
    '14.01.2009
    Dim Letzte As Long 'Letzte Zeile
    Dim a As Integer
    Letzte = Cells(Rows.Count, 1).End(xlUp).Row
    If Worksheets("Eingabe").TextBox1.Visible = True Then
        Letzte = Letzte + 10
    End If
    Worksheets("Eingabe").PageSetup.PrintArea = "A1:L" & Letzte
    'Tabellen formatieren u. Doku-Eigenschaften setzen
    '25.03.2014
    Application.ScreenUpdating = False
    Dim wks As Worksheet
    Dim v As String
    Dim checkdate As String
    v = Worksheets("Steuerung").Range("B178") & "/" & Worksheets("Steuerung").Range("A178")
    checkdate = Worksheets("Steuerung").Range("B179")
    For Each wks In Worksheets
        With wks.PageSetup
            '.LeftHeader = ActiveWorkbook.Name
            '.RightHeader = Format(Date, "dd.mmmm.yyyy")
            .LeftFooter = "&""Verdana""&06" & Application.UserName & " " & format("&D", "dd.mm.yy") & " " & "&T" & " / Fehlerprfg.: " & checkdate
            .CenterFooter = "&""Verdana""&06" & Chr(10) & ActiveWorkbook.Path & "\" & ActiveWorkbook.Name
            .RightFooter = "&""Verdana""&06" & " (v" & v & ")"
        End With
    Next wks
    Call Dokumenteigenschaften_Soll
    Application.ScreenUpdating = True
End Sub
