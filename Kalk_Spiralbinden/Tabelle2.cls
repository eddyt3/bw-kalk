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
Sub Schneidbogen()
' Loeschen d. vorhandenen Eingaben wenn d. Bg. nicht ausgewaehlt wurde
If Worksheets("SSchneiden").Range("B103") < 1 Then
    Worksheets("Schneiden").Range("D3") = "0"
    Else
    Worksheets("Schneiden").Range("D3").FormulaLocal = "=SVERWEIS(SSchneiden!C104;SSchneiden!I13:K30;3)"
End If
If Worksheets("SSchneiden").Range("C103") < 1 Then
    Worksheets("Schneiden").Range("D10") = "0"
    Else
    Worksheets("Schneiden").Range("D10").FormulaLocal = "=SVERWEIS(SSchneiden!C104;SSchneiden!I13:K30;3)"
End If
If Worksheets("SSchneiden").Range("D103") < 1 Then
    Worksheets("Schneiden").Range("D17") = "0"
    Else
    Worksheets("Schneiden").Range("D17").FormulaLocal = "=SVERWEIS(SSchneiden!D104;SSchneiden!I13:K30;3)"
End If
If Worksheets("SSchneiden").Range("E103") < 1 Then
    Worksheets("Schneiden").Range("D24") = "0"
    Else
    Worksheets("Schneiden").Range("D24").FormulaLocal = "=SVERWEIS(SSchneiden!E104;SSchneiden!I13:K30;3)"
End If
If Worksheets("SSchneiden").Range("F103") < 1 Then
    Worksheets("Schneiden").Range("D24") = "0"
    Else
    Worksheets("Schneiden").Range("D24").FormulaLocal = "=SVERWEIS(SSchneiden!E104;SSchneiden!I13:K30;3)"
End If
If Worksheets("SSchneiden").Range("F111").Value < 1 Then
    Worksheets("Schneiden").Range("D3") = "0"
    Worksheets("Schneiden").Range("D10") = "0"
    Worksheets("Schneiden").Range("D17") = "0"
    Worksheets("Schneiden").Range("D24") = "0"
    Worksheets("Schneiden").Range("D31") = "0"
    Else
    Worksheets("Schneiden").Range("D3").FormulaLocal = "=SVERWEIS(SSchneiden!B104;SSchneiden!I13:K30;3)"
 End If
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    Worksheets("Schneiden").EnableOutlining = True
End Sub
Private Sub CheckBox1_Click()
 Run Worksheets("SSchneiden").SchneidenAnzeige
End Sub
Private Sub CheckBox2_Click()
 Run Worksheets("SSchneiden").SchneidenAnzeige
End Sub
Private Sub CheckBox3_Click()
 Run Worksheets("SSchneiden").SchneidenAnzeige
End Sub
Private Sub CheckBox4_Click()
 Run Worksheets("SSchneiden").SchneidenAnzeige
End Sub
Private Sub CheckBox5_Click()
 Run Worksheets("SSchneiden").SchneidenAnzeige
End Sub
