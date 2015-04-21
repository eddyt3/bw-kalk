VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Private Sub Worksheet_Change(ByVal Target As Excel.Range)
    On Error Resume Next
    Application.ScreenUpdating = False
    Run Worksheets("SSammelheften").Zusammentragen
    'Run Worksheets("Falzen").Falzen
    Application.ScreenUpdating = True
End Sub
Private Sub ComboBox1_Change()
'Bogen A, Art
    'Run Worksheets("Falzen").Falzen
End Sub
Private Sub ComboBox2_Change()
'Bogen B, Art
    'Run Worksheets("Falzen").Falzen
End Sub
Private Sub ComboBox3_Change()
'Bogen C, Art
    'Run Worksheets("Falzen").Falzen
End Sub
Private Sub ComboBox4_Change()
'Bogen D, Art
    'Run Worksheets("Falzen").Falzen
End Sub
Private Sub ComboBox5_Change()
'Auswahl Produktformat
'20150420
    Run Worksheets("Eingabe").Produktformat
End Sub
Sub CheckBox1_Click()
' Auswahl Schneiden
On Error Resume Next
Application.ScreenUpdating = False
    If CheckBox1.Value = True Then
        Worksheets("Eingabe").Range("G60") = "1"
        Worksheets("Schneiden").CheckBox1.Value = True
        Worksheets("Schneiden").CheckBox1.Visible = True
        Worksheets("Schneiden").CheckBox2.Value = True
        Worksheets("Schneiden").CheckBox2.Visible = True
        Worksheets("Schneiden").CheckBox3.Value = True
        Worksheets("Schneiden").CheckBox3.Visible = True
        Worksheets("Schneiden").CheckBox4.Value = True
        Worksheets("Schneiden").CheckBox4.Visible = True
    Else
        Worksheets("Eingabe").Range("G60") = "0"
        Worksheets("Schneiden").CheckBox1.Value = False
        Worksheets("Schneiden").CheckBox1.Visible = False
        Worksheets("Schneiden").CheckBox2.Value = False
        Worksheets("Schneiden").CheckBox2.Visible = False
        Worksheets("Schneiden").CheckBox3.Value = False
        Worksheets("Schneiden").CheckBox3.Visible = False
        Worksheets("Schneiden").CheckBox4.Value = False
        Worksheets("Schneiden").CheckBox4.Visible = False
    End If
    Run Worksheets("SSchneiden").Schneidbogen
    Run Worksheets("SSchneiden").SchneidenAnzeige
Application.ScreenUpdating = True
End Sub
Sub CheckBox2_Click()
Attribute CheckBox2_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Falzen
On Error Resume Next
Application.ScreenUpdating = False
    If CheckBox2.Value = True Then
        Worksheets("Eingabe").Range("G62") = "1"
    Else: Worksheets("Eingabe").Range("G62") = "0"
    End If
Application.ScreenUpdating = True
End Sub
Sub CheckBox3_Click()
Attribute CheckBox3_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Heften
On Error Resume Next
Application.ScreenUpdating = False
    If CheckBox3.Value = True Then
        Worksheets("Eingabe").Range("G64") = "1"
    Else: Worksheets("Eingabe").Range("G64") = "0"
    End If
Application.ScreenUpdating = True
End Sub
Sub CheckBox4_Click()
Attribute CheckBox4_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Verpacken
On Error Resume Next
Application.ScreenUpdating = False
    If CheckBox4.Value = True Then
        Worksheets("Eingabe").Range("G68") = "1"
    Else: Worksheets("Eingabe").Range("G68") = "0"
    End If
Application.ScreenUpdating = True
End Sub
Sub CheckBox5_Click()
Attribute CheckBox5_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Bogen B
On Error Resume Next
Application.ScreenUpdating = False
    If CheckBox5.Value = True Then
        Worksheets("Eingabe").Range("A18") = "1"
        Worksheets("Schneiden").CheckBox2.Value = True
        Worksheets("Schneiden").CheckBox2.Visible = True
        Else
        Worksheets("Eingabe").Range("A18") = "0"
        Worksheets("Schneiden").CheckBox2.Value = False
        Worksheets("Schneiden").CheckBox2.Visible = False
        Worksheets("SFalzen").Range("I5") = 6
    End If
    Run Worksheets("SSchneiden").Schneidbogen
    Run Worksheets("SSchneiden").SchneidenAnzeige
    'Run Worksheets("Falzen").Falzen
Application.ScreenUpdating = True
End Sub
Sub CheckBox6_Click()
Attribute CheckBox6_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Bogen C
On Error Resume Next
Application.ScreenUpdating = False
    If CheckBox6.Value = True Then
        Worksheets("Eingabe").Range("A28") = "1"
        Worksheets("Schneiden").CheckBox3.Value = True
        Worksheets("Schneiden").CheckBox3.Visible = True
        Else
        Worksheets("Eingabe").Range("A28") = "0"
        Worksheets("Schneiden").CheckBox3.Value = False
        Worksheets("Schneiden").CheckBox3.Visible = False
        Worksheets("SFalzen").Range("I6") = 6
    End If
    Run Worksheets("SSchneiden").Schneidbogen
    Run Worksheets("SSchneiden").SchneidenAnzeige
    'Run Worksheets("Falzen").Falzen
Application.ScreenUpdating = True
End Sub
Sub CheckBox7_Click()
Attribute CheckBox7_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Bogen D
On Error Resume Next
Application.ScreenUpdating = False
    If CheckBox7.Value = True Then
        Worksheets("Eingabe").Range("A38") = "1"
        Worksheets("Schneiden").CheckBox4.Value = True
        Worksheets("Schneiden").CheckBox4.Visible = True
        Else
        Worksheets("Eingabe").Range("A38") = "0"
        Worksheets("Schneiden").CheckBox4.Value = False
        Worksheets("Schneiden").CheckBox4.Visible = False
        Worksheets("SFalzen").Range("I7") = 6
    End If
    Run Worksheets("SSchneiden").Schneidbogen
    Run Worksheets("SSchneiden").SchneidenAnzeige
    'Run Worksheets("Falzen").Falzen
Application.ScreenUpdating = True
End Sub
Sub CheckBox8_Click()
    On Error Resume Next
    Application.ScreenUpdating = False
    ' Auswahl Bohren
        If CheckBox8.Value = True Then
            Worksheets("Eingabe").Range("G66") = "1"
            Worksheets("Bohren").Visible = True
        Else
            Worksheets("Eingabe").Range("G66") = "0"
            Worksheets("Bohren").Visible = False
        End If
    Application.ScreenUpdating = True
End Sub
Private Sub CommandButton3_DropButtonClick()
'Speichern
    Application.ScreenUpdating = False
    Call Speichern
    Application.ScreenUpdating = True
End Sub
Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Produktbreite
    Select Case KeyAscii
      Case 48 To 57, 43 To 46
      Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Tabsprung zur nächsten Eingabe
    If KeyCode = vbKeyTab Then TextBox3.Activate
End Sub
Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Produktlänge
    Select Case KeyAscii
      Case 48 To 57, 43 To 46
      Case Else: KeyAscii = 0
    End Select
End Sub
Sub Produktformat()
'Auswahl Produktformat
'20150420
    If Worksheets("SEingabe").Range("F27") = 7 Then 'Sonderformat
        TextBox2.Visible = True
        TextBox3.Visible = True
        With Worksheets("Eingabe").Range("E9")
            .Interior.ColorIndex = 6
            .Value = "x"
        End With
        Worksheets("Eingabe").Range("F9") = "Sonderformat"
    Else
        TextBox2.Visible = False 'Standardformate
        TextBox3.Visible = False
        With Worksheets("Eingabe").Range("E9")
            .Interior.ColorIndex = 0
            .Value = ""
        End With
        Worksheets("Eingabe").Range("E9").Value = Worksheets("SEingabe").Range("G26").Value
        Worksheets("Eingabe").Range("F9") = ""
    End If
End Sub

