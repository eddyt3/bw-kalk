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
Application.ScreenUpdating = False
    If Worksheets("Eingabe").CheckBox1.Value = True Then
        Worksheets("SEingabe").Range("B25") = "1"
        Worksheets("Schneiden").CheckBox1.Value = True
        Worksheets("Schneiden").CheckBox1.Visible = True
        Worksheets("Schneiden").CheckBox2.Value = True
        Worksheets("Schneiden").CheckBox2.Visible = True
        Worksheets("Schneiden").CheckBox3.Value = True
        Worksheets("Schneiden").CheckBox3.Visible = True
        Worksheets("Schneiden").CheckBox4.Value = True
        Worksheets("Schneiden").CheckBox4.Visible = True
    Else
        Worksheets("SEingabe").Range("B25") = "0"
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
Application.ScreenUpdating = False
    If Worksheets("Eingabe").CheckBox2.Value = True Then
        Worksheets("SEingabe").Range("B26") = "1"
    Else: Worksheets("SEingabe").Range("B26") = "0"
    End If
Application.ScreenUpdating = True
End Sub
Sub CheckBox3_Click()
Attribute CheckBox3_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Heften
Application.ScreenUpdating = False
    If Worksheets("Eingabe").CheckBox3.Value = True Then
        Worksheets("SEingabe").Range("B27") = "1"
    Else: Worksheets("SEingabe").Range("B27") = "0"
    End If
Application.ScreenUpdating = True
End Sub
Sub CheckBox4_Click()
Attribute CheckBox4_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Verpacken
Application.ScreenUpdating = False
    If Worksheets("Eingabe").CheckBox4.Value = True Then
        Worksheets("SEingabe").Range("B29") = "1"
    Else: Worksheets("SEingabe").Range("B29") = "0"
    End If
Application.ScreenUpdating = True
End Sub
Sub CheckBox5_Click()
Attribute CheckBox5_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Bogen B
Application.ScreenUpdating = False
    If Worksheets("Eingabe").CheckBox5.Value = True Then
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
Application.ScreenUpdating = False
    If Worksheets("Eingabe").CheckBox6.Value = True Then
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
Application.ScreenUpdating = False
    If Worksheets("Eingabe").CheckBox7.Value = True Then
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
    Application.ScreenUpdating = False
    ' Auswahl Bohren
        If Worksheets("Eingabe").CheckBox8.Value = True Then
            Worksheets("SEingabe").Range("B28") = "1"
            Worksheets("Bohren").Visible = True
        Else
            Worksheets("SEingabe").Range("B28") = "0"
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
    If KeyCode = vbKeyTab Then
        Worksheets("Eingabe").TextBox3.Activate
        With Worksheets("Eingabe").TextBox3
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Sub
Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Produktlänge
    Select Case KeyAscii
      Case 48 To 57, 43 To 46
      Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub TextBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Tabsprung zur nächsten Eingabe
    If KeyCode = vbKeyTab Then Worksheets("Eingabe").Range("C10").Activate
End Sub
Sub Produktformat()
'Auswahl Produktformat
'20150420
    If Worksheets("SEingabe").Range("F27") = 7 Then 'Sonderformat
        Worksheets("Eingabe").TextBox2.Visible = True
        Worksheets("Eingabe").TextBox3.Visible = True
        With Worksheets("Eingabe").Range("E9")
            .Interior.ColorIndex = 6
            .Value = "x"
        End With
        Worksheets("Eingabe").Range("F9") = "Sonderformat"
    Else
        Worksheets("Eingabe").TextBox2.Visible = False 'Standardformate
        Worksheets("Eingabe").TextBox3.Visible = False
        With Worksheets("Eingabe").Range("E9")
            .Interior.ColorIndex = 0
            .Value = ""
        End With
        Worksheets("Eingabe").Range("E9").Value = Worksheets("SEingabe").Range("G26").Value
        Worksheets("Eingabe").Range("F9") = ""
    End If
End Sub

