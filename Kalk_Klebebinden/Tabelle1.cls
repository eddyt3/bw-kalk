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
Private Sub CommandButton2_Click()
    'Speichern
    Application.ScreenUpdating = False
    Call Speichern
    Application.ScreenUpdating = True
End Sub
Private Sub Worksheet_Change(ByVal Target As Excel.Range)
    On Error Resume Next
    Application.ScreenUpdating = False
    Run Worksheets("Zusammentragen").Zusammentragen
    'Run Worksheets("Falzen").Falzen
    Application.ScreenUpdating = True
End Sub
Private Sub ComboBox1_DropButtonClick()
    'Auswahl Broschure od. Buchblock
    On Error Resume Next
    ComboBox1.ListFillRange = "Steuerung!H2:H3"
    ComboBox1.LinkedCell = "A6"
    If Worksheets("Steuerung").Range("I3") = "1" Then 'Buchblock
        Worksheets("Kleben").Visible = True
        CheckBox7.Value = True 'Kalk. Kleben aktiviert
        Worksheets("Eingabe").Range("G69") = "1"
        CheckBox8.Value = True 'Kalk. Fadenheften aktiviert
        Worksheets("Eingabe").Range("G71") = "1"
    Else
        Worksheets("Kleben").Visible = False
        CheckBox7.Value = False 'Kalk. Kleben deaktiviert
        Worksheets("Eingabe").Range("G69") = "0"
        CheckBox8.Value = False 'Kalk. Fadenheften deaktiviert
        Worksheets("Eingabe").Range("G71") = "0"
    End If
End Sub
Sub ComboBox7_DropButtonClick()
    'Bogen A, Art
    'Run Worksheets("SKleben").kleben
    'Run Worksheets("Falzen").Falzen
End Sub
Sub ComboBox8_DropButtonClick()
    'Bogen B, Art
    'Run Worksheets("SKleben").kleben
    'Run Worksheets("Falzen").Falzen
End Sub
Sub ComboBox9_DropButtonClick()
    'Bogen C, Art
    'Run Worksheets("SKleben").kleben
    'Run Worksheets("Falzen").Falzen
End Sub
Sub ComboBox10_DropButtonClick()
    'Bogen D, Art
    'Run Worksheets("SKleben").kleben
    'Run Worksheets("Falzen").Falzen
End Sub
Private Sub ComboBox11_Change()
    '8seitiger Umschlag
    If Worksheets("SEingabe").Range("L5") > 1 Then
        MsgBox ("Bitte das Falzen der Umschläge nicht vergessen." & vbCrLf & vbCrLf & "Automatisch voreingestellt" _
        & " bei 8seitigen Umschlag:" & vbCrLf & "2 x 3 Seitenbeschnitt" & vbCrLf & "2 x Kleben & Grundeinrichten (Binden)")
        Worksheets("3Seitenbeschnitt").Range("D6") = Worksheets("3Seitenbeschnitt").Range("D5")
    Else
        Worksheets("3Seitenbeschnitt").Range("D6") = 1
    End If
End Sub
Private Sub ComboBox12_Change()
'Auswahl Buchblockformat
'20150421
    Run Worksheets("Eingabe").Produktformat
End Sub
Sub CheckBox1_Click()
Attribute CheckBox1_Click.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error Resume Next
    Application.ScreenUpdating = False
    ' Auswahl Bogen B
        If CheckBox1.Value = True Then
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
        Run Worksheets("SKleben").kleben
    Application.ScreenUpdating = True
End Sub
Sub CheckBox2_Click()
Attribute CheckBox2_Click.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error Resume Next
    Application.ScreenUpdating = False
    ' Auswahl Bogen C
        If CheckBox2.Value = True Then
            Worksheets("Eingabe").Range("A27") = "1"
            Worksheets("Schneiden").CheckBox3.Value = True
            Worksheets("Schneiden").CheckBox3.Visible = True
        Else
            Worksheets("Eingabe").Range("A27") = "0"
            Worksheets("Schneiden").CheckBox3.Value = False
            Worksheets("Schneiden").CheckBox3.Visible = False
            Worksheets("SFalzen").Range("I6") = 6
        End If
        Run Worksheets("SSchneiden").Schneidbogen
        Run Worksheets("SSchneiden").SchneidenAnzeige
        'Run Worksheets("Falzen").Falzen
       Run Worksheets("SKleben").kleben
    Application.ScreenUpdating = True
End Sub
Sub CheckBox3_Click()
Attribute CheckBox3_Click.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error Resume Next
    Application.ScreenUpdating = False
    ' Auswahl Bogen D
        If CheckBox3.Value = True Then
            Worksheets("Eingabe").Range("A36") = "1"
            Worksheets("Schneiden").CheckBox4.Value = True
            Worksheets("Schneiden").CheckBox4.Visible = True
        Else
            Worksheets("Eingabe").Range("A36") = "0"
            Worksheets("Schneiden").CheckBox4.Value = False
            Worksheets("Schneiden").CheckBox4.Visible = False
            Worksheets("SFalzen").Range("I7") = 6
       End If
        Run Worksheets("SSchneiden").Schneidbogen
        Run Worksheets("SSchneiden").SchneidenAnzeige
        'Run Worksheets("Falzen").Falzen
        Run Worksheets("SKleben").kleben
    Application.ScreenUpdating = True
End Sub
Sub CheckBox4_Click()
Attribute CheckBox4_Click.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error Resume Next
    Application.ScreenUpdating = False
    ' Auswahl Schneiden
        If CheckBox4.Value = True Then
            Worksheets("Eingabe").Range("G61") = "1"
            Worksheets("Schneiden").CheckBox1.Value = True
            Worksheets("Schneiden").CheckBox1.Visible = True
            If Worksheets("Eingabe").CheckBox5.Value = True Then
                Worksheets("Schneiden").CheckBox2.Value = True
                Worksheets("Schneiden").CheckBox2.Visible = True
            End If
            If Worksheets("Eingabe").CheckBox6.Value = True Then
                Worksheets("Schneiden").CheckBox3.Value = True
                Worksheets("Schneiden").CheckBox3.Visible = True
            End If
            If Worksheets("Eingabe").CheckBox7.Value = True Then
                Worksheets("Schneiden").CheckBox4.Value = True
                Worksheets("Schneiden").CheckBox4.Visible = True
            End If
        Else
            Worksheets("Eingabe").Range("G61") = "0"
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
Sub CheckBox5_Click()
Attribute CheckBox5_Click.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error Resume Next
    Application.ScreenUpdating = False
    ' Auswahl Zusammentragen
        If CheckBox5.Value = True Then
            Worksheets("Eingabe").Range("G63") = "1"
        Else: Worksheets("Eingabe").Range("G63") = "0"
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox6_Click()
Attribute CheckBox6_Click.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error Resume Next
    Application.ScreenUpdating = False
    ' Auswahl Falzen
        If CheckBox6.Value = True Then
            Worksheets("Eingabe").Range("G65") = "1"
        Else: Worksheets("Eingabe").Range("G65") = "0"
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox7_Click()
Attribute CheckBox7_Click.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error Resume Next
    Application.ScreenUpdating = False
    ' Auswahl Kleben
        If CheckBox7.Value = True Then
            Worksheets("Eingabe").Range("G69") = "1"
        Else: Worksheets("Eingabe").Range("G69") = "0"
        End If
    Run Worksheets("SKleben").kleben
    Application.ScreenUpdating = True
End Sub
Sub CheckBox8_Click()
Attribute CheckBox8_Click.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error Resume Next
    Application.ScreenUpdating = False
    ' Auswahl Fadenheften
        If CheckBox8.Value = True Then
            Worksheets("Eingabe").Range("G71") = "1"
        Else: Worksheets("Eingabe").Range("G71") = "0"
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox9_Click()
Attribute CheckBox9_Click.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error Resume Next
    Application.ScreenUpdating = False
    ' Auswahl Binden
        If CheckBox9.Value = True Then
            Worksheets("Eingabe").Range("G73") = "1"
        Else: Worksheets("Eingabe").Range("G73") = "0"
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox10_Click()
Attribute CheckBox10_Click.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error Resume Next
    Application.ScreenUpdating = False
    ' Auswahl 3Seitenbeschnitt
        If CheckBox10.Value = True Then
            Worksheets("Eingabe").Range("G75") = "1"
        Else: Worksheets("Eingabe").Range("G75") = "0"
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox11_Click()
Attribute CheckBox11_Click.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error Resume Next
    Application.ScreenUpdating = False
    ' Auswahl Verpacken
        If CheckBox11.Value = True Then
            Worksheets("Eingabe").Range("G79") = "1"
        Else: Worksheets("Eingabe").Range("G79") = "0"
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox12_Click()
    On Error Resume Next
    Application.ScreenUpdating = False
    ' Auswahl Bohren
        If CheckBox12.Value = True Then
            Worksheets("Eingabe").Range("G77") = "1"
            Worksheets("Bohren").Visible = True
        Else
            Worksheets("Eingabe").Range("G77") = "0"
            Worksheets("Bohren").Visible = False
        End If
    Application.ScreenUpdating = True
End Sub
Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Produktbreite
    Select Case KeyAscii
      Case 48 To 57, 43 To 46
      Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Tabsprung zur nächsten Eingabe
   If KeyCode = vbKeyTab Then TextBox2.Activate
End Sub
Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Produktlänge
    Select Case KeyAscii
      Case 48 To 57, 43 To 46
      Case Else: KeyAscii = 0
    End Select
End Sub
Sub Produktformat()
'Auswahl Produktformat
'20150221
    If Worksheets("SEingabe").Range("F27") = 7 Then 'Sonderformat
        TextBox1.Visible = True
        TextBox2.Visible = True
        With Worksheets("Eingabe").Range("E9")
            .Interior.ColorIndex = 6
            .Value = "x"
        End With
        Worksheets("Eingabe").Range("F9") = "Sonderformat"
    Else
        'Standardformate
        TextBox1.Visible = False
        TextBox2.Visible = False
        With Worksheets("Eingabe").Range("E9")
            .Interior.ColorIndex = 0
            .Value = ""
        End With
        Worksheets("Eingabe").Range("E9").Value = Worksheets("SEingabe").Range("G26").Value
        Worksheets("Eingabe").Range("F9") = ""
    End If
End Sub

