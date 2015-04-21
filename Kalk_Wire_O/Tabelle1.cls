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
Option Explicit
Private Sub Worksheet_Activate()
    Application.ScreenUpdating = False
    Run Worksheets("Zusammentragen").Zusammentragen
    Run Worksheets("SSchneiden").Schneidbogen
    Run Worksheets("SSchneiden").SchneidenAnzeige
    Run Worksheets("Zusammentragen").Formatfehler
    Application.ScreenUpdating = True
    Call Materialkommentar
End Sub
Private Sub Worksheet_Deactivate()
    Application.ScreenUpdating = False
    Run Worksheets("Zusammentragen").Zusammentragen
    Run Worksheets("SSchneiden").Schneidbogen
    Run Worksheets("SSchneiden").SchneidenAnzeige
    Run Worksheets("Zusammentragen").Formatfehler
    Application.ScreenUpdating = True
    Call Materialkommentar
End Sub
Private Sub CheckBox1_Click()
On Error Resume Next
Application.ScreenUpdating = False
' Auswahl Beschneiden
    If CheckBox1.Value = True Then
        Worksheets("Eingabe").Range("G58") = "1"
    Else
        Worksheets("Eingabe").Range("G58") = "0"
    End If
    Run Worksheets("SSchneiden").Schneidbogen
    Run Worksheets("SSchneiden").SchneidenAnzeige
Application.ScreenUpdating = True
End Sub
Private Sub CheckBox2_Click()
On Error Resume Next
Application.ScreenUpdating = False
' Auswahl Zusammentragen
    If CheckBox2.Value = True Then
        Worksheets("Eingabe").Range("G60") = "1"
    Else: Worksheets("Eingabe").Range("G60") = "0"
    End If
Application.ScreenUpdating = True
End Sub
Private Sub CheckBox3_Click()
On Error Resume Next
Application.ScreenUpdating = False
' Auswahl 3Seitenbeschnitt
    If CheckBox3.Value = True Then
        Worksheets("Eingabe").Range("G62") = "1"
    Else: Worksheets("Eingabe").Range("G62") = "0"
    End If
Application.ScreenUpdating = True
End Sub
Private Sub CheckBox4_Click()
On Error Resume Next
Application.ScreenUpdating = False
' Auswahl Stanzen
    If CheckBox4.Value = True Then
        Worksheets("Eingabe").Range("G64") = "1"
    Else: Worksheets("Eingabe").Range("G64") = "0"
    End If
Application.ScreenUpdating = True
End Sub
Private Sub CheckBox5_Click()
On Error Resume Next
Application.ScreenUpdating = False
' Auswahl Binden
    If CheckBox5.Value = True Then
        Worksheets("Eingabe").Range("G66") = "1"
    Else: Worksheets("Eingabe").Range("G66") = "0"
    End If
Application.ScreenUpdating = True
End Sub
Private Sub CheckBox6_Click()
On Error Resume Next
Application.ScreenUpdating = False
' Auswahl Verpacken
    If CheckBox6.Value = True Then
        Worksheets("Eingabe").Range("G70") = "1"
    Else: Worksheets("Eingabe").Range("G70") = "0"
    End If
Application.ScreenUpdating = True
End Sub
Sub CheckBox7_Click()
    On Error Resume Next
    Application.ScreenUpdating = False
    ' Auswahl Bohren
        If CheckBox7.Value = True Then
            Worksheets("Eingabe").Range("G68") = "1"
            Worksheets("Bohren").Visible = True
        Else
            Worksheets("Eingabe").Range("G68") = "0"
            Worksheets("Bohren").Visible = False
        End If
    Application.ScreenUpdating = True
End Sub
Private Sub ComboBox1_DropButtonClick()
'Grammatur Folie
    If Worksheets("Steuerung").Range("D70") = 1 Then
        CommandButton5.Visible = True 'Formatänderung Alternativ-Bogen
        Else: CommandButton5.Visible = False
    End If
    Call NutzenCheck_Folie
End Sub
Private Sub ComboBox2_DropButtonClick()
'Grammatur Rückpappe
    If Worksheets("Steuerung").Range("H70") = 1 Then
        CommandButton1.Visible = True 'Formatänderung Alternativ-Bogen
        Else: CommandButton1.Visible = False
    End If
    Call NutzenCheck_Pappe
End Sub
Private Sub ComboBox3_DropButtonClick()
'Verarbeitungsart Folie
    Application.ScreenUpdating = False
    Call Alternativfolie
    Run Worksheets("SSchneiden").Schneidbogen
    Application.ScreenUpdating = True
End Sub
Private Sub ComboBox4_DropButtonClick()
'Verarbeitungsart Deckblatt
    Application.ScreenUpdating = False
    Run Worksheets("SSchneiden").Schneidbogen
    Application.ScreenUpdating = True
End Sub
Private Sub ComboBox5_DropButtonClick()
'Verarbeitungsart Inhalt
    Application.ScreenUpdating = False
    Run Worksheets("SSchneiden").Schneidbogen
    Application.ScreenUpdating = True
End Sub
Private Sub ComboBox6_DropButtonClick()
'Verarbeitungsart Rückblatt
    Application.ScreenUpdating = False
    Run Worksheets("SSchneiden").Schneidbogen
    Application.ScreenUpdating = True
End Sub
Private Sub ComboBox7_DropButtonClick()
'Verarbeitungsart Rückpappe
    Application.ScreenUpdating = False
    Call Rueckpappe
    Run Worksheets("SSchneiden").Schneidbogen
    Application.ScreenUpdating = True
End Sub
Private Sub ComboBox8_DropButtonClick()
    Call NutzenCheck_Folie
End Sub
Private Sub ComboBox12_DropButtonClick()
'Nutzen Pappe
    Call NutzenCheck_Pappe
End Sub
Private Sub ComboBox13_Change()
'Auswahl Buchblockformat
'20150421
    Run Worksheets("Eingabe").Produktformat
End Sub
Private Sub CommandButton1_Click()
'Nutzen Folie
    Call Bogenformat_Pappe
End Sub
Private Sub CommandButton3_Click()
'Speichern
    Call Speichern
End Sub
Private Sub CommandButton5_Click()
'Bogenformat Folie
    Call Bogenformat_Folie
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

