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
Private Sub Worksheet_Deactivate()
    Application.ScreenUpdating = False
    Run Worksheets("Zusammentragen").Zusammentragen
    Run Worksheets("SSchneiden").Schneidbogen
    Run Worksheets("SSchneiden").SchneidenAnzeige
    Run Worksheets("Zusammentragen").Formatfehler
    Application.ScreenUpdating = True
    Call Materialkommentar
    Call Fehlermeldung
End Sub
Private Sub Worksheet_Activate()
    Application.ScreenUpdating = False
    Run Worksheets("Zusammentragen").Zusammentragen
    Run Worksheets("SSchneiden").Schneidbogen
    Run Worksheets("SSchneiden").SchneidenAnzeige
    Run Worksheets("Zusammentragen").Formatfehler
    Application.ScreenUpdating = True
    Call Materialkommentar
    Call Fehlermeldung
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
Private Sub ComboBox12_DropButtonClick()
'Nutzen Pappe
    Call NutzenCheck_Pappe
End Sub
Private Sub ComboBox8_DropButtonClick()
    Call NutzenCheck_Folie
End Sub
Private Sub CommandButton1_Click()
'Nutzen Folie
    Call Bogenformat_Pappe
End Sub
Private Sub CommandButton2_Click()
'Eingabe Produktformat
    Run Worksheets("SEingabe").Formatinput
End Sub
Private Sub CommandButton3_Click()
'Speichern
    Call Speichern
End Sub
Private Sub CommandButton4_Click()
'Fehlerprüfung
    Call Fehlerpruefung
    Call checkdate
End Sub
Private Sub CommandButton5_Click()
'Bogenformat Folie
    Call Bogenformat_Folie
End Sub
Sub Fehlerpruefung()
    On Error Resume Next
    Call Fehlermeldung
    If FFormat Like "" And FSchlaufe Like "" And FSchlaufeS Like "" And FTeilung Like "" And FSchaftMin Like "" And _
    FSchaftMax Like "" And FStanzen Like "" Then
        MsgBox ("Alle Eingaben scheinen richtig zu sein." & vbCrLf & vbCrLf & _
        "Hinweis: Dies ist keine Auftragsprüfung!")
        Else
    End If
End Sub
Sub Fehlermeldung()
On Error Resume Next
If FFormat Like "" And FSchlaufe Like "" And FSchlaufeS Like "" And FTeilung Like "" And FSchaftMin Like "" And _
FSchaftMax Like "" And FStanzen Like "" Then
    TextBox1.Visible = False
    Else
    Worksheets("Eingabe").Unprotect "bw"
    With TextBox1
        .Visible = True
        .Font = "Verdana"
        .Font.Size = 10
        .Font.Bold = True
        .BorderColor = vbRed
        .BackColor = vbRed
        .Height = 120
        .Width = 525
        .MaxLength = 600
        .WordWrap = False
        .Value = "Fehlermeldungen:" & vbCrLf & "=============" & vbCrLf & FFormat & vbCrLf & FSchlaufe & vbCrLf & FSchlaufeS & vbCrLf _
        & FTeilung & vbCrLf & FSchaftMin & vbCrLf & FSchaftMax & vbCrLf & FStanzen
        '.Locked = True
    End With
End If
Worksheets("Eingabe").Protect "bw"
End Sub

