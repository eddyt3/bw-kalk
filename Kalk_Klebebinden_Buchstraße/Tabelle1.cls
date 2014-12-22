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
Sub ComboBox2_Change()
    'Bindeseite
    Run Worksheets("SKleben").kleben
End Sub
Sub ComboBox3_DropButtonClick()
    'Buchdecke
    Run Worksheets("SBuchdecke").Thickness
    Run Worksheets("SBuchdecke").Beschriftung
End Sub
Private Sub ComboBox4_Change()
    'Bogen A, Art
    'Run Worksheets("Falzen").Falzen
End Sub
Sub ComboBox5_DropButtonClick()
    'Zusammentragen Autom./Hand
    Run Worksheets("Zusammentragen").Zusammentragen
End Sub
Sub ComboBox9_Change()
    'Bogen B, Art
    'Run Worksheets("Falzen").Falzen
End Sub
Sub ComboBox10_Change()
    'Bogen C, Art
    'Run Worksheets("Falzen").Falzen
End Sub
Sub ComboBox11_Change()
'Bogen D, Art
    'Run Worksheets("Falzen").Falzen
End Sub
Sub CommandButton1_Click()
    Application.ScreenUpdating = False
        Call Fehlerpruefung
        Call checkdate
    Application.ScreenUpdating = True
End Sub
Sub CommandButton2_Click()
    Application.ScreenUpdating = False
        Call Speichern
    Application.ScreenUpdating = True
End Sub
Sub CommandButton4_Click()
    'Materialseite einblenden
    On Error Resume Next
    Worksheets("Material").Visible = xlSheetVisible
    Worksheets("Material").Activate
    Worksheets("Material").Range("C5").Select
End Sub
Sub Worksheet_Change(ByVal Target As Excel.Range)
    Application.ScreenUpdating = False
        On Error Resume Next
        Call Fehlermeldung
        Run Worksheets("Zusammentragen").Zusammentragen
        'Run Worksheets("Falzen").Falzen
        Run Worksheets("SBuchdecke").Beschriftung
    Application.ScreenUpdating = True
End Sub
Sub Fehlerpruefung()
Attribute Fehlerpruefung.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.ScreenUpdating = False
        On Error Resume Next
        Call FormatMin
        Call Fehlermeldung
        If FFormat Like "" And FFormatMin Like "" Then
                MsgBox ("Alle Eingaben scheinen richtig zu sein." & vbCrLf & vbCrLf & _
                "Hinweis: Dies ist keine Auftragsprüfung!")
                Else
        End If
    Application.ScreenUpdating = True
End Sub
Sub Fehlermeldung()
Attribute Fehlermeldung.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error Resume Next
    If FFormat Like "" And FFormatMin Like "" Then
        TextBox1.Visible = False
        Else
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
            .Value = "Fehlermeldungen:" & vbCrLf & "=============" & vbCrLf & FFormat & vbCrLf & FFormatMin
            '.Locked = True
        End With
    End If
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
        Call Fehlerpruefung
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
       Call Fehlerpruefung
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
        Call Fehlerpruefung
    Application.ScreenUpdating = True
End Sub
Sub CheckBox4_Click()
Attribute CheckBox4_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Schneiden
    On Error Resume Next
    Application.ScreenUpdating = False
        If CheckBox4.Value = True Then 'Schneiden
            Worksheets("Eingabe").Range("G61") = "1"
            Worksheets("Schneiden").CheckBox1.Value = False 'Nur Umschlag soll standardmäßig geschnitte werden
            Worksheets("Schneiden").CheckBox1.Visible = True
            Worksheets("Schneiden").CheckBox2.Value = False
            Worksheets("Schneiden").CheckBox2.Visible = True
            Worksheets("Schneiden").CheckBox3.Value = False
            Worksheets("Schneiden").CheckBox3.Visible = True
            Worksheets("Schneiden").CheckBox4.Value = True 'Umschlag
            Worksheets("Schneiden").CheckBox4.Visible = True
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
' Auswahl Falzen
    Application.ScreenUpdating = False
        If CheckBox5.Value = True Then
            Worksheets("Eingabe").Range("G63") = "1"
        Else: Worksheets("Eingabe").Range("G63") = "0"
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox6_Click()
Attribute CheckBox6_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Kolbus
    Application.ScreenUpdating = False
        If CheckBox6.Value = True Then
            Worksheets("Eingabe").Range("G73") = "1"
        Else: Worksheets("Eingabe").Range("G73") = "0"
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox7_Click()
Attribute CheckBox7_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Verpacken
    Application.ScreenUpdating = False
        If CheckBox7.Value = True Then
            Worksheets("Eingabe").Range("G79") = "1"
        Else: Worksheets("Eingabe").Range("G79") = "0"
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox8_Click()
' Auswahl Buchdecke
    Application.ScreenUpdating = False
        If CheckBox8.Value = True Then
            Worksheets("Eingabe").Range("G75") = "1"
            Worksheets("Buchdecke").Visible = True
        Else
            Worksheets("Eingabe").Range("G75") = "0"
            Worksheets("Buchdecke").Visible = False
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox9_Click()
' Auswahl Kleben
    Application.ScreenUpdating = False
        If CheckBox9.Value = True Then
            Worksheets("Eingabe").Range("G67") = "1"
        Else: Worksheets("Eingabe").Range("G67") = "0"
        End If
    Run Worksheets("SKleben").kleben
    Application.ScreenUpdating = True
End Sub
Sub CheckBox10_Click()
' Auswahl Fadenheften
    Application.ScreenUpdating = False
        If CheckBox10.Value = True Then
            Worksheets("Eingabe").Range("G71") = "1"
            Worksheets("Fadenheften").Visible = True
        Else
            Worksheets("Eingabe").Range("G71") = "0"
            Worksheets("Fadenheften").Visible = False
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox11_Click()
'Auswahl Zusammentragen
'20140516
    If CheckBox11.Value = True Then
        Worksheets("Eingabe").Range("G69") = "1"
    Else: Worksheets("Eingabe").Range("G69") = "0"
    End If
End Sub
Sub CheckBox12_Click()
    'Auswahl Einhängen
    Application.ScreenUpdating = False
        If CheckBox12.Value = True Then
            Worksheets("Eingabe").Range("G77") = "1"
            Worksheets("Einhängen").Visible = True
        Else
            Worksheets("Eingabe").Range("G77") = "0"
            Worksheets("Einhängen").Visible = False
        End If
    Application.ScreenUpdating = True
End Sub
Sub ComboBox13_Change()
'Auswahl Leimart über Eingabeseite
'20140422
    On Error Resume Next
    Run Worksheets("Binden").Leimart
End Sub
Sub ComboBox14_Change()
'Auswahl Produktart (Broschüre, Buchblock, Ableimen)
'20140424
    'On Error Resume Next
    Dim intArt As Integer
    intArt = ThisWorkbook.Worksheets("SBinden").Range("B34").Value
    Select Case intArt
    Case 1
        'Broschüre
        ThisWorkbook.Worksheets("Kleben").Visible = False
        CheckBox9.Value = False 'Kalk. Kleben deaktiviert
        Worksheets("Eingabe").Range("G67") = "0"
        CheckBox10.Value = False 'Kalk. Fadenheften deaktiviert
        Worksheets("Eingabe").Range("G71") = "0"
        Worksheets("Fadenheften").Visible = False
        Worksheets("SZusammentragen").Range("B42") = "1" 'Kolbus ZU autom. Inline
        Worksheets("SZusammentragen").Range("C11") = "4" 'Kolbus ZU autom. Inline
    Case 2
        'Buchblock
        ThisWorkbook.Worksheets("Kleben").Visible = True
        CheckBox9.Value = True 'Kalk. Kleben aktiviert
        Worksheets("Eingabe").Range("G67") = "1"
        CheckBox10.Value = False 'Kalk. Fadenheften deaktiviert
        Worksheets("Eingabe").Range("G71") = "0"
        Worksheets("Fadenheften").Visible = False
        Worksheets("SZusammentragen").Range("B42") = "1" 'Kolbus ZU autom. Inline
        Worksheets("SZusammentragen").Range("C11") = "4" 'Kolbus ZU autom. Inline
    Case 3
        'Buchblock fadengeheftet
        ThisWorkbook.Worksheets("Kleben").Visible = True
        CheckBox9.Value = True 'Kalk. Kleben aktiviert
        Worksheets("Eingabe").Range("G67") = "1"
        CheckBox10.Value = True 'Kalk. Fadenheften aktiviert
        Worksheets("Eingabe").Range("G71") = "1"
        Worksheets("Fadenheften").Visible = True
        Worksheets("SZusammentragen").Range("B42") = "2" 'Kolbus ZU handanlage
        Worksheets("SZusammentragen").Range("C11") = "2" 'Zusammentragen auf ZTM370
    Case 4
        'Ableimen
        ThisWorkbook.Worksheets("Kleben").Visible = False
        CheckBox9.Value = False 'Kalk. Kleben deaktiviert
        Worksheets("Eingabe").Range("G67") = "0"
        CheckBox10.Value = False 'Kalk. Fadenheften deaktiviert
        Worksheets("Eingabe").Range("G71") = "0"
        ThisWorkbook.Worksheets("Fadenheften").Visible = False
        Worksheets("SZusammentragen").Range("B42") = "2" 'Kolbus ZU handanlage
        Worksheets("SZusammentragen").Range("C11") = "4" 'Kolbus ZU autom. Inline
    End Select
    Application.Run Worksheets("Binden").Produktart
End Sub
Private Sub ComboBox16_Change()
'Buchrückenform
    If Worksheets("SBuchdecke").Range("B38") = "4" Then
        Worksheets("Einhängen").CheckBox2.Value = True 'Runden beim Einhängen aktivieren
    Else: Worksheets("Einhängen").CheckBox2.Value = False
    End If
End Sub
Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Buchdecke, alternative Pappstärke mit Nachkommastelle
    Select Case KeyAscii
      Case 48 To 57, 43 To 46
      Case Else: KeyAscii = 0
    End Select
    Run Worksheets("SBuchdecke").Beschriftung
End Sub
Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Produktbreite
    Select Case KeyAscii
      Case 48 To 57
      Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Produktlänge
    Select Case KeyAscii
      Case 48 To 57
      Case Else: KeyAscii = 0
    End Select
End Sub
'Private Sub TextBox6_Change(ByVal KeyAscii As MSForms.ReturnInteger)
'' Bezug Nutzenanzahl
'    Select Case KeyAscii
'      Case 48 To 57
'      Case Else: KeyAscii = 0
'    End Select
'End Sub
Sub test()
    Run Worksheets("Binden").Produktart
End Sub
