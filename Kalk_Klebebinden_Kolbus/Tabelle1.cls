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
Private Sub ComboBox2_Change()
'Bindeseite
    Run Worksheets("SKleben").kleben
End Sub
Private Sub ComboBox3_DropButtonClick()
    '8seitiger Umschlag
    If Worksheets("SEingabe").Range("L5") > 1 Then
        MsgBox ("Bitte das Falzen der Umschläge nicht vergessen.")
    End If
End Sub
Private Sub ComboBox4_Change()
'Bogen A, Art
    'Run Worksheets("Falzen").Falzen
End Sub
Private Sub ComboBox5_DropButtonClick()
    Run Worksheets("SZusammentragen").Zusammentragen
End Sub
Private Sub ComboBox9_Change()
'Bogen B, Art
    'Run Worksheets("Falzen").Falzen
End Sub
Private Sub ComboBox10_Change()
'Bogen C, Art
    'Run Worksheets("Falzen").Falzen
End Sub
Private Sub ComboBox11_Change()
'Bogen D, Art
    'Run Worksheets("Falzen").Falzen
End Sub
Private Sub CommandButton1_Click()
    Application.ScreenUpdating = False
        Call Fehlerpruefung
        Call checkdate
    Application.ScreenUpdating = True
End Sub
Private Sub CommandButton2_Click()
    Application.ScreenUpdating = False
        Call Speichern
    Application.ScreenUpdating = True
End Sub
Private Sub CommandButton3_Click()
'Eingabe Produktformat
    Run Worksheets("SEingabe").Formatinput
End Sub

Private Sub CommandButton4_Click()
    'Materialseite einblenden
    On Error Resume Next
    Worksheets("Material").Visible = xlSheetVisible
    Worksheets("Material").Activate
    Worksheets("Material").Range("C5").Select
End Sub
Private Sub Worksheet_Change(ByVal Target As Excel.Range)
    Application.ScreenUpdating = False
        On Error Resume Next
        Call Fehlermeldung
        Run Worksheets("SZusammentragen").Zusammentragen
        'Run Worksheets("Falzen").Falzen
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
            .Value = "Fehlermeldungen:" & vbCrLf & "=============" & vbCrLf & FFormat & vbCrLf & FFormatMin
            '.Locked = True
        End With
    End If
Worksheets("Eingabe").Protect "bw"
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
    On Error Resume Next
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
    On Error Resume Next
    Application.ScreenUpdating = False
        If CheckBox6.Value = True Then
            Worksheets("Eingabe").Range("G71") = "1"
        Else: Worksheets("Eingabe").Range("G71") = "0"
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox7_Click()
Attribute CheckBox7_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Verpacken
    On Error Resume Next
    Application.ScreenUpdating = False
        If CheckBox7.Value = True Then
            Worksheets("Eingabe").Range("G75") = "1"
        Else: Worksheets("Eingabe").Range("G75") = "0"
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox8_Click()
' Auswahl Bohren
    On Error Resume Next
    Application.ScreenUpdating = False
        If CheckBox8.Value = True Then
            Worksheets("Eingabe").Range("G73") = "1"
            Worksheets("Bohren").Visible = True
        Else
            Worksheets("Eingabe").Range("G73") = "0"
            Worksheets("Bohren").Visible = False
        End If
    Application.ScreenUpdating = True
End Sub
Private Sub CheckBox9_Click()
' Auswahl Kleben
    On Error Resume Next
    Application.ScreenUpdating = False
        If CheckBox9.Value = True Then
            Worksheets("Eingabe").Range("G67") = "1"
        Else: Worksheets("Eingabe").Range("G67") = "0"
        End If
    Run Worksheets("SKleben").kleben
    Application.ScreenUpdating = True
End Sub
Private Sub CheckBox10_Click()
' Auswahl Fadenheften
    On Error Resume Next
    Application.ScreenUpdating = False
        If CheckBox10.Value = True Then
            Worksheets("Eingabe").Range("G69") = "1"
        Else: Worksheets("Eingabe").Range("G69") = "0"
        End If
    Application.ScreenUpdating = True
End Sub
Private Sub ComboBox13_Change()
'Auswahl Leimart über Eingabeseite
'20140422
    On Error Resume Next
    Run Worksheets("Binden").Leimart
End Sub
Private Sub ComboBox14_Change()
'Auswahl Produktart (Broschüre, Buchblock, Ableimen)
'20140424
    'On Error Resume Next
    Dim intArt As Integer
    intArt = Worksheets("SBinden").Range("B34")
    Select Case intArt
    Case 1
        'Broschüre
        Worksheets("Kleben").Visible = False
        CheckBox9.Value = False 'Kalk. Kleben deaktiviert
        Worksheets("Eingabe").Range("G67") = "0"
        CheckBox10.Value = True 'Kalk. Fadenheften aktiviert
        Worksheets("Eingabe").Range("G69") = "0"
        Application.Run Worksheets("Binden").Produktart
    Case 2
        'Buchblock
        Worksheets("Kleben").Visible = True
        CheckBox9.Value = True 'Kalk. Kleben aktiviert
        Worksheets("Eingabe").Range("G67") = "1"
        CheckBox10.Value = True 'Kalk. Fadenheften aktiviert
        Worksheets("Eingabe").Range("G69") = "1"
        Application.Run Worksheets("Binden").Produktart
    Case 3
        'Ableimen
        Worksheets("Kleben").Visible = False
        CheckBox9.Value = False 'Kalk. Kleben deaktiviert
        Worksheets("Eingabe").Range("G67") = "0"
        CheckBox10.Value = False 'Kalk. Fadenheften deaktiviert
        Worksheets("Eingabe").Range("G69") = "0"
        Application.Run Worksheets("Binden").Produktart
    End Select

End Sub
Sub test()
    Run Worksheets("Binden").Produktart
End Sub
