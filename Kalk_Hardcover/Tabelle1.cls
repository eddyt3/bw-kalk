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
'Standardwerte
'Buchrücken gerade Schrenz 600g
'Bindeseite lang
Private Sub Worksheet_Activate()
        Run Worksheets("SBuchdecke").Thickness
End Sub
Sub Worksheet_Change(ByVal Target As Excel.Range)
    Application.ScreenUpdating = False
        Run Worksheets("Zusammentragen").Zusammentragen
        'Run Worksheets("Falzen").Falzen
        Run Worksheets("SBuchdecke").Beschriftung
    Application.ScreenUpdating = True
End Sub
Sub ComboBox2_Change()
    'Bindeseite
    Run Worksheets("SKleben").kleben
End Sub
Sub ComboBox3_DropButtonClick()
    'Buchdecke
    Run Worksheets("SBuchdecke").Thickness
    Run Worksheets("SBuchdecke").Beschriftung
    Run Worksheets("Eingabe").Buchblockformat
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
Sub ComboBox13_Change()
'Auswahl Leimart über Eingabeseite
'20140422
    Run Worksheets("Binden").Leimart
End Sub
Sub ComboBox14_Change()
'Auswahl Produktart (Broschüre, Hardcover, Ableimen)
'20140424
    Dim intArt As Integer
    intArt = ThisWorkbook.Worksheets("SBinden").Range("B34").Value
    Select Case intArt
    Case 1
        'Broschüre
        ThisWorkbook.Worksheets("Kleben").Visible = False
        Worksheets("Eingabe").CheckBox9.Value = False 'Kalk. Kleben deaktiviert
        Worksheets("SEingabe").Range("B32") = "0"
        Worksheets("Eingabe").CheckBox10.Value = False 'Kalk. Fadenheften deaktiviert
        Worksheets("SEingabe").Range("B34") = "0"
        Worksheets("Fadenheften").Visible = False
        Worksheets("SZusammentragen").Range("B42") = "1" 'Kolbus ZU autom. Inline
        Worksheets("SZusammentragen").Range("C11") = "4" 'Kolbus ZU autom. Inline
    Case 2
        'Hardcover
        ThisWorkbook.Worksheets("Kleben").Visible = True
        Worksheets("Eingabe").CheckBox9.Value = True 'Kalk. Kleben aktiviert
        Worksheets("SEingabe").Range("B32") = "1"
        Worksheets("Eingabe").CheckBox10.Value = False 'Kalk. Fadenheften deaktiviert
        Worksheets("SEingabe").Range("B34") = "0"
        Worksheets("Fadenheften").Visible = False
        Worksheets("SZusammentragen").Range("B42") = "1" 'Kolbus ZU autom. Inline
        Worksheets("SZusammentragen").Range("C11") = "4" 'Kolbus ZU autom. Inline
    Case 3
        'Hardcover fadengeheftet
        ThisWorkbook.Worksheets("Kleben").Visible = True
        Worksheets("Eingabe").CheckBox9.Value = True 'Kalk. Kleben aktiviert
        Worksheets("SEingabe").Range("B32") = "1"
        Worksheets("Eingabe").CheckBox10.Value = True 'Kalk. Fadenheften aktiviert
        Worksheets("SEingabe").Range("B34") = "1"
        Worksheets("Fadenheften").Visible = True
        Worksheets("SZusammentragen").Range("B42") = "2" 'Kolbus ZU handanlage
        Worksheets("SZusammentragen").Range("C11") = "2" 'Zusammentragen auf ZTM370
    Case 4
        'Ableimen
        ThisWorkbook.Worksheets("Kleben").Visible = False
        Worksheets("Eingabe").CheckBox9.Value = False 'Kalk. Kleben deaktiviert
        Worksheets("SEingabe").Range("B32") = "0"
        Worksheets("Eingabe").CheckBox10.Value = False 'Kalk. Fadenheften deaktiviert
        Worksheets("SEingabe").Range("B34") = "0"
        Worksheets("Fadenheften").Visible = False
        Worksheets("SZusammentragen").Range("B42") = "1" 'Kolbus ZU autom. Inline
        Worksheets("SZusammentragen").Range("C11") = "4" 'Kolbus ZU autom. Inline
    End Select
    Worksheets("SZusammentragen").Calculate
    Application.Run Worksheets("Binden").Produktart
    Application.Run Worksheets("Zusammentragen").GeschwindigkeitZus
End Sub
Private Sub ComboBox16_Change()
'Buchrückenform
    If ThisWorkbook.Sheets("SBuchdecke").Range("B38") = 4 Then
        ThisWorkbook.Sheets("Einhängen").CheckBox2.Value = True 'Runden beim Einhängen aktivieren
    Else: ThisWorkbook.Sheets("Einhängen").CheckBox2.Value = False
    End If
End Sub
Private Sub ComboBox21_Change()
'Auswahl Buchblockformat
'20150220
    Run Worksheets("Eingabe").Buchblockformat
End Sub
Sub CommandButton2_Click()
    Application.ScreenUpdating = False
        Call Speichern
    Application.ScreenUpdating = True
End Sub
Sub CommandButton4_Click()
    'Materialseite einblenden
    Worksheets("Material").Visible = xlSheetVisible
    Worksheets("Material").Activate
    Worksheets("Material").Range("C5").Select
End Sub
Sub CheckBox1_Click()
Attribute CheckBox1_Click.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.ScreenUpdating = False
    ' Auswahl Bogen B
        If Worksheets("Eingabe").CheckBox1.Value = True Then
            Worksheets("Eingabe").Range("A18") = "1"
            Worksheets("Schneiden").CheckBox2.Value = True
            Worksheets("Schneiden").CheckBox2.Visible = True
            If Worksheets("Eingabe").CheckBox5.Value = True And Worksheets("Eingabe").ComboBox9.ListIndex = 0 Then 'Falzen und Inhalt aktiv?
                Worksheets("Falzen").ComboBox2.ListIndex = 0 'Bg. B Falzen aktivieren
            End If
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
Sub CheckBox2_Click()
Attribute CheckBox2_Click.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.ScreenUpdating = False
    ' Auswahl Bogen C
        If Worksheets("Eingabe").CheckBox2.Value = True Then
            Worksheets("Eingabe").Range("A27") = "1"
            Worksheets("Schneiden").CheckBox3.Value = True
            Worksheets("Schneiden").CheckBox3.Visible = True
            If Worksheets("Eingabe").CheckBox5.Value = True And Worksheets("Eingabe").ComboBox10.ListIndex = 0 Then 'Falzen und Inhalt aktiv?
                Worksheets("Falzen").ComboBox3.ListIndex = 0 'Bg. C Falzen aktivieren
            End If
        Else
            Worksheets("Eingabe").Range("A27") = "0"
            Worksheets("Schneiden").CheckBox3.Value = False
            Worksheets("Schneiden").CheckBox3.Visible = False
            Worksheets("SFalzen").Range("I6") = 6
        End If
        Run Worksheets("SSchneiden").Schneidbogen
        Run Worksheets("SSchneiden").SchneidenAnzeige
        'Run Worksheets("Falzen").Falzen
    Application.ScreenUpdating = True
End Sub
Sub CheckBox3_Click()
Attribute CheckBox3_Click.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.ScreenUpdating = False
    ' Auswahl Bogen D
        If Worksheets("Eingabe").CheckBox3.Value = True Then
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
    Application.ScreenUpdating = True
End Sub
Sub CheckBox4_Click()
Attribute CheckBox4_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Schneiden
    Application.ScreenUpdating = False
        If Worksheets("Eingabe").CheckBox4.Value = True Then 'Schneiden
            Worksheets("SEingabe").Range("B30") = "1"
            Worksheets("Schneiden").CheckBox1.Value = False 'Nur Umschlag soll standardmäßig geschnitte werden
            Worksheets("Schneiden").CheckBox1.Visible = True
            Worksheets("Schneiden").CheckBox2.Value = False
            Worksheets("Schneiden").CheckBox2.Visible = True
            Worksheets("Schneiden").CheckBox3.Value = False
            Worksheets("Schneiden").CheckBox3.Visible = True
            Worksheets("Schneiden").CheckBox4.Value = True 'Umschlag
            Worksheets("Schneiden").CheckBox4.Visible = True
        Else
            Worksheets("SEingabe").Range("B30") = "0"
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
        If Worksheets("Eingabe").CheckBox5.Value = True Then
            Worksheets("SEingabe").Range("B31") = "1"
        Else: Worksheets("SEingabe").Range("B31") = "0"
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox6_Click()
Attribute CheckBox6_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Kolbus
    Application.ScreenUpdating = False
        If Worksheets("Eingabe").CheckBox6.Value = True Then
            Worksheets("SEingabe").Range("B35") = "1"
        Else: Worksheets("SEingabe").Range("B35") = "0"
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox7_Click()
Attribute CheckBox7_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Verpacken
    Application.ScreenUpdating = False
        If Worksheets("Eingabe").CheckBox7.Value = True Then
            Worksheets("SEingabe").Range("B39") = "1"
        Else: Worksheets("SEingabe").Range("B39") = "0"
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox8_Click()
' Auswahl Buchdecke
    Application.ScreenUpdating = False
        If Worksheets("Eingabe").CheckBox8.Value = True Then
            Worksheets("SEingabe").Range("B36") = "1"
            Worksheets("Buchdecke").Visible = True
        Else
            Worksheets("SEingabe").Range("B36") = "0"
            Worksheets("Buchdecke").Visible = False
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox9_Click()
' Auswahl Kleben
    Application.ScreenUpdating = False
        If Worksheets("Eingabe").CheckBox9.Value = True Then
            Worksheets("SEingabe").Range("B32") = "1"
        Else: Worksheets("SEingabe").Range("B32") = "0"
        End If
    Run Worksheets("SKleben").kleben
    Application.ScreenUpdating = True
End Sub
Sub CheckBox10_Click()
' Auswahl Fadenheften
    Application.ScreenUpdating = False
        If Worksheets("Eingabe").CheckBox10.Value = True Then
            Worksheets("SEingabe").Range("B34") = "1"
            Worksheets("Fadenheften").Visible = True
        Else
            Worksheets("SEingabe").Range("B34") = "0"
            Worksheets("Fadenheften").Visible = False
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox11_Click()
'Auswahl Zusammentragen
'20140516
    If Worksheets("Eingabe").CheckBox11.Value = True Then
        Worksheets("SEingabe").Range("B33") = "1"
    Else: Worksheets("SEingabe").Range("B33") = "0"
    End If
End Sub
Sub CheckBox12_Click()
    'Auswahl Einhängen
    Application.ScreenUpdating = False
        If Worksheets("Eingabe").CheckBox12.Value = True Then
            Worksheets("SEingabe").Range("B37") = "1"
            Worksheets("Einhängen").Visible = True
        Else
            Worksheets("SEingabe").Range("B37") = "0"
            Worksheets("Einhängen").Visible = False
        End If
    Application.ScreenUpdating = True
End Sub
Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Produktbreite
    Select Case KeyAscii
      Case 48 To 57, 43 To 46
      Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub TextBox4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Tabsprung zur nächsten Eingabe
    If KeyCode = vbKeyTab Then
        Worksheets("Eingabe").TextBox5.Activate
        With Worksheets("Eingabe").TextBox5
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Sub
Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Produktlänge
    Select Case KeyAscii
      Case 48 To 57, 43 To 46
      Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Bezug Nutzenanzahl
    Select Case KeyAscii
      Case 48 To 57
      Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub TextBox6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
   If KeyCode = vbKeyTab Then
        Worksheets("Eingabe").TextBox4.Activate
        With Worksheets("Eingabe").TextBox4
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Sub
Sub Test()
    Run Worksheets("Binden").Produktart
End Sub
Sub Buchblockformat()
'Auswahl Buchblockformat
'20150220
    If Worksheets("SEingabe").Range("F27") = 7 Then 'Sonderformat
        Worksheets("Eingabe").TextBox4.Visible = True
        Worksheets("Eingabe").TextBox5.Visible = True
        With Worksheets("Eingabe").Range("E9")
            .Interior.ColorIndex = 6
            .Value = "x"
        End With
        Worksheets("Eingabe").Range("F9") = "Sonderformat"
        Worksheets("SEingabe").Range("E127") = 1
    Else
        Worksheets("Eingabe").ComboBox3.ListIndex = 6 'Standardformate nur mit 2mm Pappstärke
        Worksheets("Eingabe").Label3.Visible = True
        Worksheets("Eingabe").TextBox4.Visible = False
        Worksheets("Eingabe").TextBox5.Visible = False
        With Worksheets("Eingabe").Range("E9")
            .Interior.ColorIndex = 0
            .Value = ""
        End With
        Worksheets("Eingabe").Range("E9").Value = Worksheets("SEingabe").Range("G26").Value
        Worksheets("Eingabe").Range("F9") = ""
        Worksheets("SEingabe").Range("E127") = 0
    End If
End Sub


