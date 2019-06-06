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
Sub Worksheet_Change(ByVal Target As Excel.Range)
    Application.ScreenUpdating = False
        Run Worksheets("SSchneiden").Schneidbogen
        Run Worksheets("SSchneiden").SchneidenAnzeige
        Run Worksheets("Zusammentragen").Zusammentragen
        'Run Worksheets("Falzen").Falzen
    Application.ScreenUpdating = True
End Sub
Sub ComboBox2_Change()
    'Bindeseite
    Run Worksheets("SKleben").kleben
End Sub
Sub ComboBox3_DropButtonClick()
    ' Umschlag Art
    ' 20181016
    Dim intUS As Integer
    intUS = Worksheets("SEingabe").Range("L7")
    Select Case intUS
        Case 1
        '4-seitiger Umschlag
            Worksheets("Binden").ComboBox3.ListIndex = 0 'ohne Fälzel
            Call Worksheets("Binden").Stoff
        Case 2, 3
        '6/8-seitiger Umschlag
            MsgBox ("Bitte das Falzen der Umschläge nicht vergessen.")
            Worksheets("Binden").ComboBox3.ListIndex = 0 'ohne Fälzel
            Call Worksheets("Binden").Stoff
        Case 4, 5
        'Otabind immer mit Fälzel
        Worksheets("Binden").ComboBox3.ListIndex = 2
        Call Worksheets("Binden").Stoff
    End Select
    Run Worksheets("Falzen").umschlag
End Sub
Private Sub ComboBox4_Change()
    'Bogen A, Art
    Run Worksheets("SSchneiden").Schneidbogen
    Run Worksheets("SSchneiden").SchneidenAnzeige
    'Run Worksheets("Falzen").Falzen
End Sub
Sub ComboBox5_DropButtonClick()
    'Zusammentragen Autom./Hand
    Run Worksheets("Zusammentragen").Zusammentragen
End Sub
Sub ComboBox9_Change()
    'Bogen B, Art
    Run Worksheets("SSchneiden").Schneidbogen
    Run Worksheets("SSchneiden").SchneidenAnzeige
    'Run Worksheets("Falzen").Falzen
    Run Worksheets("Falzen").umschlag
End Sub
Sub ComboBox10_Change()
    'Bogen C, Art
    Run Worksheets("SSchneiden").Schneidbogen
    Run Worksheets("SSchneiden").SchneidenAnzeige
    'Run Worksheets("Falzen").Falzen
    Run Worksheets("Falzen").umschlag
End Sub
Sub ComboBox11_Change()
    'Bogen D, Art
    Run Worksheets("SSchneiden").Schneidbogen
    Run Worksheets("SSchneiden").SchneidenAnzeige
    'Run Worksheets("Falzen").Falzen
    Run Worksheets("Falzen").umschlag
End Sub
Private Sub ComboBox12_Change()
    'Tiefdrucklagen
End Sub
Private Sub ComboBox15_Change()
'Auswahl Buchblockformat
'20150421
    Run Worksheets("Eingabe").Produktformat
End Sub
Sub CommandButton1_Click()
' Drucken
'Auswahl neuen Standard Drucker
    Application.Dialogs(xlDialogPrinterSetup).Show
    'MsgBox "Alter Drucker: " & x & vbNewLine & "Neuer Drucker: " & Application.ActivePrinter
    With ActiveSheet
        .PageSetup.BlackAndWhite = True
        .PrintOut Copies:=1, Collate:=True
        '.PageSetup.BlackAndWhite = False
        '.SelectedSheets.PrintOut Copies:=1, Collate:=False
    End With
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
            'Worksheets("Schneiden").CheckBox2.Visible = True
            If Worksheets("Eingabe").CheckBox5.Value = True And Worksheets("Eingabe").ComboBox9.ListIndex = 0 Then 'Falzen und Inhalt aktiv?
                Worksheets("Falzen").ComboBox2.ListIndex = 0 'Bg. B Falzen aktivieren
            End If
        Else
            Worksheets("Eingabe").Range("A18") = "0"
            Worksheets("Schneiden").CheckBox2.Value = False
            'Worksheets("Schneiden").CheckBox2.Visible = False
        End If
        Run Worksheets("SSchneiden").Schneidbogen
        Run Worksheets("SSchneiden").SchneidenAnzeige
        Run Worksheets("Falzen").umschlag
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
            'Worksheets("Schneiden").CheckBox3.Visible = True
            If Worksheets("Eingabe").CheckBox5.Value = True And Worksheets("Eingabe").ComboBox10.ListIndex = 0 Then 'Falzen und Inhalt aktiv?
                Worksheets("Falzen").ComboBox3.ListIndex = 0 'Bg. C Falzen aktivieren
            End If
        Else
            Worksheets("Eingabe").Range("A27") = "0"
            Worksheets("Schneiden").CheckBox3.Value = False
            'Worksheets("Schneiden").CheckBox3.Visible = False
        End If
        Run Worksheets("SSchneiden").Schneidbogen
        Run Worksheets("SSchneiden").SchneidenAnzeige
        Run Worksheets("Falzen").umschlag
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
            'Worksheets("Schneiden").CheckBox4.Visible = True
        Else
            Worksheets("Eingabe").Range("A36") = "0"
            Worksheets("Schneiden").CheckBox4.Value = False
            'Worksheets("Schneiden").CheckBox4.Visible = False
        End If
        Run Worksheets("SSchneiden").Schneidbogen
        Run Worksheets("SSchneiden").SchneidenAnzeige
        Run Worksheets("Falzen").umschlag
        'Run Worksheets("Falzen").Falzen
    Application.ScreenUpdating = True
End Sub
Sub CheckBox4_Click()
Attribute CheckBox4_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Schneiden
    Application.ScreenUpdating = False
        If Worksheets("Eingabe").CheckBox4.Value = True Then 'Schneiden
            Worksheets("SEingabe").Range("B30") = "1"
'            Worksheets("Schneiden").CheckBox1.Value = False 'Nur Umschlag soll standardmäßig geschnitte werden
'            Worksheets("Schneiden").CheckBox1.Visible = True
'            Worksheets("Schneiden").CheckBox2.Value = False
'            Worksheets("Schneiden").CheckBox2.Visible = True
'            Worksheets("Schneiden").CheckBox3.Value = False
'            Worksheets("Schneiden").CheckBox3.Visible = True
'            Worksheets("Schneiden").CheckBox4.Value = True 'Umschlag
'            Worksheets("Schneiden").CheckBox4.Visible = True
        Else
            Worksheets("SEingabe").Range("B30") = "0"
'            Worksheets("Schneiden").CheckBox1.Value = False
'            Worksheets("Schneiden").CheckBox1.Visible = False
'            Worksheets("Schneiden").CheckBox2.Value = False
'            Worksheets("Schneiden").CheckBox2.Visible = False
'            Worksheets("Schneiden").CheckBox3.Value = False
'            Worksheets("Schneiden").CheckBox3.Visible = False
'            Worksheets("Schneiden").CheckBox4.Value = False
'            Worksheets("Schneiden").CheckBox4.Visible = False
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
            Worksheets("SEingabe").Range("B36") = "1"
        Else: Worksheets("SEingabe").Range("B36") = "0"
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
' Auswahl Bohren
    Application.ScreenUpdating = False
        If Worksheets("Eingabe").CheckBox8.Value = True Then
            Worksheets("SEingabe").Range("B38") = "1"
            Worksheets("Bohren").Visible = True
        Else
            Worksheets("SEingabe").Range("B38") = "0"
            Worksheets("Bohren").Visible = False
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox9_Click()
' Auswahl Kleben
    Application.ScreenUpdating = False
        If Worksheets("Eingabe").CheckBox9.Value = True Then
            Worksheets("SEingabe").Range("B33") = "1"
        Else: Worksheets("SEingabe").Range("B33") = "0"
        End If
    Run Worksheets("SKleben").kleben
    Application.ScreenUpdating = True
End Sub
Sub CheckBox10_Click()
' Auswahl Fadenheften
    Application.ScreenUpdating = False
        If Worksheets("Eingabe").CheckBox10.Value = True Then
            Worksheets("SEingabe").Range("B35") = "1"
            Worksheets("Fadenheften").Visible = True
            Worksheets("Zusammentragen").ComboBox2.ListIndex = 4
        Else
            Worksheets("SEingabe").Range("B35") = "0"
            Worksheets("Fadenheften").Visible = False
            Worksheets("Zusammentragen").ComboBox2.ListIndex = 3
        End If
    Application.ScreenUpdating = True
End Sub
Sub CheckBox11_Click()
'Auswahl Zusammentragen
'20140516
    If Worksheets("Eingabe").CheckBox11.Value = True Then
        Worksheets("SEingabe").Range("B34") = "1"
    Else: Worksheets("SEingabe").Range("B34") = "0"
    End If
End Sub
Sub ComboBox13_Change()
'Auswahl Leimart über Eingabeseite
'20180531
    Run Worksheets("Binden").Leimart
    With Worksheets("Eingabe").ComboBox13
        .FontSize = 20
    End With
End Sub
Sub ComboBox14_Change()
'Auswahl Produktart (Broschüre, Buchblock, Ableimen)
'20170418
    Dim intArt As Integer
    intArt = ThisWorkbook.Worksheets("SBinden").Range("B34").Value
    Select Case intArt
    Case 1
        'Broschüre
        ThisWorkbook.Worksheets("Kleben").Visible = False
        Worksheets("Eingabe").CheckBox9.Value = False 'Kalk. Kleben deaktiviert
        Worksheets("SEingabe").Range("B33") = "0"
        Worksheets("Eingabe").CheckBox10.Value = False 'Kalk. Fadenheften deaktiviert
        Worksheets("SEingabe").Range("B35") = "0"
        Worksheets("Fadenheften").Visible = False
        Worksheets("SZusammentragen").Range("B43") = "1" 'Kolbus ZU autom. Inline
        Worksheets("SZusammentragen").Range("C11") = "4" 'Kolbus ZU autom. Inline
    Case 2
        'Buchblock
        ThisWorkbook.Worksheets("Kleben").Visible = True
        Worksheets("Eingabe").CheckBox9.Value = True 'Kalk. Kleben aktiviert
        Worksheets("SEingabe").Range("B33") = "1"
        Worksheets("Eingabe").CheckBox10.Value = False 'Kalk. Fadenheften deaktiviert
        Worksheets("SEingabe").Range("B35") = "0"
        Worksheets("Fadenheften").Visible = False
        Worksheets("SZusammentragen").Range("B43") = "1" 'Kolbus ZU autom. Inline
        Worksheets("SZusammentragen").Range("C11") = "4" 'Kolbus ZU autom. Inline
    Case 3
        'Buchblock fadengeheftet
        ThisWorkbook.Worksheets("Kleben").Visible = True
        Worksheets("Eingabe").CheckBox9.Value = False 'Kalk. Kleben deaktiviert
        Worksheets("SEingabe").Range("B33") = "1"
        Worksheets("Eingabe").CheckBox10.Value = True 'Kalk. Fadenheften aktiviert
        Worksheets("SEingabe").Range("B35") = "1"
        Worksheets("Fadenheften").Visible = True
        Worksheets("SZusammentragen").Range("B43") = "2" 'Kolbus ZU handanlage
        Worksheets("SZusammentragen").Range("C11") = "5" 'Zusammentragen auf Astronico
    Case 4
        'Ableimen
        ThisWorkbook.Worksheets("Kleben").Visible = False
        Worksheets("Eingabe").CheckBox9.Value = False 'Kalk. Kleben deaktiviert
        Worksheets("SEingabe").Range("B33") = "0"
        Worksheets("Eingabe").CheckBox10.Value = False 'Kalk. Fadenheften deaktiviert
        Worksheets("SEingabe").Range("B35") = "0"
        ThisWorkbook.Worksheets("Fadenheften").Visible = False
        Worksheets("SZusammentragen").Range("B43") = "1" 'Kolbus ZU autom. Inline
        Worksheets("SZusammentragen").Range("C11") = "4" 'Kolbus ZU autom. Inline
    End Select
    Worksheets("SZusammentragen").Calculate
    Application.Run Worksheets("Binden").Produktart
    Application.Run Worksheets("Zusammentragen").GeschwindigkeitZus
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
'20150221
    If Worksheets("SEingabe").Range("F27") = 8 Then 'Sonderformat
        Worksheets("Eingabe").TextBox1.Visible = True
        Worksheets("Eingabe").TextBox3.Visible = True
        With Worksheets("Eingabe").Range("E9")
            .Interior.ColorIndex = 6
            .Value = "x"
        End With
        Worksheets("Eingabe").Range("F9") = "Sonderformat"
    Else
        'Standardformate
        Worksheets("Eingabe").TextBox1.Visible = False
        Worksheets("Eingabe").TextBox3.Visible = False
        With Worksheets("Eingabe").Range("E9")
            .Interior.ColorIndex = 0
            .Value = ""
        End With
        Worksheets("Eingabe").Range("E9").Value = Worksheets("SEingabe").Range("G26").Value
        Worksheets("Eingabe").Range("F9") = ""
    End If
End Sub
Sub Test()
    Run Worksheets("Binden").Produktart
End Sub
