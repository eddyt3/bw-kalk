VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Public VNFormatN, VNFormatB As String 'Vor-, Nachsatzformat f. Kleben
Private Sub Worksheet_Activate()
    Run Worksheets("SKleben").Bogen_erstellen
    Call Worksheets("Kleben").CheckBox1_Click
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    Call Worksheets("Kleben").VNFormat
    Call Worksheets("Kleben").TextBox1_Change
    Run Worksheets("SKleben").Bogen_erstellen
End Sub
Private Sub CommandButton1_Click()
    'Materialseite einblenden
    Worksheets("Material").Visible = xlSheetVisible
    Worksheets("Material").Activate
    Worksheets("Material").Range("G348").Select
End Sub
Private Sub CommandButton21_Click()
    'Reset Schnittanzahl auf Standard
    '20150225
    Worksheets("Kleben").Range("D12").FormulaLocal = "=SKleben!D76"
End Sub
Sub CheckBox1_Click()
    ' Auswahl Vor-, Nachsatzkleben
        If Worksheets("Kleben").CheckBox1.Value = True Then
            Worksheets("SKleben").Range("I2") = "1"
            'Worksheets("Kleben").Rows("4:37").Hidden = False
            'Worksheets("Kleben").ComboBox1.Visible = True
            'Worksheets("Kleben").ComboBox2.Visible = True
            Worksheets("Kleben").ComboBox3.Visible = True 'Falzen
            Worksheets("Kleben").ComboBox3.Top = 234
            Worksheets("Kleben").TextBox1.Visible = True
            'Worksheets("Kleben").CommandButton21.Visible = True
            Call Worksheets("Kleben").ComboBox1_DropButtonClick
            Call Worksheets("Kleben").Maschinendaten
        Else
            Worksheets("SKleben").Range("I2") = "0"
            'Worksheets("Kleben").Rows("4:37").Hidden = True
            'Worksheets("Kleben").ComboBox1.Visible = False
            'Worksheets("Kleben").ComboBox2.Visible = False
            'Worksheets("Kleben").ComboBox3.Visible = False 'Falzen
            'Worksheets("Kleben").ComboBox4.Visible = False
            Worksheets("Kleben").TextBox1.Visible = False
            'Worksheets("Kleben").CommandButton21.Visible = False
        End If
        Run Worksheets("SKleben").KlebenAnzeige
End Sub
Sub CheckBox2_Click()
Attribute CheckBox2_Click.VB_ProcData.VB_Invoke_Func = " \n14"
    ' Auswahl Bogen ankleben
        If Worksheets("Kleben").CheckBox2.Value = True Then
            Worksheets("SKleben").Range("I3") = "1"
            'Worksheets("Kleben").Rows("40:51").Hidden = False
            Else
            Worksheets("SKleben").Range("I3") = "0"
            'Worksheets("Kleben").Rows("40:51").Hidden = True
        End If
        Run Worksheets("SKleben").KlebenAnzeige
End Sub
Private Sub CheckBox3_Click()
    Run Worksheets("SKleben").Bogen_erstellen
End Sub
Sub ComboBox1_DropButtonClick()
    'Papierberechnung je nach Auswahl
    Dim shpShape As Shape
    Dim intI As Integer
        If Worksheets("SKleben").Range("C33") = "1" Then 'Bubi eingekauft
            Worksheets("Kleben").Rows("4:6").Hidden = False
            Worksheets("Kleben").Rows("8:12").Hidden = False
            Worksheets("Kleben").Rows("7").Hidden = True
            Worksheets("Kleben").TextBox1.Visible = True 'Formatanzeige
            Worksheets("Kleben").TextBox8.Visible = False 'Nutzen gelieferte Bg.
            Worksheets("Kleben").ComboBox4.Visible = True 'Materialauswahl
            Worksheets("Kleben").ComboBox1.Top = 67.5 'Auswahl Bubi eingekauft od. Kunde geliefert
            Worksheets("Kleben").ComboBox4.Top = 66.75
            'Worksheets("Kleben").ComboBox4.ListIndex = 0
            Run Worksheets("Kleben").Materialauswahl
        Else 'Kunde geliefert
            Worksheets("Kleben").Rows("4:6").Hidden = True
            Worksheets("Kleben").Rows("8:12").Hidden = True
            Worksheets("Kleben").Rows("7").Hidden = False
            'Worksheets("Kleben").Range("D11:G11") = "0"
            Worksheets("Kleben").TextBox1.Visible = False
            For intI = 3 To 7
            With OLEObjects("TextBox" & intI)
                .Visible = False
            End With
            Next
            Worksheets("Kleben").TextBox8.Visible = True 'Nutzen gelieferte Bg.
            Worksheets("Kleben").TextBox8.Top = 53.25
            Worksheets("Kleben").ComboBox1.Top = 54
            Worksheets("Kleben").CheckBox2.Top = 408.75 'Bg. ankleben
            Worksheets("Kleben").ComboBox3.Top = 151.5
            Worksheets("Kleben").ComboBox4.Visible = False
            For Each shpShape In Worksheets("Kleben").Shapes
                If shpShape.Type = msoAutoShape Then
                    shpShape.Delete
                End If
            Next shpShape
            'Worksheets("Kleben").Shapes("Textfeld 23").Visible = False
        End If
End Sub
Private Sub ComboBox4_Change()
    'Materialauswahl Vor-/Nachsatz
    Run Worksheets("Kleben").Materialauswahl
    Run Worksheets("SKleben").Bogen_erstellen
End Sub
Sub TextBox1_Change()
    'Anzeige Bogeninfos
    If Worksheets("SKleben").Range("C33") = 1 Then
            With Worksheets("Kleben").TextBox1
                .Visible = True
                .Top = 58
                .Left = 744
                .Font.Name = "Calibri"
                .Font.Size = 10
                .Font.Bold = True
                .BorderColor = vbRed
                .BackColor = &HE0E0E0
                .Height = 72
                .Width = 200
                .WordWrap = False
                .Value = "Vor-, Nachsatzformat:" & vbCrLf & "================" & vbCrLf _
                & "Mit Zuschlag: " & VNFormatB & vbCrLf & vbCrLf & "Netto..........: " & VNFormatN
                '.Locked = True
            End With
    End If
End Sub
Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Vor-/Nachsatz alternatives Bogenformat Breite
    Select Case KeyAscii
      Case 48 To 57
      Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub TextBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Tabsprung zur nächsten Eingabe
    If KeyCode = vbKeyTab Then
        Worksheets("Kleben").TextBox4.Activate
        With Worksheets("Kleben").TextBox4
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Sub
Private Sub TextBox3_LostFocus()
    Run Worksheets("SKleben").Bogen_erstellen
End Sub
Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Vor-/Nachsatz alternatives Bogenformat Länge
    Select Case KeyAscii
      Case 48 To 57
      Case Else: KeyAscii = 0
    End Select
    Run Worksheets("SKleben").Bogen_erstellen
End Sub
Private Sub TextBox4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Tabsprung zur nächsten Eingabe
    If KeyCode = vbKeyTab Then
        Worksheets("Kleben").TextBox5.Activate
        With Worksheets("Kleben").TextBox5
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Sub
Private Sub TextBox4_LostFocus()
    Run Worksheets("SKleben").Bogen_erstellen
End Sub
Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Vor-/Nachsatz alternatives Bogenformat Grammatur
    Select Case KeyAscii
      Case 48 To 57
      Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub TextBox5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Tabsprung zur nächsten Eingabe
    If KeyCode = vbKeyTab Then
        Worksheets("Kleben").TextBox6.Activate
        With Worksheets("Kleben").TextBox6
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Sub
Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Vor-/Nachsatz alternatives Bogenformat Preis
    Select Case KeyAscii
      Case 48 To 57, 43 To 46
      Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub TextBox6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Tabsprung zur nächsten Eingabe
    If KeyCode = vbKeyTab Then
        Worksheets("Kleben").Range("D10").Activate
    End If
End Sub
Private Sub TextBox7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Vor-/Nachsatz alternatives Bogenformat Volumenfaktor
    Select Case KeyAscii
      Case 48 To 57, 43 To 46
      Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub TextBox7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Tabsprung zur nächsten Eingabe
    If KeyCode = vbKeyTab Then
        Worksheets("Kleben").TextBox3.Activate
        With Worksheets("Kleben").TextBox3
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Sub
Private Sub TextBox8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Nutzenanzahl gelieferte Bogen
    Select Case KeyAscii
      Case 48 To 57
      Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub TextBox8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Tabsprung zur nächsten Eingabe
    If KeyCode = vbKeyTab Then
        Worksheets("Kleben").Range("D7").Activate
    End If
End Sub
Sub VNFormat()
    'Berechnung Vor-, Nachsatzformat für Nutzenauswertung (Papierbestellung)
    VNFormatN = Worksheets("SKleben").Range("B52") & " cm x " & Worksheets("SKleben").Range("C52") & " cm"
    VNFormatB = Worksheets("SKleben").Range("B53") & " cm x " & Worksheets("SKleben").Range("C53") & " cm"
End Sub
Sub Maschinendaten()
    'MinMax Format Prüfen
    If Worksheets("SKleben").Range("H10") = True Then
        MsgBox ("Achtung das Produktformat über-/unterschreitet die technischen Maschinendaten!")
    End If
End Sub
Sub Materialauswahl()
    'Materialauswahl für eingekaufte Bögen
    Dim intI As Integer
        If Worksheets("SKleben").Range("B36") = "11" Then 'freie Eingabe
        Worksheets("Kleben").Rows("6").Hidden = False 'alternativer Bogen, Eingaben aktivieren
        For intI = 3 To 6
        With OLEObjects("TextBox" & intI)
            .Visible = True
            .Top = 85.5
        End With
        Next
        With OLEObjects("TextBox7")
            .Visible = True
            .Top = 65.25
        End With
    Else
        Worksheets("Kleben").Rows("6").Hidden = True
        For intI = 3 To 6
         With OLEObjects("TextBox" & intI)
            .Visible = False
            .Top = 85.5
        End With
        Next
        With OLEObjects("TextBox7")
            .Visible = False
            .Top = 65.25
        End With
    End If
End Sub

