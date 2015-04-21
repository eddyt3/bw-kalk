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
    On Error Resume Next
    Run Worksheets("SKleben").Bogen_erstellen_netto
    Call CheckBox1_Click
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    Call VNFormat
    Call TextBox1_Change
    Run Worksheets("SKleben").Produktformat_Kleben
    Run Worksheets("SKleben").Bogen_erstellen_netto
End Sub
Private Sub CommandButton21_Click()
    'Reset Schnittanzahl auf Standard
    '20150225
    Worksheets("Kleben").Range("D12").FormulaLocal = "=SKleben!D76"
End Sub
Sub CheckBox1_Click()
Attribute CheckBox1_Click.VB_ProcData.VB_Invoke_Func = " \n14"
    ' Auswahl Vor-, Nachsatzkleben
    On Error Resume Next
        If CheckBox1.Value = True Then
            Worksheets("SKleben").Range("I2") = "1"
            Worksheets("Kleben").Rows("4:36").Hidden = False
            ComboBox1.Visible = True
            ComboBox2.Visible = True
            ComboBox3.Visible = True
            TextBox1.Visible = True
            CheckBox3.Visible = True
            Call ComboBox1_DropButtonClick
            Call Maschinendaten
            Else
            Worksheets("SKleben").Range("I2") = "0"
            Worksheets("Kleben").Rows("4:36").Hidden = True
            ComboBox1.Visible = False
            ComboBox2.Visible = False
            ComboBox3.Visible = False
            TextBox1.Visible = False
            CheckBox3.Visible = False
        End If
        Run Worksheets("SKleben").KlebenAnzeige
End Sub
Sub CheckBox2_Click()
Attribute CheckBox2_Click.VB_ProcData.VB_Invoke_Func = " \n14"
    ' Auswahl Bogen ankleben
        If CheckBox2.Value = True Then
            Worksheets("SKleben").Range("I3") = "1"
            Worksheets("Kleben").Rows("39:50").Hidden = False
            Else
            Worksheets("SKleben").Range("I3") = "0"
            Worksheets("Kleben").Rows("39:50").Hidden = True
        End If
        Run Worksheets("SKleben").KlebenAnzeige
End Sub
Private Sub CheckBox3_Click()
    On Error Resume Next
    Run Worksheets("SKleben").Bogen_erstellen_netto
End Sub
Private Sub ComboBox1_DropButtonClick()
    'Papierberechnung je nach Auswahl
    Dim shpShape As Shape
        If Worksheets("SKleben").Range("C33") = "1" Then 'Bubi eingekauft
            Worksheets("Kleben").Rows("4:5").Hidden = False
            Worksheets("Kleben").Rows("7:11").Hidden = False
            Worksheets("Kleben").Rows("6").Hidden = True
            Worksheets("Kleben").Range("D6:G6") = "0"
            ThisWorkbook.Sheets("Kleben").TextBox8.Visible = False 'Nutzen gelieferte Bg.
            TextBox1.Visible = True
            CheckBox3.Visible = True
            Else 'Kunde geliefert
            Worksheets("Kleben").Rows("4:5").Hidden = True
            Worksheets("Kleben").Rows("7:11").Hidden = True
            Worksheets("Kleben").Rows("6").Hidden = False
            Worksheets("Kleben").Range("D10:G10") = "0"
            ThisWorkbook.Sheets("Kleben").TextBox8.Visible = True 'Nutzen gelieferte Bg.
            Worksheets("Kleben").TextBox8.Top = 53.25
            TextBox1.Visible = False
            CheckBox3.Visible = False
            For Each shpShape In Worksheets("Kleben").Shapes
                If shpShape.Type = msoAutoShape Then
                    shpShape.Delete
                End If
            Next shpShape
        End If
End Sub
Private Sub TextBox1_Change()
    'Anzeige Bogeninfos
    On Error Resume Next
    If Worksheets("SKleben").Range("C33") = 1 Then
            With TextBox1
                .Visible = True
                .Top = 58
                .Left = 744
                .Font = "Verdana"
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
Private Sub TextBox8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Nutzenanzahl gelieferte Bogen
    Select Case KeyAscii
      Case 48 To 57
      Case Else: KeyAscii = 0
    End Select
End Sub
Sub VNFormat()
    '
    'Berechnung Vor-, Nachsatzformat für Nutzenauswertung (Papierbestellung)
    '
    On Error Resume Next
    VNFormatN = Worksheets("SKleben").Range("B52") & " cm x " & Worksheets("SKleben").Range("C52") & " cm"
    VNFormatB = Worksheets("SKleben").Range("B53") & " cm x " & Worksheets("SKleben").Range("C53") & " cm"
End Sub
Sub Maschinendaten()
    'MinMax Format Prüfen
    If Worksheets("SKleben").Range("H10") = True Then
        MsgBox ("Achtung das Produktformat über-/unterschreitet die technischen Maschinendaten!")
    End If
End Sub
