VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Sub CheckBox1_Click()
Attribute CheckBox1_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Leimart PUR
    On Error Resume Next
        If CheckBox1.Value = True Then
            Worksheets("Binden").CheckBox2.Value = False
            Worksheets("SBinden").Range("D2") = "1"
            ComboBox1.ListFillRange = "Material!D9:D13" 'Auswahl PUR Kleber
            ComboBox1.ListIndex = 0 'Ersten Wert d. Auswahl autom. anzeigen
            ComboBox1.LinkedCell = "Material!B8" 'Ausgabe Leimname
            ComboBox4.ListFillRange = "Material!D38:D43" 'Auswahl PUR Reiniger
            ComboBox4.ListIndex = 0
            ComboBox4.LinkedCell = "Material!B37" 'Ausgabe PUR Reiniger
        Else: Worksheets("Binden").CheckBox2.Value = True
            Worksheets("SBinden").Range("D2") = "0"
            ComboBox1.ListFillRange = "Material!D14:D18" 'Auswahl HM Kleber
            ComboBox1.ListIndex = 0
            ComboBox1.LinkedCell = "Material!B8" 'Ausgabe Leimname
            ComboBox4.ListFillRange = "Material!D43" 'kein Reiniger
            ComboBox4.ListIndex = 0
            ComboBox4.LinkedCell = "Material!B37" 'kein Reiniger
        End If
    ComboBox2.ListFillRange = "Material!D26:D38" 'Auswahl Seitenleim
    ComboBox2.ListIndex = 0
    ComboBox2.LinkedCell = "Material!B25"
End Sub
Sub CheckBox2_Click()
Attribute CheckBox2_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Leimart Hotmelt
    On Error Resume Next
        If CheckBox2.Value = True Then
            Worksheets("Binden").CheckBox1.Value = False
            Worksheets("SBinden").Range("D3") = "1"
            ComboBox1.ListFillRange = "Material!D14:D18"
            ComboBox1.ListIndex = 0
            ComboBox1.LinkedCell = "Material!B8"
            ComboBox4.ListFillRange = "Material!D43" 'kein Reiniger
            ComboBox4.ListIndex = 0
            ComboBox4.LinkedCell = "Material!B37" 'kein Reiniger
        Else: Worksheets("Binden").CheckBox1.Value = True
            Worksheets("SBinden").Range("D3") = "0"
            ComboBox1.ListFillRange = "Material!D9:D13"
            ComboBox1.ListIndex = 0
            ComboBox1.LinkedCell = "Material!B8"
            ComboBox4.ListFillRange = "Material!D38:D43" 'Auswahl PUR Reiniger
            ComboBox4.ListIndex = 0
            ComboBox4.LinkedCell = "Material!B37" 'Ausgabe PUR Reiniger
        End If
    ComboBox2.ListFillRange = "Material!D26:D38"
    ComboBox2.ListIndex = 0
    ComboBox2.LinkedCell = "Material!B25"
End Sub
Private Sub CheckBox3_Click()
'Einblenden Blatt Zusammentragen
    On Error Resume Next
    Application.ScreenUpdating = False
        If CheckBox3.Value = True Then
            Worksheets("Zusammentragen").Visible = True
        Else
            Worksheets("Zusammentragen").Visible = False
        End If
    Application.ScreenUpdating = True
End Sub
Private Sub ComboBox3_DropButtonClick()
' Auswahl Scharnierstoff od. Fälzel
' 20130219
    On Error Resume Next
    With ComboBox3
        .ListFillRange = "Material!D52:D54"
        .LinkedCell = "Material!B51"
    End With
    If Worksheets("Material").Range("C51") > 1 Then
        If Worksheets("Material").Range("C51") = 2 Then 'Scharnierstoff
            Worksheets("Binden").Unprotect "bw"
            Worksheets("Binden").Rows("20:27").Hidden = False
            Worksheets("Binden").Rows("28").Hidden = True
            Else 'Fälzel
            Worksheets("Binden").Unprotect "bw"
            'Worksheets("Binden").Range("D28:G28") = ""
            Worksheets("Binden").Rows("20:28").Hidden = False
            Worksheets("Binden").Rows("27").Hidden = True
        End If
    Else
        Worksheets("Binden").Unprotect "bw"
        Worksheets("Binden").Rows("20:28").Hidden = True
    End If
    Worksheets("Binden").Protect "bw"
End Sub

Private Sub CommandButton1_Click()
    'Reset V auf Standard
    '20140325
    On Error Resume Next
    Worksheets("Binden").Range("D8").FormulaLocal = "=Basis_Kolbus!H17"
    Worksheets("Binden").Range("F8").FormulaLocal = "=Basis_Kolbus!I17"
    Worksheets("Binden").Range("G8").FormulaLocal = "=Basis_Kolbus!J17"
End Sub
