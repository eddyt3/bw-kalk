VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CheckBox3_Click()
    Application.ScreenUpdating = False
    If Worksheets("Binden").CheckBox3.Value = True Then
        Worksheets("Zusammentragen").Visible = True
    Else
        Worksheets("Zusammentragen").Visible = False
    End If
    Application.ScreenUpdating = True
End Sub
Private Sub ComboBox3_DropButtonClick()
    Call Worksheets("Binden").Stoff
End Sub
Private Sub ComboBox5_Change()
    Call Worksheets("Binden").Leimart
End Sub
Private Sub CommandButton1_Click()
    Worksheets("Binden").Range("D8").FormulaLocal = "=Basis_Kolbus!H20"
    Worksheets("Binden").Range("F8").FormulaLocal = "=Basis_Kolbus!I20"
    Worksheets("Binden").Range("G8").FormulaLocal = "=Basis_Kolbus!J20"
End Sub
Sub Leimart()
    If Worksheets("Binden").ComboBox5.ListIndex = 0 Then
        Worksheets("Binden").Range("D3") = Worksheets("Binden").ComboBox5.ListIndex + 1
        Worksheets("Binden").ComboBox1.ListFillRange = "Material!D9:D13"
        Worksheets("Binden").ComboBox1.ListIndex = 0
        Worksheets("Binden").ComboBox1.LinkedCell = "Material!B8"
        Worksheets("Binden").ComboBox4.ListFillRange = "Material!D38:D43"
        Worksheets("Binden").ComboBox4.ListIndex = 0
        Worksheets("Binden").ComboBox4.LinkedCell = "Material!B37"
    Else
        Worksheets("Binden").ComboBox1.ListFillRange = "Material!D14:D18"
        Worksheets("Binden").ComboBox1.ListIndex = 0
        Worksheets("Binden").ComboBox1.LinkedCell = "Material!B8"
        Worksheets("Binden").ComboBox4.ListFillRange = "Material!D38:D43"
        Worksheets("Binden").ComboBox4.ListIndex = 0
        Worksheets("Binden").ComboBox4.LinkedCell = "Material!B37"
    End If
    If Worksheets("Binden").ComboBox5.ListIndex = 1 Then
        Worksheets("Binden").Range("D3") = Worksheets("Binden").ComboBox5.ListIndex + 1
        Worksheets("Binden").ComboBox1.ListFillRange = "Material!D14:D18"
        Worksheets("Binden").ComboBox1.ListIndex = 0
        Worksheets("Binden").ComboBox1.LinkedCell = "Material!B8"
        Worksheets("Binden").ComboBox4.ListFillRange = "Material!D38:D43"
        Worksheets("Binden").ComboBox4.ListIndex = 5
        Worksheets("Binden").ComboBox4.LinkedCell = "Material!B37"
    Else
        Worksheets("Binden").ComboBox1.ListFillRange = "Material!D9:D13"
        Worksheets("Binden").ComboBox1.ListIndex = 0
        Worksheets("Binden").ComboBox1.LinkedCell = "Material!B8"
        Worksheets("Binden").ComboBox4.ListFillRange = "Material!D38:D43"
        Worksheets("Binden").ComboBox4.ListIndex = 5
        Worksheets("Binden").ComboBox4.LinkedCell = "Material!B37"
    End If
        Worksheets("Binden").ComboBox2.ListFillRange = "Material!D26:D36"
        Worksheets("Binden").ComboBox2.ListIndex = 0
        Worksheets("Binden").ComboBox2.LinkedCell = "Material!B25"
End Sub
Sub Produktart()
    Dim intArt As Integer
    intArt = Worksheets("SBinden").Range("B34")
    Select Case intArt
        Case 1
            Worksheets("Binden").ComboBox5.ListIndex = 0
            Worksheets("Binden").Range("D3") = Worksheets("Binden").ComboBox5.ListIndex + 1
            Worksheets("Binden").ComboBox1.ListFillRange = "Material!D9:D13"
            Worksheets("Binden").ComboBox1.LinkedCell = "Material!B8"
            Worksheets("Binden").ComboBox1.ListIndex = 0
            Worksheets("Binden").ComboBox2.ListFillRange = "Material!D26:D36"
            Worksheets("Binden").ComboBox2.ListIndex = 0
            Worksheets("Binden").ComboBox2.LinkedCell = "Material!B25"
            Worksheets("Binden").ComboBox3.ListFillRange = "Material!D52:D55"
            Worksheets("Binden").ComboBox3.LinkedCell = "Material!B51"
            Worksheets("Binden").ComboBox3.ListIndex = 0
            Worksheets("Binden").ComboBox4.ListFillRange = "Material!D38:D43"
            Worksheets("Binden").ComboBox4.LinkedCell = "Material!B37"
            Worksheets("Binden").ComboBox4.ListIndex = 5
        Case 2
            Worksheets("Binden").ComboBox5.ListIndex = 0
            Worksheets("Binden").Range("D3") = Worksheets("Binden").ComboBox5.ListIndex + 1
            Worksheets("Binden").ComboBox1.ListFillRange = "Material!D9:D13"
            Worksheets("Binden").ComboBox1.LinkedCell = "Material!B8"
            Worksheets("Binden").ComboBox1.ListIndex = 0
            Worksheets("Binden").ComboBox2.ListFillRange = "Material!D26:D36"
            Worksheets("Binden").ComboBox2.ListIndex = 0
            Worksheets("Binden").ComboBox2.LinkedCell = "Material!B25"
            Worksheets("Binden").ComboBox3.ListFillRange = "Material!D52:D55"
            Worksheets("Binden").ComboBox3.LinkedCell = "Material!B51"
            Worksheets("Binden").ComboBox3.ListIndex = 3
            Worksheets("Binden").ComboBox4.ListFillRange = "Material!D38:D43"
            Worksheets("Binden").ComboBox4.LinkedCell = "Material!B37"
            Worksheets("Binden").ComboBox4.ListIndex = 5
        Case 3
            Worksheets("Binden").ComboBox5.ListIndex = 1
            Worksheets("Binden").Range("D3") = Worksheets("Binden").ComboBox5.ListIndex + 1
            Worksheets("Binden").ComboBox1.ListFillRange = "Material!D14:D18"
            Worksheets("Binden").ComboBox1.LinkedCell = "Material!B8"
            Worksheets("Binden").ComboBox1.ListIndex = 0
            Worksheets("Binden").ComboBox2.ListFillRange = "Material!D26:D36"
            Worksheets("Binden").ComboBox2.ListIndex = 0
            Worksheets("Binden").ComboBox2.LinkedCell = "Material!B25"
            Worksheets("Binden").ComboBox3.ListFillRange = "Material!D52:D55"
            Worksheets("Binden").ComboBox3.LinkedCell = "Material!B51"
            Worksheets("Binden").ComboBox3.ListIndex = 3
            Worksheets("Binden").ComboBox4.ListFillRange = "Material!D38:D43"
            Worksheets("Binden").ComboBox4.LinkedCell = "Material!B37"
            Worksheets("Binden").ComboBox4.ListIndex = 5
        Case 4
            Worksheets("Binden").ComboBox5.ListIndex = 1
            Worksheets("Binden").Range("D3") = Worksheets("Binden").ComboBox5.ListIndex + 1
            Worksheets("Binden").ComboBox1.ListFillRange = "Material!D14:D18"
            Worksheets("Binden").ComboBox1.LinkedCell = "Material!B8"
            Worksheets("Binden").ComboBox1.ListIndex = 0
            Worksheets("Binden").ComboBox2.ListFillRange = "Material!D26:D36"
            Worksheets("Binden").ComboBox2.ListIndex = 10
            Worksheets("Binden").ComboBox2.LinkedCell = "Material!B25"
            Worksheets("Binden").ComboBox3.ListFillRange = "Material!D52:D55"
            Worksheets("Binden").ComboBox3.LinkedCell = "Material!B51"
            Worksheets("Binden").ComboBox3.ListIndex = 0
            Worksheets("Binden").ComboBox4.ListFillRange = "Material!D38:D43"
            Worksheets("Binden").ComboBox4.LinkedCell = "Material!B37"
            Worksheets("Binden").ComboBox4.ListIndex = 5
    End Select
    Call Stoff
End Sub
Sub Stoff()
    If Worksheets("Material").Range("C51") > 1 Then
        Worksheets("Binden").Rows("20:28").Hidden = False
        Worksheets("Binden").Rows("27").Hidden = True
    Else
        Worksheets("Binden").Rows("20:28").Hidden = True
    End If
End Sub
