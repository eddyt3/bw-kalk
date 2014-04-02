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
Private Sub Worksheet_Activate()

End Sub
Sub CheckBox1_Click()
Attribute CheckBox1_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Leimart PUR
On Error Resume Next
    If CheckBox1.Value = True Then
        Worksheets("Binden").CheckBox2.Value = False
        Worksheets("Steuerung").Range("K2") = "1"
        ComboBox1.ListFillRange = "Material!C9:C10" 'Auswahl PUR Kleber
        ComboBox1.ListIndex = 0 'Ersten Wert d. Auswahl autom. anzeigen
        ComboBox1.LinkedCell = "Steuerung!K4" 'Ausgabe Leimname
    Else: Worksheets("Binden").CheckBox2.Value = True
        Worksheets("Steuerung").Range("K2") = "0"
        ComboBox1.ListFillRange = "Material!C15" 'Auswahl HM Kleber
        ComboBox1.ListIndex = 0
        ComboBox1.LinkedCell = "Steuerung!K4" 'Ausgabe Leimname
    End If
ComboBox2.ListFillRange = "Material!C16:C19" 'Auswahl Seitenleim
ComboBox2.ListIndex = 0
ComboBox2.LinkedCell = "Steuerung!K5"
End Sub
Sub CheckBox2_Click()
Attribute CheckBox2_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Leimart Hotmelt
On Error Resume Next
    If CheckBox2.Value = True Then
        Worksheets("Binden").CheckBox1.Value = False
        Worksheets("Steuerung").Range("K3") = "1"
        ComboBox1.ListFillRange = "Material!C15"
        ComboBox1.ListIndex = 0
        ComboBox1.LinkedCell = "Steuerung!K4"
    Else: Worksheets("Binden").CheckBox1.Value = True
        Worksheets("Steuerung").Range("K3") = "0"
        ComboBox1.ListFillRange = "Material!C9:C10"
        ComboBox1.ListIndex = 0
        ComboBox1.LinkedCell = "Steuerung!K4"
    End If
ComboBox2.ListFillRange = "Material!C16:C19"
ComboBox2.ListIndex = 0
ComboBox2.LinkedCell = "Steuerung!K5"
End Sub
Sub CheckBox3_Click()
Attribute CheckBox3_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Quickbinder
On Error Resume Next
    If CheckBox3.Value = True Then
        Worksheets("Binden").CheckBox4.Value = False
        Worksheets("Steuerung").Range("C34") = "1"
    Else: Worksheets("Binden").CheckBox4.Value = True
        Worksheets("Steuerung").Range("C34") = "0"
    End If
End Sub
Sub CheckBox4_Click()
Attribute CheckBox4_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Handbinden
On Error Resume Next
    If CheckBox4.Value = True Then
        Worksheets("Binden").CheckBox3.Value = False
        Worksheets("Steuerung").Range("C35") = "1"
    Else: Worksheets("Binden").CheckBox3.Value = True
        Worksheets("Steuerung").Range("C35") = "0"
    End If
End Sub
Private Sub ComboBox3_DropButtonClick()
' Auswahl Scharnierstoff od. Faelzel
On Error Resume Next
With ComboBox3
    .ListFillRange = "Steuerung!B30:B32"
    .LinkedCell = "Steuerung!E31"
End With
If Worksheets("Steuerung").Range("C30") = "0" Then
    If Worksheets("Steuerung").Range("C31") = "1" Then
        Worksheets("Binden").Unprotect "bw"
        Worksheets("Binden").Rows("19:26").Hidden = False
        Worksheets("Binden").Rows("27").Hidden = True
        Else
        Worksheets("Binden").Unprotect "bw"
        Worksheets("Binden").Range("D27:G27") = ""
        Worksheets("Binden").Rows("19:26").Hidden = True
        Worksheets("Binden").Rows("27").Hidden = False
    End If
Else
    Worksheets("Binden").Unprotect "bw"
    Worksheets("Binden").Rows("19:26").Hidden = True
End If
Worksheets("Binden").Protect "bw"
End Sub
