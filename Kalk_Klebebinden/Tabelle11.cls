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
    If Worksheets("Binden").CheckBox1.Value = True Then
        Worksheets("Binden").CheckBox2.Value = False
        Worksheets("Steuerung").Range("K2") = "1"
        Worksheets("Binden").ComboBox1.ListFillRange = "Material!C9:C10" 'Auswahl PUR Kleber
        Worksheets("Binden").ComboBox1.ListIndex = 0 'Ersten Wert d. Auswahl autom. anzeigen
        Worksheets("Binden").ComboBox1.LinkedCell = "Steuerung!K4" 'Ausgabe Leimname
    Else: Worksheets("Binden").CheckBox2.Value = True
        Worksheets("Steuerung").Range("K2") = "0"
        Worksheets("Binden").ComboBox1.ListFillRange = "Material!C15" 'Auswahl HM Kleber
        Worksheets("Binden").ComboBox1.ListIndex = 0
        Worksheets("Binden").ComboBox1.LinkedCell = "Steuerung!K4" 'Ausgabe Leimname
    End If
Worksheets("Binden").ComboBox2.ListFillRange = "Material!C16:C19" 'Auswahl Seitenleim
Worksheets("Binden").ComboBox2.ListIndex = 0
Worksheets("Binden").ComboBox2.LinkedCell = "Steuerung!K5"
End Sub
Sub CheckBox2_Click()
Attribute CheckBox2_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Leimart Hotmelt
    If Worksheets("Binden").CheckBox2.Value = True Then
        Worksheets("Binden").CheckBox1.Value = False
        Worksheets("Steuerung").Range("K3") = "1"
        Worksheets("Binden").ComboBox1.ListFillRange = "Material!C15"
        Worksheets("Binden").ComboBox1.ListIndex = 0
        Worksheets("Binden").ComboBox1.LinkedCell = "Steuerung!K4"
    Else: Worksheets("Binden").CheckBox1.Value = True
        Worksheets("Steuerung").Range("K3") = "0"
        Worksheets("Binden").ComboBox1.ListFillRange = "Material!C9:C10"
        Worksheets("Binden").ComboBox1.ListIndex = 0
        Worksheets("Binden").ComboBox1.LinkedCell = "Steuerung!K4"
    End If
Worksheets("Binden").ComboBox2.ListFillRange = "Material!C16:C19"
Worksheets("Binden").ComboBox2.ListIndex = 0
Worksheets("Binden").ComboBox2.LinkedCell = "Steuerung!K5"
End Sub
Sub CheckBox3_Click()
Attribute CheckBox3_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Quickbinder
    If Worksheets("Binden").CheckBox3.Value = True Then
        Worksheets("Binden").CheckBox4.Value = False
        Worksheets("Steuerung").Range("C34") = "1"
    Else: Worksheets("Binden").CheckBox4.Value = True
        Worksheets("Steuerung").Range("C34") = "0"
    End If
End Sub
Sub CheckBox4_Click()
Attribute CheckBox4_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Handbinden
    If Worksheets("Binden").CheckBox4.Value = True Then
        Worksheets("Binden").CheckBox3.Value = False
        Worksheets("Steuerung").Range("C35") = "1"
    Else: Worksheets("Binden").CheckBox3.Value = True
        Worksheets("Steuerung").Range("C35") = "0"
    End If
End Sub
Sub ComboBox3_DropButtonClick()
' Auswahl Scharnierstoff od. Faelzel
With Worksheets("Binden").ComboBox3
    .ListFillRange = "Steuerung!B30:B32"
    .LinkedCell = "Steuerung!E31"
End With
If Worksheets("Steuerung").Range("C30") = "0" Then
    If Worksheets("Steuerung").Range("C31") = "1" Then
        Worksheets("Binden").Rows("19:26").Hidden = False
        Worksheets("Binden").Rows("27").Hidden = True
        Else
        Worksheets("Binden").Range("D27:G27") = ""
        Worksheets("Binden").Rows("19:26").Hidden = True
        Worksheets("Binden").Rows("27").Hidden = False
    End If
Else
    Worksheets("Binden").Rows("19:26").Hidden = True
End If
End Sub
