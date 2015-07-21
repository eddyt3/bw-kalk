VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Sub CheckBox1_Click()
Attribute CheckBox1_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Circut
    If Worksheets("3Seitenbeschnitt").CheckBox1.Value = True Then
        'Worksheets("3Seitenbeschnitt").CheckBox2.Value = False
        Worksheets("Steuerung").Range("E2") = "1"
        Else
        'Worksheets("3Seitenbeschnitt").CheckBox2.Value = True
        Worksheets("Steuerung").Range("E2") = "0"
    End If
End Sub
Sub CheckBox2_Click()
Attribute CheckBox2_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Circut
    If Worksheets("3Seitenbeschnitt").CheckBox2.Value = True Then
        'Worksheets("3Seitenbeschnitt").CheckBox1.Value = False
        Worksheets("Steuerung").Range("E3") = "1"
        Else
        'Worksheets("3Seitenbeschnitt").CheckBox1.Value = True
        Worksheets("Steuerung").Range("E3") = "0"
    End If
End Sub
