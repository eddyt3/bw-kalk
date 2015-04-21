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
' 20150220
On Error Resume Next
    If CheckBox1.Value = True Then
        Worksheets("3Seitenbeschnitt").CheckBox2.Value = False
        Worksheets("S3Seitenbeschnitt").Range("H4") = "1"
        Else
        Worksheets("3Seitenbeschnitt").CheckBox2.Value = True
        Worksheets("S3Seitenbeschnitt").Range("H4") = "0"
    End If
End Sub
Sub CheckBox2_Click()
Attribute CheckBox2_Click.VB_ProcData.VB_Invoke_Func = " \n14"
' Auswahl Circut
' 20150220
On Error Resume Next
    If CheckBox2.Value = True Then
        Worksheets("3Seitenbeschnitt").CheckBox1.Value = False
        Worksheets("S3Seitenbeschnitt").Range("H5") = "1"
        Else
        Worksheets("3Seitenbeschnitt").CheckBox1.Value = True
        Worksheets("S3Seitenbeschnitt").Range("H5") = "0"
    End If
End Sub
