Attribute VB_Name = "Modul1"
Sub Produkt()
Attribute Produkt.VB_ProcData.VB_Invoke_Func = " \n14"
' Anzeigen d. Produktangaben
'
On Error Resume Next
Dim format, Gewicht, Dicke As String
    format = Worksheets("Eingabe").Range("E9")
    Dicke = Range("Eingabe!C48")
    Gewicht = Range("Eingabe!C49")
    Worksheets("Verpacken").Label1.Caption = "Produkt:" & vbLf & "======" & vbLf & vbLf & "Format: " & vbLf & format _
    & vbLf & vbLf & "Stärke: " & vbLf & Dicke & " mm" & vbLf & vbLf & "Gewicht: " & vbLf & Gewicht & " g"
End Sub

