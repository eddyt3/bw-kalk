Attribute VB_Name = "Modul1"
Sub Produkt()
Attribute Produkt.VB_ProcData.VB_Invoke_Func = " \n14"
' Anzeigen d. Produktangaben
Dim format, Gewicht, Dicke As String
    format = Worksheets("Eingabe").Range("E9")
    Dicke = Worksheets("Eingabe").Range("C48")
    Gewicht = Worksheets("Eingabe").Range("C49")
    Worksheets("Verpacken").Label1.Caption = "Produkt:" & vbLf & "======" & vbLf & vbLf & "Format: " & vbLf & format _
    & vbLf & vbLf & "Stärke: " & vbLf & Dicke & " mm" & vbLf & vbLf & "Gewicht: " & vbLf & Gewicht & " g"
End Sub

