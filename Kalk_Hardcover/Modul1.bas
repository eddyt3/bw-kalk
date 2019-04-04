Attribute VB_Name = "Modul1"
Public Function Aufrunden(Zahl) As Long
If IsNumeric(Zahl) Then
If Zahl - Int(Zahl) < 0.5 Then
Aufrunden = Int(Zahl)
Else
Aufrunden = Int(Zahl) + 1
End If
End If
End Function
Public Function Abrunden(Zahl) As Long
If IsNumeric(Zahl) Then
If Zahl - Int(Zahl) < 0.5 Then
Abrunden = Int(Zahl)
Else
Abrunden = Int(Zahl) + 1
End If
End If
End Function
Public Function Interpolation(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X0 As Double) As Variant
Dim Y0 As Double
If X1 <> X2 Then
Y0 = (Y2 - Y1) / (X2 - X1) * (X0 - X1) + Y1
Interpolation = Y0
Else
End If
End Function
Public Function Interpolation2(ZeitBereich As Range, WertBereich As Range, t As Double) As Double
n = WertBereich.Cells.Count
ReDim Werte(1 To n)
For i = 1 To n
Werte(i) = WertBereich(i)
Next i
For i = 1 To n
For j = n To i + 1 Step -1
Werte(j) = (Werte(j) - Werte(j - 1)) / (ZeitBereich(j) - ZeitBereich(j - i))
Next j
Next i
Interpolation2 = 0
For i = n To 1 Step -1
Interpolation2 = Interpolation2 * (t - ZeitBereich(i)) + Werte(i)
Next i
End Function
Sub produkt()
Attribute produkt.VB_ProcData.VB_Invoke_Func = " \n14"
Dim format, Gewicht, Dicke As String
format = Worksheets("SEingabe").Range("G26")
Dicke = Round(Worksheets("SEingabe").Range("D127"), 1)
Gewicht = Worksheets("SEingabe").Range("B123")
Worksheets("Verpacken").Label1.Caption = "Produkt:" & vbLf & "======" & vbLf & vbLf & "Format: " & vbLf & format & vbLf & vbLf & "Stärke: " & vbLf & Dicke & "cm" & vbLf & vbLf & "Gewicht: " & vbLf & Gewicht & "g"
End Sub
