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
'        If (X0 < X1 And X0 < X2) Or (X0 > X1 And X0 > X2) Then
'            MsgBox "X0 liegt nicht zwischen X1 und X2", vbInformation, "Trendberechnung"
'        End If
    Else
'        MsgBox "Tabellenwerte X1 und X2 d�rfen nicht �bereinstimmen", vbCritical, "Achtung"
'        Interpolation = "#Fehler!"
    End If
End Function
Public Function Interpolation2(ZeitBereich As Range, WertBereich As Range, t As Double) As Double
    ' Die Funktion untersucht die �bergebenen Arrays und ermittelt
    ' aus n Elementen von [Zeit] und [Werte] ein Polynom n-1 Grades,
    ' setzt [t] ein und gibt den interpolierten Wert zur�ck    Dim i As Long, j As Long, Werte() As Double, n As Long
    n = WertBereich.Cells.Count
    ReDim Werte(1 To n)
    For i = 1 To n
        Werte(i) = WertBereich(i)
    Next i
    ' Interpolation nach Newton
    For i = 1 To n
        For j = n To i + 1 Step -1
            Werte(j) = (Werte(j) - Werte(j - 1)) / (ZeitBereich(j) - ZeitBereich(j - i))
        Next j
    Next i
    ' Hornerschema anwenden, um Interpolationspolynom auszuwerten
    Interpolation2 = 0
    For i = n To 1 Step -1
        Interpolation2 = Interpolation2 * (t - ZeitBereich(i)) + Werte(i)
    Next i
End Function
Sub produkt()
Attribute produkt.VB_ProcData.VB_Invoke_Func = " \n14"
    ' Anzeigen d. Produktangaben
    Dim format, Gewicht, Dicke As String
        format = Worksheets("SEingabe").Range("B127") & "cm x " & Worksheets("SEingabe").Range("C127") & "cm"
        Dicke = Round(Worksheets("SEingabe").Range("D127"), 1)
        Gewicht = Worksheets("SEingabe").Range("B123")
        Worksheets("Verpacken").Label1.Caption = "Produkt:" & vbLf & "======" & vbLf & vbLf & "Format: " & vbLf & format _
        & vbLf & vbLf & "St�rke: " & vbLf & Dicke & "cm" & vbLf & vbLf & "Gewicht: " & vbLf & Gewicht & "g"
End Sub
