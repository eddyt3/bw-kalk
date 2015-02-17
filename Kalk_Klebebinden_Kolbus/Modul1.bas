Attribute VB_Name = "Modul1"
Public FFormat, FFormatMin As String 'Fehlervariabeln
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
        If (X0 < X1 And X0 < X2) Or (X0 > X1 And X0 > X2) Then
            MsgBox "X0 liegt nicht zwischen X1 und X2", vbInformation, "Trendberechnung"
        End If
    Else
        MsgBox "Tabellenwerte X1 und X2 dürfen nicht übereinstimmen", vbCritical, "Achtung"
        Interpolation = "#Fehler!"
    End If
End Function
Public Function Interpolation2(ZeitBereich As Range, WertBereich As Range, t As Double) As Double
    ' Die Funktion untersucht die übergebenen Arrays und ermittelt
    ' aus n Elementen von [Zeit] und [Werte] ein Polynom n-1 Grades,
    ' setzt [t] ein und gibt den interpolierten Wert zurück    Dim i As Long, j As Long, Werte() As Double, n As Long
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
Sub FormatMin()
Attribute FormatMin.VB_ProcData.VB_Invoke_Func = " \n14"
    'Prüfung Mindestformat für Zusammentragen
    If Worksheets("Steuerung").Range("N128") = 1 Then
        MsgBox ("Achtung das Mindestformat f. Das Zusammentragen wurde unterschritten!" & vbCrLf & vbCrLf & "(Mindestformat: " _
        & Worksheets("Zusammentragen").Range("K2") & " x " & Worksheets("Zusammentragen").Range("M2") & " cm)")
        FFormatMin = "Das Mindestformat für das Zusammentragen wurde unterschritten!"
        Else
        FFormatMin = ""
    End If
End Sub
Sub produkt()
Attribute produkt.VB_ProcData.VB_Invoke_Func = " \n14"
    ' Anzeigen d. Produktangaben
    On Error Resume Next
    Dim format, Gewicht, Dicke As String
        format = Worksheets("Eingabe").CommandButton3.Caption
        Dicke = Range("Eingabe!C45")
        Gewicht = Range("Eingabe!C46")
        Worksheets("Verpacken").Label1.Caption = "Produkt:" & vbLf & "======" & vbLf & vbLf & "Format: " & vbLf & format _
        & vbLf & vbLf & "Stärke: " & vbLf & Dicke & " mm" & vbLf & vbLf & "Gewicht: " & vbLf & Gewicht & " g"
End Sub
