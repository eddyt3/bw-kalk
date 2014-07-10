Attribute VB_Name = "Modul1"
'all rights by E.Dargel ed@dissenter.de
Public FNutzen, FFormat, FDicke As String 'Fehlervariabeln
Public FBindenS, FBindenB, FBindenG As String 'Fehlervariabeln Binden
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
            Aufrunden = Int(Zahl) + 1
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
Sub Nutzenauswertung()
    '
    ' Nutzenauswertung Makro
    ' Makro von Enrico Dargel
    ' Überprüfung der Angaben Seitenzahl, Bogenzahl, Seiten pro Bogen
    On Error Resume Next
    If Range("Steuerung!H57") > 0 Then
        Dim BogenA, BogenB, BogenC, BogenD As String
            If Range("Steuerung!H53") > 0 Then
                MsgBox ("Fehlerhafte Eingabe(n) bei Bogen A! " & vbCrLf & vbCrLf & _
                "Bitte 'Seitenzahl', 'Nutzen/Druckbogen', 'Buchbindebogen' u. 'Seiten/Buchbindebogen' kontrollieren.")
                BogenA = " A,"
                Else
                BogenA = ""
            End If
            If Range("Steuerung!H54") > 0 Then
                MsgBox ("Fehlerhafte Eingabe(n) bei Bogen B! " & vbCrLf & vbCrLf & _
                "Bitte 'Seitenzahl', 'Nutzen/Druckbogen', 'Buchbindebogen' u. 'Seiten/Buchbindebogen' kontrollieren.")
                BogenB = " B,"
                Else
                BogenB = ""
            End If
            If Range("Steuerung!H55") > 0 Then
                MsgBox ("Fehlerhafte Eingabe(n) bei Bogen C! " & vbCrLf & vbCrLf & _
                "Bitte 'Seitenzahl', 'Nutzen/Druckbogen', 'Buchbindebogen' u. 'Seiten/Buchbindebogen' kontrollieren.")
                BogenC = " C,"
                Else
                BogenC = ""
            End If
            If Range("Steuerung!H56") > 0 Then
                MsgBox ("Fehlerhafte Eingabe(n) bei Bogen D! " & vbCrLf & vbCrLf & _
                "Bitte 'Seitenzahl', 'Nutzen/Druckbogen', 'Buchbindebogen' u. 'Seiten/Buchbindebogen' kontrollieren.")
                BogenD = " D"
                Else
                BogenD = ""
            End If
        FNutzen = "Fehlerhafte Seitenzahl, Bogenzahl od. Seiten/Bogen bei Bogen: " & BogenA & BogenB & BogenC & BogenD & "."
        Else
        FNutzen = ""
    End If
End Sub
Sub produkt()
Attribute produkt.VB_ProcData.VB_Invoke_Func = " \n14"
    '
    ' Anzeigen d. Produktangaben
    ' Makro von Enrico Dargel
    '
    Dim format, Gewicht, Dicke As String
        format = Worksheets("Eingabe").CommandButton3.Caption
        Dicke = Range("Eingabe!C44")
        Gewicht = Range("Eingabe!C45")
        Worksheets("Verpacken").Label1.Caption = "Produkt:" & vbLf & "======" & vbLf & vbLf & "Format: " & vbLf & format _
        & vbLf & vbLf & "Stärke: " & vbLf & Dicke & " mm" & vbLf & vbLf & "Gewicht: " & vbLf & Gewicht & " g"
End Sub
Sub Evaluierung_Binden() 'in Arbeit
Attribute Evaluierung_Binden.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error Resume Next
    Dim bgAS, bgBS, bgCS, bgDS As Integer 'Seiten pro Bogen
    Dim bgAB, bgBB, bgCB, bgDB As Integer 'Bogenanzahl
    Dim bgAN, bgBN, bgCN, bgDN As Integer 'Nutzen
    Dim bgAG, bgBG, bgCG, bgDG As Integer 'Grammatur
    Dim FbgAS, FbgBS, FbgCS, FbgDS, FbgAB, FbgBB, FbgCB, FbgDB, FbgAN, FbgBN, FbgCN, FbgDN, FbgAG, FbgBG, FbgCG, FbgDG As String
    bgAS = Worksheets("Steuerung").Range("D61")
    bgAB = Worksheets("Steuerung").Range("D62")
    bgAN = Worksheets("Steuerung").Range("D63")
    bgAG = Worksheets("Steuerung").Range("D64")
    bgBS = Worksheets("Steuerung").Range("E61")
    bgBB = Worksheets("Steuerung").Range("E62")
    bgBN = Worksheets("Steuerung").Range("E63")
    bgBG = Worksheets("Steuerung").Range("E64")
    bgCS = Worksheets("Steuerung").Range("F61")
    bgCB = Worksheets("Steuerung").Range("F62")
    bgCN = Worksheets("Steuerung").Range("F63")
    bgCG = Worksheets("Steuerung").Range("F64")
    bgDS = Worksheets("Steuerung").Range("G61")
    bgDB = Worksheets("Steuerung").Range("G62")
    bgDN = Worksheets("Steuerung").Range("G63")
    bgDG = Worksheets("Steuerung").Range("G64")
    'Kontrolle Seitenanzahl pro Bogen
    If bgAS < 2 Then
        If bgAS = 1 Then FbgAS = " A," Else: FbgAS = ""
        Else: FbgAS = ""
    End If
    If bgBS < 2 Then
        If bgBS = 1 Then FbgBS = " B," Else: FbgBS = ""
        Else: FbgBS = ""
    End If
    If bgCS < 2 Then
        If bgCS = 1 Then FbgCS = " C," Else: FbgCS = ""
        Else: FbgCS = ""
    End If
    If bgDS < 2 Then
        If bgDS = 1 Then FbgDS = " D," Else: FbgDS = ""
        Else: FbgDS = ""
    End If
    If Worksheets("Steuerung").Range("H61") > 0 Then
        MsgBox ("Fehlerhafte Seitenzahl pro Bogen:" & FbgAS & FbgBS & FbgCS & FbgDS & " (min. 8 u. max. 24 Seiten/Bg.)")
        FBindenS = "Fehlerhafte Seitenzahl pro Bogen:" & FbgAS & FbgBS & FbgCS & FbgDS & " (min. 8 u. max. 24 Seiten/Bg.)"
        Else: FBindenS = ""
    End If
    'Kontrolle Bogenzahl
    If bgAB < 2 Then
        If bgAB = 1 Then FbgAB = " A," Else: FbgAB = ""
        Else: FbgAB = ""
    End If
    If bgBB < 2 Then
        If bgBB = 1 Then FbgBB = " B," Else: FbgBB = ""
        Else: FbgBB = ""
    End If
    If bgCB < 2 Then
        If bgCB = 1 Then FbgCB = " C," Else: FbgCB = ""
        Else: FbgCB = ""
    End If
    If bgDB < 2 Then
        If bgDB = 1 Then FbgDB = " D," Else: FbgDB = ""
        Else: FbgDB = ""
    End If
    If Worksheets("Steuerung").Range("H62") > 0 Then
        MsgBox ("Fehlerhafte Bogenzahl:" & FbgAB & FbgBB & FbgCB & FbgDB & " (min. 3 u. max. 256 Bögen)")
        FBindenB = "Fehlerhafte Bogenzahl:" & FbgAB & FbgBB & FbgCB & FbgDB & " (min. 3 u. max. 256 Bögen)"
        Else: FBindenB = ""
    End If
    'Kontrolle Nutzen
    '
    ' in Arbeit
    '
    'Kontrolle Grammatur
    If bgAG < 2 Then
        If bgAG = 1 Then FbgAG = " A," Else: FbgAG = ""
        Else: FbgAG = ""
    End If
    If bgBG < 2 Then
        If bgBBG = 1 Then FbgBG = " B," Else: FbgBG = ""
        Else: FbgBG = ""
    End If
    If bgCG < 2 Then
        If bgCG = 1 Then FbgCG = " C," Else: FbgCG = ""
        Else: FbgCG = ""
    End If
    If bgDG < 2 Then
        If bgDG = 1 Then FbgDG = " D," Else: FbgDG = ""
        Else: FbgDG = ""
    End If
    If Worksheets("Steuerung").Range("H64") > 0 Then
        MsgBox ("Fehlerhafte Grammatur Bogen:" & FbgAB & FbgBB & FbgCB & FbgDB & " (min. 100 g/qm u. max. 300 g/qm)")
        FBindenG = "Fehlerhafte Grammatur Bogen:" & FbgAB & FbgBB & FbgCB & FbgDB & " (min. 100 g/qm u. max. 300 g/qm)"
        Else: FBindenG = ""
    End If
    ' Ueberpruefung d. Mindeststärke
    Dicke = Worksheets("Steuerung").Range("D59")
    If Worksheets("Steuerung").Range("B59") < Worksheets("Steuerung").Range("C59") Then
        MsgBox ("Achtung zu geringe Produktstärke!" & vbCrLf & vbCrLf & "Die Mindeststärke beträgt 3 mm.")
        FDicke = "Das Produkt ist für das Binden " & Dicke & " mm zu dünn (Mindeststärke: 3 mm)."
        Else
        FDicke = ""
    End If
End Sub
