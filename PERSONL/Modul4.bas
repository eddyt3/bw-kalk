Attribute VB_Name = "Modul4"
'Modul für alle Funktionen
'
Public Function Farbcode(rngFarb As Range) As Byte
    Application.Volatile
    If rngFarb.Interior.ColorIndex > 0 Then
        Farbcode = rngFarb.Interior.ColorIndex
    Else: Farbcode = 0
    End If
End Function
Public Function SummeWennFarbe(Bereich As Range, _
                               SuchFarbe As Variant, _
                               Optional Summe_Bereich As Range, _
                               Optional bolFont As Boolean = False) _
       As Variant
    '© t.ramel@mvps.org, 30.05.2003
    'erweitert 01.07.2004, 31.08.2004, 11.12.2004, 18.04.2005
    'Funktion zur Anwendung von SUMMEWENN() mit Hintergrund- oder Schriftfarbe
    'als Kriterium
    '
    'Die Parametereingabe erfolgt in derselben Reihenfolge
    'wie in der Funktion SUMMEWENN():
    ' - Der erste Parameter erwartet den Suchbereich
    ' - Der zweite Parameter erwartet einen Zellbezug
    '   (Hintergrund/Schriftfarbe) oder einen Farbindex (Zahl)
    '   Farbindex '0' zählt Zellen ohne Hintergrund/Standard-Schriftfarbe
    ' - Der dritte Parameter erwartet optional den zu summierenden Bereich
    ' - Der vierte Parameter erwartet Wahr/Falsch für die Festlegung
    '   ob nach Hintergrund- oder Schriftfarbe summiert werden soll
    
    'Zur automatischen Aktualisierung im Tabellenblatt den folgenden Term
    'anhängen: +(0*JETZT()) und durch F9 drücken die Funktion aktualisieren
    'Also z.B. wie folgt:
    Dim intColor As Integer
    Dim lngI As Long
    Dim Summe As Variant
   If Summe_Bereich Is Nothing Then Set Summe_Bereich = Bereich
   If bolFont Then
      If IsObject(SuchFarbe) Then
         intColor = SuchFarbe(1).Font.ColorIndex
      Else
         intColor = SuchFarbe
      End If
      For lngI = 1 To Bereich.count
         If Bereich(lngI).Font.ColorIndex = intColor Then
            Summe = Summe + CDec(Summe_Bereich(lngI))
         End If
      Next
   Else
      If IsObject(SuchFarbe) Then
         intColor = SuchFarbe(1).Interior.ColorIndex
      Else
         intColor = SuchFarbe
      End If

      For lngI = 1 To Bereich.count
         If Bereich(lngI).Interior.ColorIndex = intColor Then
            Summe = Summe + CDec(Summe_Bereich(lngI))
         End If
      Next lngI
   End If
   SummeWennFarbe = Summe
End Function
Public Function AnzahlWennFarbe(Bereich As Range, _
                               SuchFarbe As Variant, _
                               Optional Anzahl_Bereich As Range, _
                               Optional bolFont As Boolean = False) _
       As Variant
    Dim intColor, Anzahl As Integer
    Dim lngI            As Long
    Anzahl = 0
   If Anzahl_Bereich Is Nothing Then Set Anzahl_Bereich = Bereich
   If bolFont Then
      If IsObject(SuchFarbe) Then
         intColor = SuchFarbe(1).Font.ColorIndex
      Else
         intColor = SuchFarbe
      End If
      For lngI = 1 To Bereich.count
         If Bereich(lngI).Font.ColorIndex = intColor Then
            Anzahl = Anzahl + 1
         End If
      Next
   Else
      If IsObject(SuchFarbe) Then
         intColor = SuchFarbe(1).Interior.ColorIndex
      Else
         intColor = SuchFarbe
      End If

      For lngI = 1 To Bereich.count
         If Bereich(lngI).Interior.ColorIndex = intColor Then
            Anzahl = Anzahl + 1
         End If
      Next lngI
   End If
   AnzahlWennFarbe = Anzahl
End Function
Public Function SuchenUndErsetzen(QuellText, Suchen, Optional Ersetzen)
    'Die Funktion sucht in einem String eine Zeichenkette beliebiger Länge und ersetzt sie durch eine beliebige andere Zeichenkette.
    'Wenn die zu ersetzende Zeichenfolge fehlt, wird der Text aus der Zeichenkette entfernt.
     Dim Pos As Long    ' Position im bearbeiteten String
     Dim LängeSuchText As Long, LängeErsatzText As Long
     ' Fehlerprüfung
     If (Nz(QuellText) = vbNullString) Then GoTo Ende
     If (Nz(Suchen) = vbNullString) Then GoTo Ende
     If IsMissing(Ersetzen) Or IsNull(Ersetzen) Then Ersetzen = vbNullString
     LängeSuchText = Len(Suchen)
     LängeErsatzText = Len(Ersetzen)
     Pos = InStr(1, QuellText, Suchen)
     While Pos <> 0
         QuellText = Left(QuellText, Pos - 1) & Ersetzen & _
                     Mid(QuellText, Pos + LängeSuchText)
         Pos = InStr(Pos + LängeErsatzText, QuellText, Suchen)
     Wend
Ende:
     SuchenUndErsetzen = QuellText
End Function
Function LIP(xVector As Range, yVector As Range, xValue As Double)
    Dim Dimension As Long, MinDim As Long, MaxDim As Long
    Dim I_oben As Long, I_unten As Long, i As Long
    Dimension = xVector.Cells.count
    On Error GoTo Fehler
    '1. X-Y-Wertepaar bestimmen, das verschieden ist von Leerstring
    For i = 1 To Dimension
      If xVector(i) <> "" And yVector(i) <> "" Then MinDim = i: Exit For
    Next
    'letztes X-Y-Wertepaar bestimmen, das verschieden ist von Leerstring
    For i = Dimension To 1 Step -1
      If xVector(i) <> "" And yVector(i) <> "" Then MaxDim = i: Exit For
    Next
    If xValue < xVector.Cells(MinDim).Value Or xValue > xVector.Cells(MaxDim).Value Then
    'Extrapolation der Werte
        If xValue < xVector.Cells(MinDim).Value Then
            'Nächstes X-Y-Wertepaar mit Werten verschieden von Leerstring
            For i = MinDim + 1 To Dimension
              If xVector(i) <> "" And yVector(i) <> "" Then Exit For
            Next
            m = (yVector.Cells(i) - yVector.Cells(MinDim)) / (xVector.Cells(i) - xVector.Cells(MinDim))
            n = yVector.Cells(i) - m * xVector.Cells(i)
        Else
            'Vorletztes X-Y-Wertepaar mit Werten verschieden von Leerstring
            For i = MaxDim - 1 To MinDim Step -1
              If xVector(i) <> "" And yVector(i) <> "" Then Exit For
            Next
            m = (yVector.Cells(MaxDim) - yVector.Cells(i)) / (xVector.Cells(MaxDim) - xVector.Cells(i))
            n = yVector.Cells(MaxDim) - m * xVector.Cells(MaxDim)
        End If
        LIP = m * xValue + n
    
    Else
    'Interpolation der Werte
        'X-Y-Wertepaar mit X-Wert >= gesuchten X-Wert
        For i = MinDim + 1 To MaxDim
          If xValue <= xVector.Cells(i).Value And yVector(i) <> "" Then I_oben = i: Exit For
        Next i
        'Vorheriges X-Y-Wertepaar mit Werten verschieden von Leerstring
        For i = I_oben - 1 To MinDim Step -1 '###### Korrketur in dieser Zeile ####
          If xVector(i) <> "" And yVector(i) <> "" Then I_unten = i: Exit For
        Next

        LIP = yVector.Cells(I_unten).Value _
          + (xValue - xVector.Cells(I_unten).Value) / _
          (xVector.Cells(I_oben).Value - xVector.Cells(I_unten).Value) _
          * (yVector.Cells(I_oben).Value - yVector.Cells(I_unten).Value)

    End If
    Exit Function
Fehler:
    LIP = "Interpolationsfehler"
End Function
Sub Objektliste_Wkb_erstellen()
    'alle OLE Objekte auflisten
    Dim shListe As Worksheet
    Dim sh As Worksheet
    Dim obj As OLEObject
    Dim shp As Shape
    Dim sp As Long, ze As Long
    On Error GoTo Objektliste_erstellen
    Set shListe = ActiveWorkbook.Sheets("Objektliste")
    On Error GoTo 0
    shListe.Cells.Clear
    sp = 1: ze = 2
    For Each sh In ActiveWorkbook.Worksheets
        shListe.Cells(1, sp).Value = sh.Name
        For Each obj In sh.OLEObjects
            shListe.Cells(ze, sp).Value = obj.Name
            ze = ze + 1
        Next
        For Each shp In sh.Shapes
            shListe.Cells(ze, sp).Value = shp.Name
            ze = ze + 1
        Next
        sp = sp + 1
        ze = 2
    Next
    Exit Sub
Objektliste_erstellen:
    Sheets.Add before:=ActiveWorkbook.Sheets(1)
    ActiveSheet.Name = "Objektliste"
    Resume
End Sub
Sub Objektliste_Wks_Direktbereich_erstellen()
    'alle OLE Objekte nur im Direktbereich auflisten
    Dim obj As OLEObject
    Dim shp As Shape
    For Each obj In ActiveSheet.OLEObjects
        Debug.Print obj.Name
    Next
    For Each shp In ActiveSheet.Shapes
        Debug.Print shp.Name
    Next
    Exit Sub
End Sub

