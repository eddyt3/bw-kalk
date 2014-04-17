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
    'Also z.B. wie folgt: =SummeWennFarbe(A1:A10;A1)+(0*JETZT())
    Dim intColor        As Integer
    Dim lngI            As Long
    Dim Summe           As Variant
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

