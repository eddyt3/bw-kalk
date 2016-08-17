Attribute VB_Name = "Modul3"
Public Sub letzte_zeile_1()
'Hier wird die letzte Zeile ermittelt
'Egal in welcher Spalte sich die letzte Zeile befindet
'Es werden alle Spalten geprüft und die letzte Zeile ausgegeben
letztezeile = Sheets(1).UsedRange.SpecialCells(xlCellTypeLastCell).Row
MsgBox letztezeile
End Sub
Public Sub letzte_zeile_2()
'Hier wir die letzte Zeile der Spalte A ermittelt
letztezeile = ActiveSheet.Cells(65536, 1).End(xlUp).Row
MsgBox letztezeile
End Sub
Public Sub letzte_zeile_3()
'Hier wir die letzte Zeile der Spalte A ermittelt
letztezeile = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row
MsgBox letztezeile
End Sub
Public Sub letzte_spalte_1()
'Hier wird die letzte Zeile ermittelt
'Egal in welcher Spalte sich die letzte Zeile befindet
'Es werden alle Spalten geprüft und die letzte Zeile ausgegeben
letztespalte = Sheets(1).UsedRange.SpecialCells(xlCellTypeLastCell).Column
MsgBox letztespalte
End Sub
Public Sub letzte_spalte_2()
'Hier wird die letzte spalte der Zeile 4 ermittelt
letztespalte = Sheets(1).Cells(4, 256).End(xlToLeft).Column
MsgBox letztespalte
End Sub
Public Sub letzte_zelle_1()
'Mit diesem Makro wird die Adresse der letzten Zelle (Zeile, Spalte) ermittelt
letztezelle = Range("A1").SpecialCells(xlCellTypeLastCell).Address
MsgBox letztezelle
End Sub
Public Sub letzte_zelle_2()
'Mit diesem Makro wird die letzte Zelle markiert
Range("A1").SpecialCells(xlCellTypeLastCell).Select
End Sub
Sub WertUndPosAusArrayBestimmen()
    'Ermittelt den kleinsten Wert und dessen Position aus einem ARRAY
    '04.11.2008, NoNet

    Dim arrWerte, lngZ As Long, intS As Integer, varMin
    arrWerte = [A1:C10] 'Werte aus Tabelle in ARRAY einlesen
    varMin = Application.Min(arrWerte)
    MsgBox varMin, , "Kleinster Wert der Matrix" 'Kleinsten Wert ausgeben
    'Position per MATRIX-Funktion in Tabelle ermitteln :
    '=ADRESSE(MAX(WENN(A1:C10=MIN(A1:C10);ZEILE(1:10);0));MAX(WENN(A1:C10=MIN(A1:C10 );SPALTE(A:C));0))

    'Variante 1 : Position per Schleifen ermitteln :
    For lngZ = LBound(arrWerte) To UBound(arrWerte)
        For intS = LBound(arrWerte) To UBound(Application.Transpose(arrWerte)) 'nur bis max 4561 Zeilen möglich !
            If arrWerte(lngZ, intS) = varMin Then
                MsgBox "Zeile : " & lngZ & vbLf & "Spalte : " & intS, , "Kleinster Wert an Position - Variante 1"
            End If
        Next
    Next

    'Variante 2 : Position durch Vergleich ermitteln
    For intS = LBound(arrWerte) To UBound(Application.Transpose(arrWerte)) 'nur bis max 4561 Zeilen möglich !

        If Not IsError(Application.Lookup(varMin, Application.WorksheetFunction.Index(Application.Transpose(arrWerte), intS))) Then
            lngZ = Application.Match(varMin, Application.WorksheetFunction.Index(Application.Transpose(arrWerte), intS))
            If lngZ > 0 Then
                MsgBox "Zeile : " & lngZ & vbLf & "Spalte : " & intS, , "Kleinster Wert an Position - Variante 2"
            End If
        End If
    Next
End Sub
