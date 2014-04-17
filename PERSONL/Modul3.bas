Attribute VB_Name = "Modul3"
Public Sub letzte_zeile_1()
'Hier wird die letzte Zeile ermittelt
'Egal in welcher Spalte sich die letzte Zeile befindet
'Es werden alle Spalten geprüft und die letzte Zeile ausgegeben
letztezeile = Sheets(1).UsedRange.SpecialCells(xlCellTypeLastCell).Row
Msgbox letztezeile
End Sub
Public Sub letzte_zeile_2()
'Hier wir die letzte Zeile der Spalte A ermittelt
letztezeile = ActiveSheet.Cells(65536, 1).End(xlUp).Row
Msgbox letztezeile
End Sub
Public Sub letzte_zeile_3()
'Hier wir die letzte Zeile der Spalte A ermittelt
letztezeile = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row
Msgbox letztezeile
End Sub
Public Sub letzte_spalte_1()
'Hier wird die letzte Zeile ermittelt
'Egal in welcher Spalte sich die letzte Zeile befindet
'Es werden alle Spalten geprüft und die letzte Zeile ausgegeben
letztespalte = Sheets(1).UsedRange.SpecialCells(xlCellTypeLastCell).Column
Msgbox letztespalte
End Sub
Public Sub letzte_spalte_2()
'Hier wird die letzte spalte der Zeile 4 ermittelt
letztespalte = Sheets(1).Cells(4, 256).End(xlToLeft).Column
Msgbox letztespalte
End Sub
Public Sub letzte_zelle_1()
'Mit diesem Makro wird die Adresse der letzten Zelle (Zeile, Spalte) ermittelt
letztezelle = Range("A1").SpecialCells(xlCellTypeLastCell).Address
Msgbox letztezelle
End Sub
Public Sub letzte_zelle_2()
'Mit diesem Makro wird die letzte Zelle markiert
Range("A1").SpecialCells(xlCellTypeLastCell).Select
End Sub
