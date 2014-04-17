Attribute VB_Name = "Modul2"
Sub Verknuepfungen_Suchen()
Dim rngGefundeneZelle As Range
Dim intI%, intN%, intAbfrage%, intZ�hler%, intNamenAbfrage%, intL�schZ�hler%
Dim strSuchbegriff$, strAktuelleZelle$, strErsteZelle$
Dim objName As Object
strSuchbegriff = "]"
If strSuchbegriff = "" Then Exit Sub
ReDim strBereiche(0)
ReDim intBl�tter(0)
intZ�hler = 0
For intI = 1 To Worksheets.count
    If Sheets(intI).ProtectContents Then
        Msgbox "Das Blatt " & Sheets(intI).Name & " ist gesch�tzt." & Chr(10) & "Entfernen Sie bitte zuerst den Blattschutz.", vbOKOnly + vbInformation, "Blatt gesch�tzt!"
        GoTo n�chstesBlatt
    End If
    Set rngGefundeneZelle = Worksheets(intI).Cells.Find(strSuchbegriff, lookat:=xlPart, LookIn:=xlFormulas)
    If Not rngGefundeneZelle Is Nothing Then
        'erste gefundene Zelle auf dem aktuellen Blatt:
        strErsteZelle = rngGefundeneZelle.Address(False, False)
        intZ�hler = intZ�hler + 1
        ReDim Preserve intBl�tter(intZ�hler)
        ReDim Preserve strBereiche(intZ�hler)
        intBl�tter(intZ�hler - 1) = intI
        strBereiche(intZ�hler - 1) = strErsteZelle
        Do While strAktuelleZelle <> strErsteZelle
            'n�chste gefundene Zelle(n) auf dem aktuellen Blatt
            Set rngGefundeneZelle = Worksheets(intI).Cells.FindNext(after:=rngGefundeneZelle)
            strAktuelleZelle = rngGefundeneZelle.Address(False, False)
            If strErsteZelle <> strAktuelleZelle Then
                intZ�hler = intZ�hler + 1
                ReDim Preserve intBl�tter(intZ�hler)
                ReDim Preserve strBereiche(intZ�hler)
                intBl�tter(intZ�hler - 1) = intI
                strBereiche(intZ�hler - 1) = strAktuelleZelle
            End If
        Loop
        strAktuelleZelle = ""
        strErsteZelle = ""
    End If
n�chstesBlatt:
Next
Set rngGefundeneZelle = Nothing
intL�schZ�hler = 0
For intN = 0 To intZ�hler - 1
    Sheets(intBl�tter(intN)).Select
    Range(strBereiche(intN)).Select
    intAbfrage = Msgbox("Auf dem Blatt " & Sheets(intBl�tter(intN)).Name & " wurde in der Zelle " & strBereiche(intN) & " eine Verkn�pfung gefunden." & Chr(10) & "Die Formel lautet:" & Chr(10) & Chr(10) & ActiveCell.Formula & Chr(10) & Chr(10) & "Soll sie gel�scht werden?", vbYesNo + vbQuestion, "Verkn�pfung gefunden")
    If intAbfrage = vbYes Then
        Range(strBereiche(intN)).ClearContents
        intL�schZ�hler = intL�schZ�hler + 1
    End If
Next
For Each objName In ActiveWorkbook.Names
    If InStr(1, objName.Value, strSuchbegriff) > 1 Then
        intZ�hler = intZ�hler + 1
        intNamenAbfrage = Msgbox("In einem Namen besteht eine Verkn�pfung." & Chr(10) & "Bezieht sich auf: " & objName.Value & Chr(10) & "Name: " & objName.Name, vbYesNo + vbQuestion, "Soll der Name gel�scht werden?")
        If intNamenAbfrage = vbYes Then
            objName.Delete
            intL�schZ�hler = intL�schZ�hler + 1
        End If
    End If
Next
If intZ�hler = 0 Then
    Msgbox "Keine Verkn�pfung gefunden oder die Bl�tter sind gesch�tzt.", vbOKOnly + vbInformation, "Fertig!"
Else
    Msgbox "Es wurden insgesamt " & intZ�hler & " Verkn�pfung(en) gefunden und davon " & intL�schZ�hler & " gel�scht.", vbOKOnly + vbInformation, "Fertig!"
End If
End Sub
Sub Verknuepfungen_aendern()
    'Beispiel - Verkn�pfungen i. einer Arbeitsmappe �ndern
    Dim wbkMappe As Workbook
    Dim varVLink As Variant
    Dim i, e As Integer
    Dim strPrefix, strPath, strFile, strRefFile As String
    
    strPath = ThisWorkbook.Path
    strFile = ThisWorkbook.Name
    strPrefix = Left(strFile, 4)
    strRefFile = strPrefix & "_Verrechnung.xls" 'verkn�pfte Tabelle
    'Info-Ausgabe
    Worksheets("Steuerung").Range("A6") = strPath
    Worksheets("Steuerung").Range("A7") = strFile
    Worksheets("Steuerung").Range("A8") = strRefFile

    Set wbkMappe = ThisWorkbook
    varVLink = wbkMappe.LinkSources(xlExcelLinks)
    
    If Not IsEmpty(varVLink) Then
        For i = 1 To UBound(varVLink)
            e = InStrRev(varVLink(i), "\") + 1
            'strRefFile = Mid(varVLink(i), e, 20) 'alternativ Referenzfilenamen auslesen
            ThisWorkbook.ChangeLink varVLink(i), strPath & "\" & strRefFile, xlLinkTypeExcelLinks
        Next i
    End If
    Msgbox "Fertig Master!"
End Sub
Sub datum_splitten()
    'Datum in Tag, Monat und Jahr aufsplitten
    'Achtung neben Datum drei Leerspalten einf�gen
    Dim lZeile As Long
    Dim vDatArr As Variant
    Dim vSplArr() As Variant
    result = Msgbox("Achtung, neben der Datumsspalte m�ssen 3 Leerspalten vorhanden sein!", 1, "Hinweis")
    If result = 2 Then
        Exit Sub
    End If
    ReDim vSplArr(Range("D65536").End(xlUp).Row, 2) As Variant
    vDatArr = Range("D2:D" & Range("D65536").End(xlUp).Row)
    For lZeile = LBound(vDatArr) To UBound(vDatArr)
        vSplArr(lZeile, 0) = Day(vDatArr(lZeile, 1))
        vSplArr(lZeile, 1) = Month(vDatArr(lZeile, 1))
        vSplArr(lZeile, 2) = Year(vDatArr(lZeile, 1))
    Next lZeile
    Range("E1:G" & Range("D65536").End(xlUp).Row) = vSplArr
    Msgbox "Fertig Master!"
End Sub
Sub Loesche_DoppleteZeilen()
    'doppelte Zeilen l�schen
    Dim temp
    Dim i, n, zn, ZSpalte, ZZeile, counter, tMin As Integer
    Dim Zeilenzahl As Long
    Dim t, tSumSec, tSec As Double
    t = Timer
    counter = 0
    ZSpalte = Application.InputBox(Prompt:="Welche Spalte soll verglichen werden?" & vbLf & vbLf _
    & "(Bitte Spalte als Zahl eingeben)", Title:="Vergleichsspalte", Type:=1)
        Zeilenzahl = ActiveSheet.Cells(Rows.count, ZSpalte).End(xlUp).Row
    ZZeile = Application.InputBox(Prompt:="Ab welcher Zeile soll begonnen werden?" & vbLf & vbLf _
    & "(Bitte Zeile als Zahl eingeben)", Title:="Startzeile", Type:=1)
        Zeilenzahl = ActiveSheet.Cells(Rows.count, ZSpalte).End(xlUp).Row
    For n = ZZeile To Zeilenzahl
        temp = ActiveSheet.Cells(n, ZSpalte).Value
            For i = n To Zeilenzahl
                m = ActiveSheet.Cells(i + 1, ZSpalte).Value
                Do While ActiveSheet.Cells(i + 1, ZSpalte).Value = temp
                    counter = counter + 1
                    ActiveSheet.Cells(i + 1, ZSpalte).EntireRow.Delete
                    Zeilenzahl = Zeilenzahl - 1
                Loop
            Next i
    Next n
    tSumSec = Timer - t
    tMin = CInt(tSumSec / 60)
    tSec = tSumSec - Fix(tSumSec)
    Msgbox "Fertig Master!" & vbLf & vbLf & counter & " dopplete Zeilen gel�scht." & vbLf & vbLf & tMin & " Min. " & tSec & " sec", , "Makrolaufzeit."
End Sub
Sub Loesche_Zeile_wenn_best_String()
    'L�scht alle Untergruppen Zeilen
    Dim rngC As Range
    Dim strSH As String
    Dim a, b, FRow, LRow, FColumn, LColumn, intL, intPosS, intPosH, tMin As Integer
    Dim t, tSumSec, tSec As Double
    t = Timer
    FRow = Range(Selection.Address).Row
    LRow = Range(Selection.Address).Row + Selection.Rows.count - 1
    FColumn = Range(Selection.Address).Column
    LColumn = Range(Selection.Address).Column + Selection.Columns.count - 1
    Application.ScreenUpdating = False
    For a = LRow To FRow Step -1
        If InStr(ActiveSheet.Cells(a, 1), "-�-�") > 0 Or InStr(ActiveSheet.Cells(a, 1), "-�- -�") > 0 Then
        ActiveSheet.Rows(a).Delete
        End If
    Next
    Application.ScreenUpdating = True
    tSumSec = Timer - t
    tMin = CInt(tSumSec / 60)
    tSec = tSumSec - Fix(tSumSec)
    Msgbox "Fertig Master!" & vbLf & vbLf & tMin & " Min. " & tSec & " sec", , "Makrolaufzeit."
End Sub

Sub Loesche_LeereZeilen()
    'L�scht alle Zeilen ohne Wert in erster Spalte
    Dim rngC As Range
    Dim strSH As String
    Dim a, b, FRow, LRow, FColumn, LColumn, intL, intPosS, intPosH, tMin As Integer
    Dim t, tSumSec, tSec As Double
    t = Timer
    FRow = Range(Selection.Address).Row
    LRow = Range(Selection.Address).Row + Selection.Rows.count - 1
    FColumn = Range(Selection.Address).Column
    LColumn = Range(Selection.Address).Column + Selection.Columns.count - 1
    Application.ScreenUpdating = False
    For a = LRow To FRow Step -1
        If ActiveSheet.Cells(a, 1) = "" Then
        ActiveSheet.Rows(a).Delete
        End If
    Next
    Application.ScreenUpdating = True
    tSumSec = Timer - t
    tMin = CInt(tSumSec / 60)
    tSec = tSumSec - Fix(tSumSec)
    Msgbox "Fertig Master!" & vbLf & vbLf & tMin & " Min. " & tSec & " sec", , "Makrolaufzeit."
End Sub


