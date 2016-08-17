Attribute VB_Name = "Modul2"
Option Explicit
Sub Verknuepfungen_Suchen()
Dim rngGefundeneZelle As Range
Dim intI%, intN%, intAbfrage%, intZähler%, intNamenAbfrage%, intLöschZähler%
Dim strSuchbegriff$, strAktuelleZelle$, strErsteZelle$
Dim objName As Object
strSuchbegriff = "]"
If strSuchbegriff = "" Then Exit Sub
ReDim strBereiche(0)
ReDim intBlätter(0)
intZähler = 0
For intI = 1 To Worksheets.count
    If Sheets(intI).ProtectContents Then
        MsgBox "Das Blatt " & Sheets(intI).Name & " ist geschützt." & Chr(10) & "Entfernen Sie bitte zuerst den Blattschutz.", vbOKOnly + vbInformation, "Blatt geschützt!"
        GoTo nächstesBlatt
    End If
    Set rngGefundeneZelle = Worksheets(intI).Cells.Find(strSuchbegriff, lookat:=xlPart, LookIn:=xlFormulas)
    If Not rngGefundeneZelle Is Nothing Then
        'erste gefundene Zelle auf dem aktuellen Blatt:
        strErsteZelle = rngGefundeneZelle.Address(False, False)
        intZähler = intZähler + 1
        ReDim Preserve intBlätter(intZähler)
        ReDim Preserve strBereiche(intZähler)
        intBlätter(intZähler - 1) = intI
        strBereiche(intZähler - 1) = strErsteZelle
        Do While strAktuelleZelle <> strErsteZelle
            'nächste gefundene Zelle(n) auf dem aktuellen Blatt
            Set rngGefundeneZelle = Worksheets(intI).Cells.FindNext(After:=rngGefundeneZelle)
            strAktuelleZelle = rngGefundeneZelle.Address(False, False)
            If strErsteZelle <> strAktuelleZelle Then
                intZähler = intZähler + 1
                ReDim Preserve intBlätter(intZähler)
                ReDim Preserve strBereiche(intZähler)
                intBlätter(intZähler - 1) = intI
                strBereiche(intZähler - 1) = strAktuelleZelle
            End If
        Loop
        strAktuelleZelle = ""
        strErsteZelle = ""
    End If
nächstesBlatt:
Next
Set rngGefundeneZelle = Nothing
intLöschZähler = 0
For intN = 0 To intZähler - 1
    Sheets(intBlätter(intN)).Select
    Range(strBereiche(intN)).Select
    intAbfrage = MsgBox("Auf dem Blatt " & Sheets(intBlätter(intN)).Name & " wurde in der Zelle " & strBereiche(intN) & " eine Verknüpfung gefunden." & Chr(10) & "Die Formel lautet:" & Chr(10) & Chr(10) & ActiveCell.Formula & Chr(10) & Chr(10) & "Soll sie gelöscht werden?", vbYesNo + vbQuestion, "Verknüpfung gefunden")
    If intAbfrage = vbYes Then
        Range(strBereiche(intN)).ClearContents
        intLöschZähler = intLöschZähler + 1
    End If
Next
For Each objName In ActiveWorkbook.Names
    If InStr(1, objName.Value, strSuchbegriff) > 1 Then
        intZähler = intZähler + 1
        intNamenAbfrage = MsgBox("In einem Namen besteht eine Verknüpfung." & Chr(10) & "Bezieht sich auf: " & objName.Value & Chr(10) & "Name: " & objName.Name, vbYesNo + vbQuestion, "Soll der Name gelöscht werden?")
        If intNamenAbfrage = vbYes Then
            objName.Delete
            intLöschZähler = intLöschZähler + 1
        End If
    End If
Next
If intZähler = 0 Then
    MsgBox "Keine Verknüpfung gefunden oder die Blätter sind geschützt.", vbOKOnly + vbInformation, "Fertig!"
Else
    MsgBox "Es wurden insgesamt " & intZähler & " Verknüpfung(en) gefunden und davon " & intLöschZähler & " gelöscht.", vbOKOnly + vbInformation, "Fertig!"
End If
End Sub
Sub Verknuepfungen_Suchen1()
' Verknüpfungen auflisten
' nach'* H. Ziplies
'* 22.08.03, 24.04.04; 31.07.05; 18.10.05
'* erstellt von Hajo.Ziplies@web.de
'* http://home.media-n.de/ziplies/
' geänd. von Erich G. 20.07.2008
   Dim RaZelle As Range, ByMldg As Integer, Sh As Worksheet
   Dim namN As Name, zz As Long

   For Each Sh In Worksheets
      If Sh.Name = "Verknüpfungen" Then
         ByMldg = MsgBox("Tabelle Verknüfungen schon vorhanden - Löschen?", _
            vbYesNo + vbQuestion, "Löschabfrage ?", "", 0)
         If ByMldg <> 6 Then Exit Sub
         Sh.Cells.Clear
         Exit For
      End If
   Next Sh
   If ByMldg <> 6 Then _
      Sheets.Add(After:=Worksheets(Worksheets.count)).Name = "Verknüpfungen"
   zz = 1
   Cells(zz, 1) = "Zelle"
   Cells(zz, 2) = "Tabelle"
   Cells(zz, 3) = "Formel"
   For Each Sh In Worksheets
      If Sh.Name <> "Verknüpfungen" Then
         ' Sh.Unprotect ' .unprotect "Passwort"
         For Each RaZelle In Sh.UsedRange
            If RaZelle.HasFormula And InStr(RaZelle.Formula, ":\") > 1 Then
               zz = zz + 1
               Cells(zz, 1) = RaZelle.Address(0, 0)
               Cells(zz, 2) = Sh.Name
               Cells(zz, 3) = "'" & RaZelle.Formula
            End If
         Next RaZelle
         ' Sh.Protect    ' .Protect "Passwort"
      End If
   Next Sh
'  ------------------------------------------------------------ Namen
   zz = zz + 3
   Cells(zz, 1) = "Name"
   Cells(zz, 2) = "Bezug"
   For Each namN In ActiveWorkbook.Names
      zz = zz + 1
      Cells(zz, 1) = namN.Name
      With Cells(zz, 2)
         If InStr(namN, "REF") <> 0 Then
               .Value = namN '"Fehlerhaft"
               .Font.Bold = True
               .Font.ColorIndex = 3
         ElseIf InStr(namN, "\") <> 0 Then
               .Value = namN
               .Font.Bold = True
               .Font.ColorIndex = 4
         Else
               .Value = Mid(namN, 2)
         End If
      End With
   Next
End Sub
Sub Verknuepfungen_Suchen2()
    Dim Tab1 As Object
    Dim Zelle1 As Object
    Dim AlleFormeln As Object
    Dim NeuesBlatt As Worksheet
    Dim Zeile As Integer
    On Error GoTo 0
    Set NeuesBlatt = ActiveWorkbook.Worksheets.Add(ActiveWorkbook.Worksheets(1))
    On Error Resume Next
    NeuesBlatt.Name = "Externe Verknüpfungen"
    NeuesBlatt.Range("a3").Formula = "Arbeitsblatt"
    NeuesBlatt.Range("b3").Formula = "Zelle"
    NeuesBlatt.Range("c3").Formula = "Externer Bezug"
    NeuesBlatt.Range("d3").Formula = "Aktueller Wert"
    NeuesBlatt.Range("a1").Font.Bold = True
    NeuesBlatt.Range("a3:d3").Font.Bold = True
    Zeile = 4
    Application.ScreenUpdating = False
    For Each Tab1 In ActiveWorkbook.Worksheets
     If Tab1.Name <> NeuesBlatt.Name Then
        Set AlleFormeln = Nothing
        Set AlleFormeln = Tab1.Cells.SpecialCells(xlFormulas, 23)
        'MsgBox (AlleFormeln)
        If AlleFormeln Then
        For Each Zelle1 In AlleFormeln
            If InStr(Zelle1.Formula, "\") > 0 Then
             'MsgBox (Zelle1)
            'If Tab1.Name <> NeuesBlatt.Name Then
             NeuesBlatt.Cells(Zeile, 1).Formula = Tab1.Name
             NeuesBlatt.Cells(Zeile, 2).Formula = Zelle1.AddressLocal(False, False)
             NeuesBlatt.Cells(Zeile, 3).Formula = Right$(Zelle1.FormulaLocal, Len(Zelle1. _
FormulaLocal) - 2)
             NeuesBlatt.Cells(Zeile, 4).Formula = Zelle1.Value
             Zeile = Zeile + 1
            'End If
'                If Zelle1.HasArray Then
'                    Zelle1.CurrentArray.Select
'                    Selection.Copy
'                    Selection.PasteSpecial Paste:=xlValues
'                    Application.CutCopyMode = False
'                Else
'                    Zelle1.Formula = Zelle1.Value
'                End If
            End If
        Next Zelle1
       End If
      End If
    Next Tab1
FormatiereDenRest:
    For Zeile = 1 To 4
     NeuesBlatt.Cells(1, Zeile).EntireColumn.AutoFit
    Next
    NeuesBlatt.Range("a1").Formula = "Externe Verknüpfungen in " + UCase$(ActiveWorkbook.Name)
    On Error GoTo 0
    'Aufräumen der Tabelle
    Zeile = 4
    While NeuesBlatt.Cells(Zeile, 1).Text <> ""
     If NeuesBlatt.Cells(Zeile, 2).Text = "" Then
      NeuesBlatt.Cells(Zeile, 2).EntireRow.Delete
     Else
      Zeile = Zeile + 1
     End If
    Wend
    Application.ScreenUpdating = True
    If Zeile = 4 Then
     MsgBox "Keine externen Verknüpfungen gefunden", vbOKOnly, "Keine Verknüpfungen"
     Application.DisplayAlerts = False
     NeuesBlatt.Delete
     Application.DisplayAlerts = True
    End If
    Application.ScreenUpdating = True
    Exit Sub
FehlerVerknüpfungAuflösen:
    If Err = 1004 Then Resume Next Else MsgBox ("Keine Zellen mit externen Verknüpfungen gefunden ")
    GoTo FormatiereDenRest
End Sub
Sub Verknuepfungen_aendern()
    'Beispiel - Verknüpfungen i. einer Arbeitsmappe ändern
    Dim wbkMappe As Workbook
    Dim varVLink As Variant
    Dim i, e As Integer
    Dim strPrefix, strPath, strFile, strRefFile As String
    
    strPath = ThisWorkbook.Path
    strFile = ThisWorkbook.Name
    strPrefix = Left(strFile, 4)
    strRefFile = strPrefix & "_Verrechnung.xls" 'verknüpfte Tabelle
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
    MsgBox "Fertig Master!"
End Sub
Sub datum_splitten()
    'Datum in Tag, Monat und Jahr aufsplitten
    'Achtung neben Datum drei Leerspalten einfügen
    Dim lZeile As Long
    Dim vDatArr As Variant
    Dim vSplArr() As Variant
    result = MsgBox("Achtung, neben der Datumsspalte müssen 3 Leerspalten vorhanden sein!", 1, "Hinweis")
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
    MsgBox "Fertig Master!"
End Sub
Sub Loesche_DoppleteZeilen()
    'doppelte Zeilen löschen
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
    MsgBox "Fertig Master!" & vbLf & vbLf & counter & " dopplete Zeilen gelöscht." & vbLf & vbLf & tMin & " Min. " & tSec & " sec", , "Makrolaufzeit."
End Sub
Sub Loesche_Zeile_wenn_best_String()
    'Löscht alle Untergruppen Zeilen
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
        If InStr(ActiveSheet.Cells(a, 1), "- - ") > 0 Or InStr(ActiveSheet.Cells(a, 1), "- - - ") > 0 Then
        ActiveSheet.Rows(a).Delete
        End If
    Next
    Application.ScreenUpdating = True
    tSumSec = Timer - t
    tMin = CInt(tSumSec / 60)
    tSec = tSumSec - Fix(tSumSec)
    MsgBox "Fertig Master!" & vbLf & vbLf & tMin & " Min. " & tSec & " sec", , "Makrolaufzeit."
End Sub

Sub Loesche_markierteZeilen_Spalte1_leer()
    'Löscht alle Zeilen ohne Wert in erster Spalte
    Dim rngC As Range
    Dim strSH As String
    Dim a, b, FRow, LRow, FColumn, LColumn, intL, intPosS, intPosH, tMin  As Integer
    Dim t, tSumSec, tSec As Double
    t = Timer
    FRow = Range(Selection.Address).Row
    LRow = Range(Selection.Address).Row + Selection.Rows.count - 1
    Zeilenzahl = LRow - FRow + 1
    Application.ScreenUpdating = False
    For a = LRow To FRow Step -1
        If ActiveSheet.Cells(a, 1) = "" Then
            ActiveSheet.Rows(a).Delete
            b = b + 1
        End If
    Next
    Application.ScreenUpdating = True
    tSumSec = Timer - t
    tMin = CInt(tSumSec / 60)
    tSec = tSumSec - Fix(tSumSec)
    MsgBox "Fertig Master!" & vbLf & vbLf & b & " von " & Zeilenzahl & " gelöscht." _
    & vbLf & vbLf & tMin & " Min. " & tSec & " sec", , "Makrolaufzeit."
End Sub
Sub Loesche_markierteZeilen_SpalteX_leer()
    'Löscht alle Zeilen ohne Wert in Spalte X
    Dim rngC As Range
    Dim strSH As String
    Dim a, b, FRow, LRow, intL, intPosS, intPosH, tMin, ZSpalte, Zeilenzahl As Integer
    Dim t, tSumSec, tSec As Double
    t = Timer
    FRow = Range(Selection.Address).Row
    LRow = Range(Selection.Address).Row + Selection.Rows.count - 1
    Zeilenzahl = LRow - FRow + 1
    ZSpalte = Application.InputBox(Prompt:="Welche Spalte soll verglichen werden?" & vbLf & vbLf _
    & "(Bitte Spalte als Zahl eingeben)", Title:="Vergleichsspalte", Type:=1)
    b = 0
    Application.ScreenUpdating = False
    For a = LRow To FRow Step -1
        If ActiveSheet.Cells(a, ZSpalte) = "" Then
            ActiveSheet.Rows(a).Delete
            b = b + 1
        End If
    Next
    Application.ScreenUpdating = True
    tSumSec = Timer - t
    tMin = CInt(tSumSec / 60)
    tSec = tSumSec - Fix(tSumSec)
    MsgBox "Fertig Master!" & vbLf & vbLf & b & " von " & Zeilenzahl & " gelöscht." _
    & vbLf & vbLf & tMin & " Min. " & tSec & " sec", , "Makrolaufzeit."
End Sub
Sub Loesche_Spalte_wenn_Zelle_erste_Zeile_leer()
'Wenn in irgendeiner der Zellen aus Zeile 1 kein Wert steht, wird die gesamte Spalte gelöscht
'Version ohne Schleife
    ActiveSheet.Rows(1).SpecialCells(xlCellTypeBlanks).EntireColumn.Delete
End Sub
Sub Loesche_Spalte_wenn_Zelle_erste_Zeile_leer1()
'Wenn in irgendeiner der Zellen aus Zeile 1 kein Wert steht, wird die gesamte Spalte gelöscht
'Version mit Schleife
    Dim lngSpalte As Long
    Dim wksA As Worksheet
    Dim lngLetzteSpalte As Long
    Set wksA = ActiveSheet
    lngLetzteSpalte = wksA.Cells(1, wksA.Columns.count).End(xlToLeft).Column
    For lngSpalte = lngLetzteSpalte To 1 Step -1
        If Trim(wksA.Cells(1, lngSpalte).Text) = "" Then
            wksA.Columns(lngSpalte).Delete
        End If
    Next
End Sub
Sub List_Location_Size_for_all_VB_Buttons()
    'Problem: unterschiedliche Größen der Buttons bei unterschiedlichen Bildschirmauflösungen
    'Macro liest alle Buttonformate (Standard) aus
    'den Code aus dem Direktbereich in die Workbook_Open() Sub übernehmen (Komma noch durch Punkt ersetzen)
    'Danach werden bei jedem Öffnen die Buttons auf ihre Standardwerte zurückgesetzt unabhängig der aktuellen Bildschirmauflösung
    Dim ShCounter As Long, Sh As Shape
    Dim i As Integer
    ShCounter = 0
    DebugClear
    'Debug.Print "fntSize=10"
    DebugPrint "fntSize=10"
    For i = 1 To ActiveWorkbook.Sheets.count - 1
      With ActiveWorkbook.Sheets(i)
       For Each Sh In .Shapes
        If Sh.Type = msoOLEControlObject Then  'Only list VB buttons
            ShCounter = ShCounter + 1
    ' Code für Direktbereich
    '        Debug.Print "WITH WorkSheets("; Chr(34); Sheets(i).Name; Chr(34); ")."; Sh.Name, "   '"; ShCounter
    '        Debug.Print "   .Height="; Sh.Height;
    '        Debug.Print ": .Width="; Sh.Width;
    '        Debug.Print ": .Top="; Sh.Top;
    '        Debug.Print ": .Left = "; Sh.Left;
    '        Debug.Print ": .FontSize = fntSize"
    '        Debug.Print "END WITH"
    
    'Code für Ausgabe in debug.log File, wenn Puffer Direktbereich zu klein
            DebugPrint "WITH WorkSheets(" & Chr(34) & Sheets(i).Name & Chr(34) & ")." & Sh.Name & "   '" & ShCounter
            DebugPrint "   .Height=" & Sh.Height & ": .Width=" & Sh.Width & ": .Top=" & Sh.Top & ": .Left = " & Sh.Left & ": .FontSize = fntSize"
            DebugPrint "END WITH"
    '
         End If
        Next Sh
      End With
    Next i
    MsgBox "Fertig Master!" & vbLf & vbLf & ShCounter & " VB Buttonformate exportiert."
End Sub
Sub Test()
'zu List_Location_Size_for_all_VB_Buttons()
'Test File Ausgabe
   Dim i As Long

   DebugClear
   For i = 1 To 100
      DebugPrint "Hello world.  " & Now
   Next
End Sub
Sub DebugPrint(s As String)
'zu List_Location_Size_for_all_VB_Buttons()
   Static fso As Object
   If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")
   With fso.OpenTextFile(ActiveWorkbook.Path & "\debug.log", 8, True, -1)
      .WriteLine s
      .Close
   End With
End Sub
Sub DebugClear()
'zu List_Location_Size_for_all_VB_Buttons()
   CreateObject("Scripting.FileSystemObject").CreateTextFile ActiveWorkbook.Path & "\debug.log", True, True
End Sub
Sub references_eigenschaften()
'verwendete VBA Verweise einer XLS anzeigen lassen
    Dim b, r, ref
    For Each b In Workbooks
        Set ref = b.VBProject.References
        For Each r In ref
            Debug.Print b.Name, r.Name, r.Type, r.FullPath
        Next
    Next
End Sub
