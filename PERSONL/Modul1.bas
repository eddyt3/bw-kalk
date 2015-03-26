Attribute VB_Name = "Modul1"
Sub Tabellenblatt_ausblenden()
    'On Error Resume Next
    Dim wksWorksheet As Worksheet
    ActiveWorkbook.Unprotect "bw"
    ActiveWorkbook.ActiveSheet().Visible = False
    'ActiveWorkbook.ActiveSheet().Visible = xlVeryHidden 'lässt sich vom Benutzer nicht einblenden
    For Each wksWorksheet In ActiveWorkbook.Worksheets
        If wksWorksheet.Range("A1") = "Steuerungsdaten" Then
        wksWorksheet.Visible = False
        End If
    Next wksWorksheet
End Sub
Sub Tabellenblaetter_einblenden()
    'On Error Resume Next
    Dim wksWorksheet As Worksheet
    For Each wksWorksheet In ActiveWorkbook.Worksheets
        wksWorksheet.Visible = xlSheetVisible
    Next wksWorksheet
End Sub
Sub Alle_Tabellenblaetter_protect()
    Dim wksWorksheet As Worksheet
    For Each wksWorksheet In ActiveWorkbook.Sheets
        wksWorksheet.Protect "bw"
    Next wksWorksheet
End Sub
Sub Alle_Tabellenblaetter_unprotect()
    Dim wksWorksheet As Worksheet
    For Each wksWorksheet In ActiveWorkbook.Sheets
        wksWorksheet.Unprotect "bw"
    Next wksWorksheet
End Sub
Sub KopfFusszeile_eintragen()
    Dim wks As Worksheet
    For Each wks In Worksheets
    With wks.PageSetup
        '.LeftHeader = ActiveWorkbook.Name
        '.RightHeader = Format(Date, "dd.mmmm.yyyy")
        .LeftFooter = "&""Verdana""&06" & Application.UserName
        .CenterFooter = "&""Verdana""&06" & ActiveWorkbook.Path & "\" & ActiveWorkbook.Name
        .RightFooter = "&""Verdana""&06" & Format("&D", "dd.mm.yy") & "&T"
    End With
    Next wks
    MsgBox "Fertig Master!"
End Sub
Sub Autoformen_loeschen_Alle()
    Dim shpShape As Shape
    For Each shpShape In ActiveSheet.Shapes
        If shpShape.Type = msoAutoShape Then
            shpShape.Delete
        End If
    Next shpShape
    MsgBox "Fertig Master!"
End Sub
Sub Autoformen_loeschen_Bereich()
    'Autoformen in einem bestimmten Bereich löschen
    Dim c As Range, Sh As Shape
    FRow = Range(Selection.Address).Row
    LRow = Range(Selection.Address).Row + Selection.Rows.count - 1
    FColumn = Range(Selection.Address).Column
    LColumn = Range(Selection.Address).Column + Selection.Columns.count - 1

   Set c = Range(Cells(FRow, FColumn), Cells(LRow, LColumn))
   For Each Sh In ActiveSheet.Shapes
      If Sh.Top > c.Top And Sh.Height < c.Height Then
         If Sh.Left > c.Left And Sh.Width < c.Width Then
            Sh.Delete
         End If
      End If
   Next
    MsgBox "Fertig Master!"
End Sub
Sub Anzahl_Markierte_Zellen()
    FRow = Range(Selection.Address).Row
    LRow = Range(Selection.Address).Row + Selection.Rows.count - 1
    FColumn = Range(Selection.Address).Column
    LColumn = Range(Selection.Address).Column + Selection.Columns.count - 1
    xRow = LRow - FRow + 1
    xColumn = LColumn - FColumn + 1
    xZellen = xRow * xColumn
    MsgBox (xRow & " Zeilen und " & xColumn & " Spalten mit " & xZellen & " Zellen sind ausgewählt.")
    'Debug.Print FRow & "/" & LRow & "/" & FColumn & "/" & LColumn
    'Debug.Print xRow & "/" & xColumn & "/" & xZellen
End Sub
Public Sub Zellen_Aufrunden_Ganzzahl()
    'Aufrunden aller markierter Zellen auf Ganzzahl
    Dim rngC As Range
    Dim zahl As Double
    Dim a, b, FRow, LRow, FColumn, LColumn As Integer
    
    FRow = Range(Selection.Address).Row
    LRow = Range(Selection.Address).Row + Selection.Rows.count - 1
    FColumn = Range(Selection.Address).Column
    LColumn = Range(Selection.Address).Column + Selection.Columns.count - 1
    
    For a = FRow To LRow
        For b = FColumn To LColumn
            zahl = ActiveSheet.Cells(FRow, FColumn).Value
            If zahl > 0 Then
                ActiveSheet.Cells(FRow, FColumn).Value = Int(zahl + 0.99)
            End If
            FColumn = FColumn + 1
        Next
        FRow = FRow + 1
        FColumn = Range(Selection.Address).Column
    Next
    MsgBox "Fertig Master!"
End Sub
Public Sub Topix_SollHaben2Minus_umwandeln()
    Dim rngC As Range
    Dim strSH As String
    Dim a, b, FRow, LRow, FColumn, LColumn, intL, intPosS, intPosH As Integer
    FRow = Range(Selection.Address).Row
    LRow = Range(Selection.Address).Row + Selection.Rows.count - 1
    FColumn = Range(Selection.Address).Column
    LColumn = Range(Selection.Address).Column + Selection.Columns.count - 1
    For a = FRow To LRow
        For b = FColumn To LColumn
            strSH = ActiveSheet.Cells(FRow, FColumn).Value
            intL = Len(strSH)
            intPosS = InStr(1, strSH, "S")
            intPosH = InStr(1, strSH, "H")
            If intPosS > 0 Then
                ActiveSheet.Cells(FRow, FColumn).Replace What:="S", Replacement:="", lookat:=xlPart
                strSH = ActiveSheet.Cells(FRow, FColumn).Value
                ActiveSheet.Cells(FRow, FColumn).Value = CDbl(strSH * -1)
            End If
            If intPosH > 0 Then
                ActiveSheet.Cells(FRow, FColumn).Replace What:="H", Replacement:="", lookat:=xlPart
                strSH = ActiveSheet.Cells(FRow, FColumn).Value
                ActiveSheet.Cells(FRow, FColumn).Value = CDbl(strSH * 1)
            End If
            FColumn = FColumn + 1
        Next
        FRow = FRow + 1
        FColumn = Range(Selection.Address).Column
    Next
    'Selection.NumberFormat = "#.##0,00 ;[Red]-#.##0,00"
    Selection.NumberFormat = "0,00 ;[Red]-0,00"
    MsgBox "Fertig Master!"
End Sub
Public Sub Topix_Vorzeichen_umkehren()
' alle markierte Zellen negieren
    Dim rngC As Range
    Dim curNeg As Currency
    Dim a, b, FRow, LRow, FColumn, LColumn As Integer
    FRow = Range(Selection.Address).Row
    LRow = Range(Selection.Address).Row + Selection.Rows.count - 1
    FColumn = Range(Selection.Address).Column
    LColumn = Range(Selection.Address).Column + Selection.Columns.count - 1
    For a = FRow To LRow
        For b = FColumn To LColumn
            If IsNumeric(ActiveSheet.Cells(FRow, FColumn).Value) Then
                curNeg = ActiveSheet.Cells(FRow, FColumn).Value
                curNeg = curNeg * -1
                ActiveSheet.Cells(FRow, FColumn).Value = curNeg
            End If
            FColumn = FColumn + 1
        Next
        FRow = FRow + 1
        FColumn = Range(Selection.Address).Column
    Next
    'Selection.NumberFormat = "#.##0,00 ;[Red]-#.##0,00"
    Selection.NumberFormatLocal = "#.##0 ;[Rot]-#.##0"
    'Selection.NumberFormat = "0,00 ;[Red]-0,00"
    MsgBox "Fertig Master!"
End Sub
Public Sub Punkt2Komma()
    Dim rngC As Range
    Dim strWert As String
    Dim a, b, FRow, LRow, FColumn, LColumn As Integer
    Dim Dummy As Double
    FRow = Range(Selection.Address).Row
    LRow = Range(Selection.Address).Row + Selection.Rows.count - 1
    FColumn = Range(Selection.Address).Column
    LColumn = Range(Selection.Address).Column + Selection.Columns.count - 1
    For a = FRow To LRow
        For b = FColumn To LColumn
            ActiveSheet.Cells(FRow, FColumn).NumberFormat = "@" 'Txt Format für Replace sicherstellen
            ActiveSheet.Cells(FRow, FColumn).Replace What:=".", Replacement:=",", lookat:=xlPart
            strWert = ActiveSheet.Cells(FRow, FColumn).Value
            If strWert <> "" Then
            Dummy = CDbl(strWert * 1)
                ActiveSheet.Cells(FRow, FColumn).Value = Dummy
                ActiveSheet.Cells(FRow, FColumn).NumberFormat = "#,##0 ;[Red]-#,##0"
            End If
            FColumn = FColumn + 1
        Next
        FRow = FRow + 1
        FColumn = Range(Selection.Address).Column
    Next
    ActiveSheet.Cells(FRow, FColumn).Select
    MsgBox "Fertig Master!"
End Sub
Public Sub Komma2Punkt()
    Dim rngC As Range
    Dim strWert As String
    Dim a, b, FRow, LRow, FColumn, LColumn As Integer
    Dim Dummy As Double
    FRow = Range(Selection.Address).Row
    LRow = Range(Selection.Address).Row + Selection.Rows.count - 1
    FColumn = Range(Selection.Address).Column
    LColumn = Range(Selection.Address).Column + Selection.Columns.count - 1
    For a = FRow To LRow
        For b = FColumn To LColumn
            ActiveSheet.Cells(FRow, FColumn).NumberFormat = "@" 'Txt Format für Replace sicherstellen
            ActiveSheet.Cells(FRow, FColumn).Replace What:=",", Replacement:="§", lookat:=xlPart
    '        ActiveSheet.Cells(FRow, FColumn).Replace What:=",", Replacement:=".", LookAt:=xlPart
    '        strWert = ActiveSheet.Cells(FRow, FColumn).Value
    '        Dummy = CDbl(strWert * 1)
    '        ActiveSheet.Cells(FRow, FColumn).Value = Dummy
    '        ActiveSheet.Cells(FRow, FColumn).NumberFormat = "#.##0,00 ;[Re     d]-#.##0,00"
            FColumn = FColumn + 1
        Next
        FRow = FRow + 1
        FColumn = Range(Selection.Address).Column
    Next
    ActiveSheet.Cells(FRow, FColumn).Select
    MsgBox "Fertig Master!"
End Sub
Public Sub Zellen_mit_0_leeren()
    Dim rngC As Range
    Dim strWert As String
    Dim a, b, FRow, LRow, FColumn, LColumn As Integer
    Dim Dummy As Double
    FRow = Range(Selection.Address).Row
    LRow = Range(Selection.Address).Row + Selection.Rows.count - 1
    FColumn = Range(Selection.Address).Column
    LColumn = Range(Selection.Address).Column + Selection.Columns.count - 1
    For a = FRow To LRow
        For b = FColumn To LColumn
            If ActiveSheet.Cells(FRow, FColumn).Value = "0" Then
                ActiveSheet.Cells(FRow, FColumn).Value = ""
            End If
            FColumn = FColumn + 1
        Next
        FRow = FRow + 1
        FColumn = Range(Selection.Address).Column
    Next
    ActiveSheet.Cells(FRow, FColumn).Select
    MsgBox "Fertig Master!"
End Sub
Sub Sheet_Nach_CSVDatei()
    'hierbei bleibt die Formatierung der Zellen so wie sie angezeigt wird.
    'Es muss alles so formatiert sein wie es später In der CSV sein soll.
    Dim vntFileName As Variant
    Dim lngFn As Long
    Dim rngRow As Excel.Range
    Dim rngCell As Excel.Range
    Dim strDelimiter As String
    Dim strText As String
    Dim strTextCell As String
    Dim bolErsteSpalte As Boolean
    Dim rngColumn As Excel.Range
    Dim wksQuelle As Excel.Worksheet
    strDelimiter = ";" 'deutsches CSV-Format: ";", Englishes CSV-Format: ","
    vntFileName = Application.GetSaveAsFilename("Test.csv", fileFilter:="CSV-File (*.csv),*.csv")
    If vntFileName = False Then Exit Sub
    Set wksQuelle = ActiveSheet  'Beispiel oder: = ActiveWorkbook.Worksheets("Tabelle1")
    lngFn = FreeFile
    Open vntFileName For Output As lngFn
     For Each rngRow In wksQuelle.UsedRange.Rows
      strText = ""
      bolErsteSpalte = True
      For Each rngCell In rngRow.Columns
       strTextCell = rngCell.Text 'Text! inclusive dem NumberFormat der Zelle
       If InStr(1, strTextCell, strDelimiter, 0) Then '## wenn alle Zellen mit " " eingeschlossen werden sollen zeile auskommentieren
        'bewirkt das Werte die den Delimiter enthalten (was eigentlich nicht sein sollte) mit " " eingeschlossen werden
        strTextCell = Chr(34) & strTextCell & Chr(34)
       End If '##
       If bolErsteSpalte Then
        strText = strTextCell
        bolErsteSpalte = False
       Else
        strText = strText & strDelimiter & strTextCell
       End If
      Next
      Print #lngFn, strText
     Next
    Close lngFn
    MsgBox "Fertig Master!"
End Sub
Sub Wbk_save_all()
    'Alle Arbeitsmappen sichern
    Dim wbkWorkbook As Workbook
    For Each wbkWorkbook In Application.Workbooks
        'wbkWorkbook.Activate
        wbkWorkbook.Save
    Next wbkWorkbook
End Sub
Sub Wbk_save_close_other()
    'Alle anderen Arbeitsmappen schließen
    Dim wbkWorkbook As Workbook
    For Each wbkWorkbook In Application.Workbooks
        If wbkWorkbook.Name <> ThisWorkbook.Name Then
            wbkWorkbook.Close SaveChanges:=True
        End If
    Next wbkWorkbook
End Sub
Sub Wbk_close_other()
    'Alle anderen Arbeitsmappen schließen
    Dim wbkWorkbook As Workbook
    For Each wbkWorkbook In Application.Workbooks
        If wbkWorkbook.Name <> ThisWorkbook.Name Then
            wbkWorkbook.Close SaveChanges:=False
        End If
    Next wbkWorkbook
End Sub
Sub Wbk_save_close_all()
    'Alle Arbeitsmappen speichern und schließen
    Dim wbkWorkbook As Workbook
    For Each wbkWorkbook In Application.Workbooks
        If wbkWorkbook.Name <> "PERSONL.XLSB" Then
            wbkWorkbook.Close SaveChanges:=True
        End If
    Next wbkWorkbook
End Sub
Sub Inhalt_Cutten()
    'Spalteninhalte auf eine bestimmte Länge kürzen
    On Error Resume Next
    Dim count, Anzahl, FRow, LRow, FColumn, LColumn As Integer
    FRow = Range(Selection.Address).Row
    LRow = Range(Selection.Address).Row + Selection.Rows.count - 1
    FColumn = Range(Selection.Address).Column
    LColumn = Range(Selection.Address).Column + Selection.Columns.count - 1
    count = 0
    Anzahl = Application.InputBox(Prompt:="Anzahl der Zeichen eingeben", Title:="Textlänge kürzen", Type:=1)
    On Error GoTo 0
        Application.DisplayAlerts = True
    If Anzahl > 0 Then
        Dim rng As Range
        For Each rng In Range(Cells(FRow, FColumn), Cells(LRow, LColumn)).Cells
           If WorksheetFunction.IsText(rng) Then
            If Len(rng) > Anzahl Then
                count = count + 1
            End If
            rng.Value = Left(rng.Value, Anzahl)
           End If
        Next rng
    Else
        MsgBox ("Bitte nur Werte größer 0 eingeben!")
    End If
    MsgBox (count & " Zellen waren länger als " & Anzahl & " Zeichen.")
End Sub
Sub Zeichen_einfuegen()
    'Zahl in String umwandeln und ein bestimmtes Zeichen an bestimmter Stelle einfügen
    Dim a, b, FRow, LRow, FColumn, LColumn, leftPos As Integer
    Dim strOrg, strOrg1, strOrg2, strNew, strSign As String
    FRow = Range(Selection.Address).Row
    LRow = Range(Selection.Address).Row + Selection.Rows.count - 1
    FColumn = Range(Selection.Address).Column
    LColumn = Range(Selection.Address).Column + Selection.Columns.count - 1
    leftPos = Application.InputBox(Prompt:="Einfügen nach welcher Stelle von links?", Title:="Zeichen einfügen", Type:=1)
    strSign = Application.InputBox(Prompt:="Welches Zeichen einfügen?", Title:="Zeichen einfügen", Type:=2)
    For a = FRow To LRow
        For b = FColumn To LColumn
            strOrg = CStr(ActiveSheet.Cells(FRow, FColumn))
            If leftPos > Len(strOrg) Or leftPos < 0 Then
                MsgBox ("Achtung! Zeichenkette ist kürzer (" & Len(strOrg) & " Zeichen) als einzufügende Position (nach " & leftPos & " Zeichen).")
                Exit Sub
            End If
            If leftPos < 0 Then
                MsgBox ("Achtung! Bitte nur Werte zwischen 0 und " & Len(strOrg) & " eingeben.")
                Exit Sub
            End If
            'mit zeichen auffüllen
            'strOrg = String(6 - Len(strOrg), "0") & strOrg
            strOrg1 = Mid(strOrg, 1, leftPos)
            strOrg2 = Mid(strOrg, (leftPos + 1), Len(strOrg))
            strNew = strOrg1 & strSign & strOrg2
            ActiveSheet.Cells(FRow, FColumn) = strNew
            FColumn = FColumn + 1
        Next
        FRow = FRow + 1
        FColumn = Range(Selection.Address).Column
    Next
    ActiveSheet.Cells(FRow, FColumn).Select
    '    NumAsString = CStr(Nummer)
    '    NumAsString = String(6 - Len(NumAsString), "0") & NumAsString
    '    Me!TXTlbImplantat = Mid(NumAsString, 1, 3) & "-" & Mid(NumAsString, 4, 6)
    MsgBox "Fertig Master!"
End Sub
Sub Blattschutz_loeschen()
    On Error Resume Next
    For i = 65 To 66
    For j = 65 To 66
    For k = 65 To 66
    For L = 65 To 66
    For m = 65 To 66
    For n = 65 To 66
    For o = 65 To 66
    For p = 65 To 66
    For q = 65 To 66
    For r = 65 To 66
    For s = 65 To 66
    For t = 32 To 126
    ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(L) & Chr(m) & _
    Chr(n) & Chr(o) & Chr(p) & Chr(q) & Chr(r) & Chr(s) & Chr(t)
    Next t
    Next s
    Next r
    Next q
    Next p
    Next o
    Next n
    Next m
    Next L
    Next k
    Next j
    Next i
    MsgBox "Fertig Master!"
End Sub
Sub ShapesUmbenennen()
    Dim ws As Worksheet, i As Integer, NeuerName As String
    Set ws = ThisWorkbook.ActiveSheet
    For i = 1 To ws.Shapes.count
        ws.Shapes(i).Visible = False
    Next
    For i = 1 To ws.Shapes.count
        ws.Shapes(i).Visible = True
        ActiveWindow.ScrollColumn = ws.Shapes(i).TopLeftCell.Column
        ActiveWindow.ScrollRow = ws.Shapes(i).TopLeftCell.Row
        NeuerName = InputBox("Bestätige den Namen '" & ws.Shapes(i).Name & "'" & vbLf & _
            "oder gebe einen neuen Namen ein:", "Shapes Umbenennen", ws.Shapes(i).Name)
        If NeuerName = "" Then
            Exit For
        Else
            If ws.Shapes(i).Name <> NeuerName Then
                ws.Shapes(i).Name = NeuerName
            End If
        End If
        ws.Shapes(i).Visible = False
    Next
    For i = 1 To ws.Shapes.count
        ws.Shapes(i).Visible = True
    Next
    'ws.Range("A1").Activate
End Sub
Sub ShapesAlleEinblenden()
    Dim ws As Worksheet, i As Integer
    Set ws = ThisWorkbook.ActiveSheet
    For i = 1 To ws.Shapes.count
        ws.Shapes(i).Visible = True
    Next
    'ws.Range("A1").Activate
End Sub
Sub color_trend()
    'zeilenweise Trend (Vergleich mit Vorgänger) einfärben im markierten Bereich
    Dim rngC As Range
    Dim zahl1, zahl2, toleranz As Double
    Dim a, b, FRow, LRow, FColumn, LColumn As Integer
    
    FRow = Range(Selection.Address).Row
    LRow = Range(Selection.Address).Row + Selection.Rows.count - 1
    FColumn = Range(Selection.Address).Column
    LColumn = Range(Selection.Address).Column + Selection.Columns.count - 1
    toleranz = Application.InputBox(Prompt:="Toleranz in Prozent", Title:="Toleranzgrenze", Type:=1)
    If toleranz > 0 And toleranz < 100 Then
        toleranz = toleranz / 100
    End If
    If toleranz < 0.01 And toleranz > 1 Then
        MsgBox "Bitte nur Werte zwischen 1 und 100 eingeben."
    End If
    For a = FRow To LRow
        For b = FColumn To LColumn
            If b <> FColumn Then
                zahl1 = ActiveSheet.Cells(a, b - 1).Value
                zahl2 = ActiveSheet.Cells(a, b).Value
                If zahl2 > (zahl1 * (1 + toleranz)) Or zahl2 < (zahl1 * (1 - toleranz)) Then
                    If zahl2 > (zahl1 * (1 + toleranz)) Then
                        ActiveSheet.Cells(a, b).Interior.ColorIndex = 43
                        Else
                        If zahl2 < (zahl1 * (1 - toleranz)) Then
                            ActiveSheet.Cells(a, b).Interior.ColorIndex = 46
                            Else
                            ActiveSheet.Cells(a, b).Interior.ColorIndex = 33
                        End If
                    End If
                Else
                If zahl2 = 0 Then
                    ActiveSheet.Cells(a, b).Interior.ColorIndex = 46
                    Else
                    ActiveSheet.Cells(a, b).Interior.ColorIndex = 33
                End If
                End If
                FColumn = FColumn + 1
            End If
        Next
        FRow = FRow + 1
        FColumn = Range(Selection.Address).Column
    Next
    MsgBox "Fertig Master!"
End Sub


