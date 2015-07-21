Attribute VB_Name = "Modul5"
Sub Objektliste_Wkb_erstellen()
    'alle OLE Objekte auflisten
    'alle OLEObject-Objekte befinden sich ebenfalls in der Shapes-Auflistung
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
'        For Each shp In sh.Shapes
'            shListe.Cells(ze, sp).Value = shp.Name
'            ze = ze + 1
'        Next
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
    'alle OLEObject-Objekte befinden sich ebenfalls in der Shapes-Auflistung
    Dim obj As OLEObject
    Dim shp As Shape
    Dim tmpArray() As Variant
    Dim i As Integer
    i = 1
    For Each obj In ActiveSheet.OLEObjects
        'Debug.Print obj.Name
        ReDim Preserve tmpArray(i)
        tmpArray(i) = obj.Name
        i = i + 1
    Next
'    For Each shp In ActiveSheet.Shapes
'        'Debug.Print shp.Namein837net

'        ReDim Preserve tmpArray(i)
'        tmpArray(i) = shp.Name
'        i = i + 1
'    Next
    If (Not Not tmpArray) <> 0 Then 'pr�fen ob array initialisiert ist, ansonsten Laufzeitfehler
        BubbleSort tmpArray
        For i = 1 To UBound(tmpArray)
            Debug.Print i & ":" & tmpArray(i)
            'Debug.Print tmpArray(i)
        Next i
        'MsgBox UBound(tmpArray) & " Elemente gefunden."
    End If
    Exit Sub
End Sub

