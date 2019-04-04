Attribute VB_Name = "Modul5"
Sub Objektliste_Wkb_erstellen()
    'alle OLE Objekte auflisten
    'alle OLEObject-Objekte befinden sich ebenfalls in der Shapes-Auflistung
    Dim shListe As Worksheet
    Dim Sh As Worksheet
    Dim obj As OLEObject
    Dim shp As Shape
    Dim sp As Long, ze As Long
    On Error GoTo Objektliste_erstellen
    Set shListe = ActiveWorkbook.Sheets("Objektliste")
    On Error GoTo 0
    shListe.Cells.Clear
    sp = 1: ze = 2
    For Each Sh In ActiveWorkbook.Worksheets
        shListe.Cells(1, sp).Value = Sh.Name
        For Each obj In Sh.OLEObjects
            shListe.Cells(ze, sp).Value = "Obj:" & obj.Name
            ze = ze + 1
        Next
        For Each shp In Sh.Shapes
            shListe.Cells(ze, sp).Value = "Shp:" & shp.Name
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
    'alle OLEObject-Objekte befinden sich ebenfalls in der Shapes-Auflistung
    Dim obj As OLEObject
    Dim shp As Shape
    Dim tmpArray() As Variant
    Dim I As Integer
    I = 1
    For Each obj In ActiveSheet.OLEObjects
        'Debug.Print obj.Name
        ReDim Preserve tmpArray(I)
        tmpArray(I) = "Obj:" & obj.Name
        I = I + 1
    Next
    For Each shp In ActiveSheet.Shapes
        'Debug.Print shp.Name

        ReDim Preserve tmpArray(I)
        tmpArray(I) = "Shp:" & shp.Name
        I = I + 1
    Next
    If (Not Not tmpArray) <> 0 Then 'prüfen ob array initialisiert ist, ansonsten Laufzeitfehler
        BubbleSort tmpArray
        For I = 1 To UBound(tmpArray)
            Debug.Print I & ":" & tmpArray(I)
            'Debug.Print tmpArray(i)
        Next I
        'MsgBox UBound(tmpArray) & " Elemente gefunden."
    End If
    Exit Sub
End Sub
Sub Objektliste_ActiveX_controls_ActiveSheet()
    Dim Ws As Worksheet
    Dim count, countB As Integer
    Set Ws = ActiveSheet
    count = 0
    countB = 0
    For Each OleObj In Ws.OLEObjects
        If OleObj.OLEType = xlOLEControl Then
            ' Nur ActiveX ComboBox
            If TypeName(OleObj.Object) = "ComboBox" Then
                countB = countB + 1
            End If
            count = count + 1
        End If
    Next OleObj
    MsgBox "Number of ActiveX controls: " & count & vbNewLine & "(ComboBoxes of that: " & countB & ")"
End Sub
Sub AlleZellnamenLoeschen()
    'löscht alle individuellen Zellnamen und setzt sie auf Standard z.B. "A1" zurück
  Dim varName As Name
  Dim intResponse As Integer

  intResponse = MsgBox("Alle Namen löschen?", _
  vbYesNo, "Excel Weekly")

  If intResponse = vbNo Then Exit Sub

  For Each varName In ActiveWorkbook.Names
    varName.Delete
  Next varName

End Sub

