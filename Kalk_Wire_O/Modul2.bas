Attribute VB_Name = "Modul2"
Option Explicit
Sub Speichern()
    '
    'Dateiname vorschlagen u. speichern
    '
    Dim knd, format, F, D, i, RB, RP, auflage As String
    Dim strVerzeichnis, strDateiname As String
    knd = Worksheets("Steuerung").Range("B181")
    format = Worksheets("Steuerung").Range("B182")
    F = Worksheets("Steuerung").Range("E181")
    D = Worksheets("Steuerung").Range("E182")
    i = Worksheets("Steuerung").Range("E183")
    RB = Worksheets("Steuerung").Range("E184")
    RP = Worksheets("Steuerung").Range("E185")
    auflage = Worksheets("Steuerung").Range("B184")
    strVerzeichnis = "\\nasbw0\daten\Kalkulationen\"
    strDateiname = Application.GetSaveAsFilename(InitialFileName:=strVerzeichnis & _
    knd & "_" & format & "_" & " F" & F & " I" & i & " RB" & RB & " RP" & RP & "_" & _
    auflage & ".xls", FileFilter:="Microsoft Excel-Arbeitsmappe (*.xls), *.xls")
      Select Case strDateiname
        Case False
          Exit Sub
        Case Else
          ThisWorkbook.SaveAs Filename:=strDateiname
      End Select
End Sub
Sub Druck_Form()
    On Error Resume Next
    UFDrucken.Show
End Sub
Sub version()
    ' Versionsnummer um 1 erhöhen
    '04.01.2013
    Dim v1 As Integer
    v1 = Worksheets("Steuerung").Range("B178")
    v1 = v1 + 1
    Worksheets("Steuerung").Range("B178") = v1
    Worksheets("Steuerung").Range("A178") = Date & "/" & Time
End Sub
Sub checkdate()
    'Datum d. Fehlerprüfung
    '14.01.2009
    Worksheets("Steuerung").Range("B179") = Now
End Sub
'Sub TextBox2_Change()
''Auflage 1 nur Zahleneingabe möglich
'    If Not IsNumeric(TextBox2) And Not TextBox2 = "" Then
'        TextBox2 = Left(TextBox2.Text, Len(TextBox2.Text) - 1)
'    End If
'    TextBox2 = format(Range("C8"), "#,###")
'End Sub
'Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'If IsNumeric(UserForm1.TextBox2) = False Then
'MsgBox "TB verlassen"
'End If
'End Sub
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'If Target.Address = "$C$9" Then Call Formatinput 'Macrostart bei Klick in Zelle
    'If Not Intersect(Target, Range("C32")) Is Nothing Then Call Materialkommentar 'Macrostart bei Maus over, Enter od. Tab
    'If Not Intersect(Target, Range("C9")) Is Nothing Then Call Produktformat
'End Sub
Sub Farbpalette_ausgeben()
    Dim bytIndex As Byte
    Dim bytColumn As Byte
    Dim bytColorIndex As Byte
    For bytColumn = 1 To 4
        For bytIndex = 1 To 14
            bytColorIndex = (14 * (bytColumn - 1)) + bytIndex
            ActiveSheet.Cells((bytIndex + 27), (bytColumn * 2) - 1) _
            .Value = bytColorIndex
            ActiveSheet.Cells((bytIndex + 27), bytColumn * 2) _
            .Interior.ColorIndex = bytColorIndex
        Next bytIndex
    Next bytColumn
End Sub
Sub Farbe_entfernen()
    ActiveSheet.Range("A1:H25").Interior.ColorIndex = xlColorIndexNone
End Sub
Sub Farbe_zuweisen()
    Dim bytColorIndex As Byte
    Dim a, b As Integer
    Dim Auftrag, Kunde As String
    Worksheets("Plantafel").Unprotect "bw"
    Auftrag = Worksheets("Steuerung").Range("C94")
    Kunde = Worksheets("Steuerung").Range("B94")
    bytColorIndex = ActiveSheet.Range("J1")
    If bytColorIndex > 2 And bytColorIndex < 57 Then
        For a = 1 To 22 Step 4
            ActiveSheet.Range(Cells(a, 1), Cells(a, 8)).Interior.ColorIndex = bytColorIndex
        Next
        For a = 4 To 25 Step 4
            For b = 1 To 8
                ActiveSheet.Cells(a, b) = "Auftr.:" & Auftrag & ", " & Kunde & ", C" & bytColorIndex & ", Bem.:"
            Next
        Next
    Else
    MsgBox "Bitte nur Werte zwischen 3 und 57 eingeben!"
    End If
End Sub
Sub Dokumenteigenschaften_Ist()
    'Dokumenteigenschaften auflisten
    '14.01.2009
    Dim i As Long
    Dim a As Integer
    On Error Resume Next
    a = 189
    For i = 1 To ThisWorkbook.BuiltinDocumentProperties.Count
      Worksheets("Steuerung").Cells(a + i, 1) = ThisWorkbook.BuiltinDocumentProperties(i).Name
      Worksheets("Steuerung").Cells(a + i, 2) = ThisWorkbook.BuiltinDocumentProperties(i).Value
    Next i
    'Columns("A:B").AutoFit
End Sub
Sub Dokumenteigenschaften_Soll()
    'Dokumenteigenschaften setzen
    '14.01.2009
    Dim i As Long
    Dim a As Integer
    On Error Resume Next
    Worksheets("Steuerung").Range("C190") = ThisWorkbook.Name
    Worksheets("Steuerung").Range("C218") = ThisWorkbook.FullName
    a = 189
    For i = 1 To ThisWorkbook.BuiltinDocumentProperties.Count
      ThisWorkbook.BuiltinDocumentProperties(i).Value = Worksheets("Steuerung").Cells(a + i, 3)
    Next i
End Sub

