Attribute VB_Name = "Modul2"
Option Explicit
Sub Speichern()
Attribute Speichern.VB_ProcData.VB_Invoke_Func = " \n14"
Dim knd, format, produkt, auflage As String
Dim strVerzeichnis, strDateiname As String
knd = Worksheets("Steuerung").Range("B181")
format = Worksheets("Steuerung").Range("B182")
produkt = Worksheets("Steuerung").Range("B183")
auflage = Worksheets("Steuerung").Range("B184")
strVerzeichnis = "\\192.168.100.1\daten\Kalkulationen\"
'strDateiname = Application.GetSaveAsFilename("Test", FileFilter:="Microsoft Excel-Arbeitsmappe (*.xls), *.xls")
strDateiname = Application.GetSaveAsFilename(InitialFileName:=strVerzeichnis & _
knd & "_" & format & "_" & produkt & "_" & auflage & ".xls", _
FileFilter:="Microsoft Excel-Arbeitsmappe (*.xls), *.xls")
  Select Case strDateiname
    Case False
      Exit Sub
    Case Else
      ThisWorkbook.SaveAs Filename:=strDateiname
  End Select
End Sub
Sub Druck_Form()
Attribute Druck_Form.VB_ProcData.VB_Invoke_Func = " \n14"
On Error Resume Next
UFDrucken.Show
End Sub
Sub version()
Attribute version.VB_ProcData.VB_Invoke_Func = " \n14"
    ' Versionsnummer um 1 erhöhen
    '04.01.2013
    Dim v1 As Integer
    v1 = Worksheets("Steuerung").Range("B178")
    v1 = v1 + 1
    Worksheets("Steuerung").Range("B178") = v1
    Worksheets("Steuerung").Range("A178") = Date & "/" & Time
End Sub
Sub checkdate()
Attribute checkdate.VB_ProcData.VB_Invoke_Func = " \n14"
'
'Datum d. Fehlerprüfung
'
Worksheets("Steuerung").Range("B179") = Now
End Sub
Sub Farbpalette_ausgeben()
Attribute Farbpalette_ausgeben.VB_ProcData.VB_Invoke_Func = " \n14"
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
Attribute Farbe_entfernen.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveSheet.Range("A1:H25").Interior.ColorIndex = xlColorIndexNone
End Sub
Sub Farbe_zuweisen()
Attribute Farbe_zuweisen.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim bytColorIndex As Byte
    Dim a, b As Integer
    Dim Auftrag, Kunde As String
    Worksheets("Plantafel").Unprotect "bw"
    Auftrag = Worksheets("Steuerung").Range("C181")
    Kunde = Worksheets("Steuerung").Range("B181")
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
