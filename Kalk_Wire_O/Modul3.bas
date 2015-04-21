Attribute VB_Name = "Modul3"
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
For i = 1 To Sheets.Count - 1
  With Sheets(i)
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
        
     End If
    Next Sh
  End With
Next i
MsgBox "Fertig Master!" & vbLf & vbLf & ShCounter & " VB Buttonformate exportiert."
End Sub
Sub DebugPrint(s As String)
   Static fso As Object

   If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")
   With fso.OpenTextFile(ThisWorkbook.Path & "\debug.log", 8, True, -1)
      .WriteLine s
      .Close
   End With
End Sub
Sub DebugClear()
   CreateObject("Scripting.FileSystemObject").CreateTextFile ThisWorkbook.Path & "\debug.log", True, True
End Sub

