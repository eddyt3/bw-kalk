Attribute VB_Name = "Modul6"
'Modul für Dateioperationen

Option Explicit
'Dateien suchen und umbenennen
'
Sub findAndRename()
  Dim objFiles() As Object, lngRet As Long, lngIndex As Long
  Dim strNewName As String
  
  lngRet = FileSearchINFO(objFiles, "E:\Temp", "*", True) 'Pfad anpassen!
  
  If lngRet > 0 Then
    For lngIndex = 0 To lngRet - 1
      If objFiles(lngIndex).Name Like "*_?#########v1.*" Then
        strNewName = Left(objFiles(lngIndex).Name, InStr(1, objFiles(lngIndex).Name, "_#") - 1) & _
          Mid(objFiles(lngIndex).Name, InStrRev(objFiles(lngIndex).Name, "."))
        
        Name objFiles(lngIndex) As objFiles(lngIndex).ParentFolder.Path & "\" & strNewName
      End If
    Next
  End If
  
End Sub

Private Function FileSearchINFO(ByRef Files() As Object, ByVal InitialPath As String, Optional ByVal FileName As String = "*", _
    Optional ByVal SubFolders As Boolean = False) As Long
  
  '# PARAMETERINFO:
  '# Files: Datenfeld zur Ausgabe der Suchergebnisse
  '# InitialPath: String der das zu durchsuchende Verzeichnis angibt
  '# FileName: String der den gesuchten Dateityp oder Dateinamen enthält (Optional, Standard="*.*" findet alle Dateien)
  '# Beispiele: "*.txt" - Findet alle Textdateien
  '# "*name*" - Findet alle Dateien mit "name" im Dateinamen
  '# "*.avi;*.mpg" - Findet .avi und .mpg Dateien (Dateitypen mit ; trennen)
  '# SubFolders: Boolean gibt an, ob Unterordner durchsucht werden sollen (Optional, Standard=False)
  
  
  Dim fobjFSO As Object, ffsoFolder As Object, ffsoSubFolder As Object, ffsoFile As Object
  Dim intC As Integer, varFiles As Variant
  
  Set fobjFSO = CreateObject("Scripting.FileSystemObject")
  
  Set ffsoFolder = fobjFSO.GetFolder(InitialPath)
  
  On Error GoTo ErrExit
  
  If InStr(1, FileName, ";") > 0 Then
    varFiles = Split(FileName, ";")
  Else
    ReDim varFiles(0)
    varFiles(0) = FileName
  End If
  For Each ffsoFile In ffsoFolder.Files
    If Not ffsoFile Is Nothing Then
      For intC = 0 To UBound(varFiles)
        If LCase(fobjFSO.GetFileName(ffsoFile)) Like LCase(varFiles(intC)) Then
          If IsArray(Files) Then
            ReDim Preserve Files(UBound(Files) + 1)
          Else
            ReDim Files(0)
          End If
          Set Files(UBound(Files)) = ffsoFile
          Exit For
        End If
      Next
    End If
  Next
  
  If SubFolders Then
    For Each ffsoSubFolder In ffsoFolder.SubFolders
      FileSearchINFO Files, ffsoSubFolder, FileName, SubFolders
    Next
  End If
  
  If IsArray(Files) Then FileSearchINFO = UBound(Files) + 1
ErrExit:
  Set fobjFSO = Nothing
  Set ffsoFolder = Nothing
End Function
