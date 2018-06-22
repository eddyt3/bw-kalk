VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDatenImport 
   Caption         =   "Daten importieren"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9195
   OleObjectBlob   =   "FrmDatenImport.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "FrmDatenImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Const strErsterEintrag As String = "Datendatei"


Private Sub cmdAbbrechen_Click()
Unload Me
End Sub

Private Sub cmdImport_Click()
Dim objBlatt As Object
Dim varPfad As Variant
Dim strText As String
Dim lngZ As Long
Dim strFormat As String
Dim strTabelle As String, strZelle As String, varInhalt As String, strAktBlatt As String
lngZ = 1
varPfad = lblDatei.Caption
strAktBlatt = ActiveSheet.Name
Open varPfad For Input As #1
    Do While Not EOF(1)
        Line Input #1, strText
            If lngZ > 1 Then
                strTabelle = STRINGG(strText, 1)
                For Each objBlatt In ActiveWorkbook.Sheets
                    If objBlatt.Name = strTabelle Then GoTo GEFUNDEN
                Next
                MsgBox "Das Blatt " & strTabelle & " wurde nicht gefunden, der Vorgang wird abgebrochen.", vbOKOnly + vbExclamation, "Fehler"
                Application.StatusBar = False
                Exit Sub
GEFUNDEN:
                Application.StatusBar = "Daten werden importiert, Blatt " & strTabelle & ", Zelle " & strZelle & ", Inhalt: " & varInhalt
                strZelle = STRINGG(strText, 2)
                varInhalt = STRINGG(strText, 3)
                Sheets(strTabelle).Select
                Sheets(strTabelle).Unprotect "bw"
                strFormat = Range(strZelle).NumberFormat
                Range(strZelle).NumberFormat = "General"
                If Left(varInhalt, 1) = "=" Then
                    Range(strZelle).Formula = varInhalt
                Else
                    Range(strZelle) = varInhalt
                End If
                Sheets(strTabelle).Unprotect "bw"
                Worksheets(strTabelle).Range(strZelle).NumberFormat = strFormat
            End If
            lngZ = lngZ + 1
    Loop
Close
Sheets(strAktBlatt).Select
Application.StatusBar = False
cmdAbbrechen.Caption = "Schlieﬂen"
cmdAbbrechen.SetFocus
MsgBox "Die Daten wurden erfolgreich importiert.", vbOKOnly + vbInformation, "Fertig"
End Sub



Private Sub cmddatei_Click()
Dim objBlatt As Object
Dim varPfad As Variant
Dim strText As String
Dim lngZ As Long
Dim strFormat As String
Dim strTabelle As String
Dim strZelle As String
Dim varInhalt As String
varPfad = Application.GetOpenFilename("Datendateien (*.dtn), *.dtn")
If varPfad = False Then Exit Sub
Open varPfad For Input As #1
    Line Input #1, strText
    If Left(strText, Len(strErsterEintrag)) <> strErsterEintrag Then
        MsgBox "Bei dieser Datei handelt es sich nicht um eine g¸ltige Datendatei. Der Vorgang wurde abgebrochen.", vbOKOnly + vbExclamation, "Fehler"
        Close
        Exit Sub
    End If
Close
lblDatei.Caption = varPfad
lstVorschau.Clear
lngZ = 1
Open varPfad For Input As #1
    Do While Not EOF(1)
        Line Input #1, strText
        If lngZ > 1 Then
            strTabelle = STRINGG(strText, 1)
            For Each objBlatt In ActiveWorkbook.Sheets
                If objBlatt.Name = strTabelle Then GoTo GEFUNDEN
            Next
            MsgBox "Das Blatt " & strTabelle & " wurde nicht gefunden, der Vorgang wird abgebrochen.", vbOKOnly + vbExclamation, "Fehler"
            Close
            Application.StatusBar = False
            Exit Sub
GEFUNDEN:
            strZelle = STRINGG(strText, 2)
            varInhalt = STRINGG(strText, 3)
            lstVorschau.AddItem strTabelle & ", Zelle " & strZelle & ":  " & varInhalt
        Else
            lstVorschau.AddItem strText
        End If
        lngZ = lngZ + 1
    Loop
Close
End Sub
