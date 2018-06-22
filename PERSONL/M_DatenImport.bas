Attribute VB_Name = "M_DatenImport"
'Daten ex- und importieren
'
'Dieses Add-In kopiert die markierten Zellen in eine separate Textdatei und fügt die Daten an den gleichen Positionen wieder ein.
'
'Vorgehen:
'Markieren Sie auf den einzelnen Tabellenblättern die Zellen, deren Inhalte später wieder eingefügt werden sollen.
'Getrennte Zellbereiche können Sie mit gedrückter Strg-Taste anklicken, um diese zu markieren.
'Über den Befehl "Daten exportieren" werden diese Bereiche dann in einer Datei mit der Endung *.dtn kopiert.
'Den Dateinamen und Ordner können Sie beliebig eingeben/wählen. Einzelne aktive Zellen werden dabei nicht berücksichtigt.
'
'Beispiel:
'Auf Tabelle2 ist die Zelle A1 aktiv, auf Tabelle3 ist der Bereich A5 bis A10 markiert.
'Die Zelle A1 auf Tabelle2 wird nicht berücksichtigt, aber alle Zellen von A5 bis A10 auf Tabelle3.
'
'Sollen die Daten wieder importiert werden, wählen Sie den Befehl "Daten importieren".
'Sie können dann den Ordner und die Datei wählen, aus der die Daten importiert werden sollen und sehen dann eine Vorschau.
'Wird dies bestätigt, werden die Daten importiert und dabei alle Inhalte der aktiven Mappe durch die importierten Daten überschrieben.
'
Option Explicit
Private Const strErsterEintrag As String = "Datendatei"
Private Const Trennung = "|||"
Private Const DTrennung = "!!!"

Sub DatenImport__XLS2XLS()
FrmDatenImport.Show
End Sub
Sub DatenExport_XLS2XLS()
'Unterschied: Sheets vorher auf Visible stellen
Dim objBlatt, objZelle As Object
Dim varPfad As Variant
Dim intFrage As Integer
Dim lngZ As Long
Dim strAktBlatt As String

strAktBlatt = ActiveSheet.Name

For Each objBlatt In Sheets
    If Selection.Cells.count > 1 Then GoTo MARKIERT
Next
MsgBox "Auf mindestens einem Blatt müssen mehrere Zellen markiert sein. Es sind keine Zellen markiert, Daten können nicht exportiert werden.", vbOKOnly + vbInformation, "Keine Zellen markiert"
Exit Sub
MARKIERT:

intFrage = MsgBox("Alle Zellinhalte, die sich in Markierungen befinden, werden exportiert. Zellen, die sich nicht in einer Markierung befinden, werden nicht berücksichtigt. Sind Sie sicher, daß Sie die markierten Daten exportieren möchten?", vbYesNo + vbQuestion, "Fortsetzen?")
If intFrage = vbNo Then Exit Sub
varPfad = Application.GetSaveAsFilename("", fileFilter:="Datendateien (*.dtn), *.dtn")
If varPfad = False Then Exit Sub
If Dir(varPfad) <> "" Then
    intFrage = MsgBox("Die Datei " & vbNewLine & varPfad & vbNewLine & "existiert bereits. Soll sie überschrieben werden?", vbYesNo + vbQuestion, "Datei existiert")
    If intFrage = vbNo Then
        Exit Sub
    Else
        Kill varPfad
    End If
End If
Open varPfad For Output As #1
    Print #1, strErsterEintrag & ", " & Date & ", " & Time
    For Each objBlatt In Sheets
        'MsgBox (ActiveSheet.Name)
        ActiveSheet.Visible = True 'bei ausgeblendeten Blättern kommt es sonst zum Abbruch
        ActiveSheet.Unprotect "bw"
        objBlatt.Select
        If Selection.Cells.count > 1 Then
            For Each objZelle In Selection
                If objZelle.HasFormula Then
                    Print #1, SCHREIBE(ActiveSheet.Name, objZelle.Address(False, False), objZelle.Formula)
                Else
                    Print #1, SCHREIBE(ActiveSheet.Name, objZelle.Address(False, False), objZelle)
                End If
            Next
        End If
    Next
Close #1
Sheets(strAktBlatt).Select
End Sub

Function STRINGG(Folge, welcher)
Dim intI As Integer
Dim intZaehler As Integer
Dim intBeginn As Integer, intEnde As Integer
intBeginn = 0
intEnde = 0
intZaehler = 1
For intI = 1 To Len(Folge) + 2
    If Mid(Folge, intI, 3) = Trennung Then
        If intZaehler = welcher Then intBeginn = intI + 3
        If intZaehler = welcher + 1 Then
            intEnde = intI
            Exit For
        End If
        intZaehler = intZaehler + 1
    End If
Next
On Error Resume Next
If intEnde > 0 Then STRINGG = Mid(Folge, intBeginn, intEnde - intBeginn) Else STRINGG = ""
End Function

Function STRINGD(Folge, welcher)
Dim intI As Integer
Dim intZaehler As Integer
Dim intBeginn As Integer, intEnde As Integer
intBeginn = 0
intEnde = 0
intZaehler = 1
For intI = 1 To Len(Folge) + 2
    If Mid(Folge, intI, 3) = DTrennung Then
        If intZaehler = welcher Then intBeginn = intI + 3
        If intZaehler = welcher + 1 Then
            intEnde = intI
            Exit For
        End If
        intZaehler = intZaehler + 1
    End If
Next
If intEnde > 0 Then STRINGD = Mid(Folge, intBeginn, intEnde - intBeginn) Else STRINGD = ""
End Function

Function SCHREIBE(strBlattname As String, strZelle As String, varInhalt As Variant) As String
SCHREIBE = "|||" & strBlattname & "|||" & strZelle & "|||" & varInhalt & "|||"
End Function

Function GetUserName() As String
GetUserName = Environ$("UserName")
End Function


