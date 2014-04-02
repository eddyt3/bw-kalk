Attribute VB_Name = "Modul1"
Public FNutzen, FFormat, FSeitenanz, FSeitenBgMin, FSeitenBgMax, FBgUMax As String 'Fehlervariabeln
Sub Nutzenauswertung()
Attribute Nutzenauswertung.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Überprüfung der Angaben Seitenzahl, Bogenzahl, Seiten pro Bogen
On Error Resume Next
If Range("Steuerung!L7") > 0 Then
Dim BogenA, BogenB, BogenC, BogenD As String
    If Range("Steuerung!L3") > 0 Then
        MsgBox ("Fehlerhafte Eingabe(n) bei Bogen A! " & vbCrLf & vbCrLf & _
        "Bitte 'Seitenzahl', 'Nutzen/Druckbogen', 'Buchbindebogen' u. 'Seiten/Buchbindebogen' kontrollieren.")
        BogenA = " A,"
        Else
        BogenA = ""
    End If
    If Range("Steuerung!L4") > 0 Then
        MsgBox ("Fehlerhafte Eingabe(n) bei Bogen B! " & vbCrLf & vbCrLf & _
        "Bitte 'Seitenzahl', 'Nutzen/Druckbogen', 'Buchbindebogen' u. 'Seiten/Buchbindebogen' kontrollieren.")
        BogenB = " B,"
        Else
        BogenB = ""
    End If
    If Range("Steuerung!L5") > 0 Then
        MsgBox ("Fehlerhafte Eingabe(n) bei Bogen C! " & vbCrLf & vbCrLf & _
        "Bitte 'Seitenzahl', 'Nutzen/Druckbogen', 'Buchbindebogen' u. 'Seiten/Buchbindebogen' kontrollieren.")
        BogenC = " C,"
        Else
        BogenC = ""
    End If
    If Range("Steuerung!L6") > 0 Then
        MsgBox ("Fehlerhafte Eingabe(n) bei Bogen D! " & vbCrLf & vbCrLf & _
        "Bitte 'Seitenzahl', 'Nutzen/Druckbogen', 'Buchbindebogen' u. 'Seiten/Buchbindebogen' kontrollieren.")
        BogenD = " D"
        Else
        BogenD = ""
    End If
FNutzen = "Fehlerhafte Seitenzahl, Bogenzahl od. Seiten/Bogen bei Bogen: " & BogenA & BogenB & BogenC & BogenD & "."
Else
FNutzen = ""
End If
End Sub
Sub Produkt()
Attribute Produkt.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Anzeigen d. Produktangaben
'
On Error Resume Next
Dim format, Gewicht, Dicke As String
    Worksheets("Verpacken").Unprotect "bw"
    format = Worksheets("Eingabe").CommandButton2.Caption
    Dicke = Range("Eingabe!C48")
    Gewicht = Range("Eingabe!C49")
    Worksheets("Verpacken").Label1.Caption = "Produkt:" & vbLf & "======" & vbLf & vbLf & "Format: " & vbLf & format _
    & vbLf & vbLf & "Stärke: " & vbLf & Dicke & " mm" & vbLf & vbLf & "Gewicht: " & vbLf & Gewicht & " g"
    Worksheets("Verpacken").Protect "bw"
End Sub

