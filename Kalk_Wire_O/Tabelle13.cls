VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Sub Standardwerte_Schneiden()
    '20140623
    Application.ScreenUpdating = False
    Worksheets("SSchneiden").Range("G5:G10").Copy Destination:=Worksheets("Schneiden").Range("N3")
    Worksheets("SSchneiden").Range("F4").Copy Destination:=Worksheets("Schneiden").Range("M2")
    Application.ScreenUpdating = True
End Sub
Sub Schneidbogen()
    '
    ' Loeschen d. vorhandenen Eingaben wenn d. Bg. nicht ausgewaehlt wurde
    '
    '20130218
    On Error Resume Next
    Application.ScreenUpdating = False
    Worksheets("Schneiden").Unprotect "bw"
    If Worksheets("Steuerung").Range("D49") = 2 Or Worksheets("Steuerung").Range("D49") = 4 Then
        Worksheets("Schneiden").Range("D3") = "0"
        Else
        Worksheets("Schneiden").Range("D3").FormulaLocal = "=SVERWEIS(SSchneiden!B104;SSchneiden!I13:K30;3)"
    End If
    If Worksheets("Steuerung").Range("E49") = 2 Or Worksheets("Steuerung").Range("E49") = 4 Then
        Worksheets("Schneiden").Range("D10") = "0"
        Else
        Worksheets("Schneiden").Range("D10").FormulaLocal = "=SVERWEIS(SSchneiden!C104;SSchneiden!I13:K30;3)"
    End If
    If Worksheets("Steuerung").Range("F49") = 2 Or Worksheets("Steuerung").Range("F49") = 4 Then
        Worksheets("Schneiden").Range("D17") = "0"
        Else
        Worksheets("Schneiden").Range("D17").FormulaLocal = "=SVERWEIS(SSchneiden!D104;SSchneiden!I13:K30;3)"
    End If
    If Worksheets("Steuerung").Range("G49") = 2 Or Worksheets("Steuerung").Range("G49") = 4 Then
        Worksheets("Schneiden").Range("D24") = "0"
        Else
        Worksheets("Schneiden").Range("D24").FormulaLocal = "=SVERWEIS(SSchneiden!E104;SSchneiden!I13:K30;3)"
    End If
    If Worksheets("Steuerung").Range("H49") = 2 Or Worksheets("Steuerung").Range("H49") = 4 Then
        Worksheets("Schneiden").Range("D31") = "0"
        Else
        Worksheets("Schneiden").Range("D31").FormulaLocal = "=SVERWEIS(SSchneiden!F104;SSchneiden!I13:K30;3)"
    End If
    If Worksheets("Eingabe").CheckBox1 = False Then
        Worksheets("Schneiden").Range("D3") = "0"
        Worksheets("Schneiden").Range("D10") = "0"
        Worksheets("Schneiden").Range("D17") = "0"
        Worksheets("Schneiden").Range("D24") = "0"
        Worksheets("Schneiden").Range("D31") = "0"
        Worksheets("Schneiden").Range("D38") = "0"
     End If
    Worksheets("Schneiden").Protect "bw"
    Application.ScreenUpdating = True
End Sub
Sub SchneidenAnzeige()
    'Anzeige d. zu schneidenden Bogen auf der Eingabeseite
    '20130218
    Application.ScreenUpdating = False
    Dim rngBgA, rngBgB, rngBgC, rngBgD, rngBgE, rngSum As String
    Dim BgA, BgB, BgC, BgD, anzeige As String
    
    rngBgA = Worksheets("SSchneiden").Range("D4")
    rngBgB = Worksheets("SSchneiden").Range("D5")
    rngBgC = Worksheets("SSchneiden").Range("D6")
    rngBgD = Worksheets("SSchneiden").Range("D7")
    rngBgE = Worksheets("SSchneiden").Range("D8")
    rngSum = Worksheets("SSchneiden").Range("D9")
    
    If rngSum = True Then
        If rngBgA = True Then
            BgA = " F,"
            Else: BgA = ""
        End If
        If rngBgB = True Then
            BgB = "D,"
            Else: BgB = ","
        End If
        If rngBgC = True Then
            BgC = "I,"
            Else: BgC = ","
        End If
        If rngBgD = True Then
            BgD = "Rb"
            Else: BgD = ""
        End If
            If rngBgE = True Then
            BgE = "Rp"
            Else: BgE = ""
        End If
    
        anzeige = "Bg.:" & BgA & BgB & BgC & BgD & BgE
        Else: anzeige = " "
    End If
    Worksheets("SSchneiden").Range("D10") = anzeige
    Application.ScreenUpdating = True
End Sub
