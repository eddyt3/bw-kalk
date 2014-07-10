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
Sub Standertwerte_Schneiden()
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
    'Anzeige d. zu schneidenden Bogen auf der Eingabeseite
    '20130218
    Application.ScreenUpdating = False
    If Worksheets("SSchneiden").Range("B103") < 1 Then
       Worksheets("Schneiden").Range("D3") = "0"
       Else
       Worksheets("Schneiden").Range("D3").FormulaLocal = "=SVERWEIS(SSchneiden!B104;SSchneiden!I13:K30;3)"
    End If
    If Worksheets("SSchneiden").Range("C103") < 1 Then
       Worksheets("Schneiden").Range("D10") = "0"
       Else
       Worksheets("Schneiden").Range("D10").FormulaLocal = "=SVERWEIS(SSchneiden!C104;SSchneiden!I13:K30;3)"
    End If
    If Worksheets("SSchneiden").Range("D103") < 1 Then
       Worksheets("Schneiden").Range("D17") = "0"
       Else
       Worksheets("Schneiden").Range("D17").FormulaLocal = "=SVERWEIS(SSchneiden!D104;SSchneiden!I13:K30;3)"
    End If
    If Worksheets("SSchneiden").Range("E103") < 1 Then
       Worksheets("Schneiden").Range("D24") = "0"
       Else
       Worksheets("Schneiden").Range("D24").FormulaLocal = "=SVERWEIS(SSchneiden!E104;SSchneiden!I13:K30;3)"
    End If
    If Worksheets("SSchneiden").Range("F111") < 1 Then
       Worksheets("Schneiden").Range("D3") = "0"
       Worksheets("Schneiden").Range("D10") = "0"
       Worksheets("Schneiden").Range("D17") = "0"
       Worksheets("Schneiden").Range("D24") = "0"
    End If
    Application.ScreenUpdating = True
End Sub
Sub SchneidenAnzeige()
    'Anzeige d. zu schneidenden Bogen auf der Eingabeseite
    '20130218
    Application.ScreenUpdating = False
    Dim rngBgA, rngBgB, rngBgC, rngBgD, rngSum As String
    Dim BgA, BgB, BgC, BgD, anzeige As String
    
    rngBgA = Worksheets("SSchneiden").Range("D4")
    rngBgB = Worksheets("SSchneiden").Range("D5")
    rngBgC = Worksheets("SSchneiden").Range("D6")
    rngBgD = Worksheets("SSchneiden").Range("D7")
    rngSum = Worksheets("SSchneiden").Range("D8")
    
    If rngSum = True Then
        If rngBgA = True Then
            BgA = " A,"
            Else: BgA = ""
        End If
        If rngBgB = True Then
            BgB = "B,"
            Else: BgB = ","
        End If
        If rngBgC = True Then
            BgC = "C,"
            Else: BgC = ","
        End If
        If rngBgD = True Then
            BgD = "D"
            Else: BgD = ""
        End If
        anzeige = "Bg.:" & BgA & BgB & BgC & BgD
        Else: anzeige = " "
    End If
    Worksheets("SSchneiden").Range("D9") = anzeige
    Application.ScreenUpdating = True
End Sub

