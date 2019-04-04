VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle18"
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
    ' Loeschen d. vorhandenen Eingaben wenn d. Bg. nicht ausgewaehlt wurde
    ' aktivieren wenn Nutzen
    '20190227
    Application.ScreenUpdating = False
    If Worksheets("Eingabe").CheckBox4.Value = True Then 'Schneiden
        'Bg. A
        If Worksheets("SSchneiden").Range("B105") > 0 Then 'Bg. Anzahl
            If Worksheets("SSchneiden").Range("B113") = 1 Then 'Umschlag immer schneiden
                Worksheets("Schneiden").CheckBox1.Value = True
                Worksheets("Schneiden").Range("D3").FormulaLocal = "=SVERWEIS(SSchneiden!B104;SSchneiden!I13:K30;3)"
            Else
                If Worksheets("SSchneiden").Range("B112") = 1 And Worksheets("SSchneiden").Range("B104") > 1 Then 'Schneiden wenn Inhalt ja und Nutzen > 1
                    Worksheets("Schneiden").CheckBox1.Value = True
                    Worksheets("Schneiden").Range("D3").FormulaLocal = "=SVERWEIS(SSchneiden!B104;SSchneiden!I13:K30;3)"
                End If
                'Schneiden wenn Inhalt und Nutzen = 1, manuelles deaktivieren der CheckBox möglich
                If Worksheets("SSchneiden").Range("B112") = 1 And Worksheets("SSchneiden").Range("B104") = 1 Then
                    Worksheets("Schneiden").Range("D3").FormulaLocal = "=SVERWEIS(SSchneiden!B104;SSchneiden!I13:K30;3)"
                End If
            End If
        Else
            Worksheets("Schneiden").Range("D3") = "0" 'Anzahl der Schnitte
            Worksheets("Schneiden").CheckBox1.Value = False
        End If
        'Bg. B
        If Worksheets("Eingabe").CheckBox1.Value = True Then 'Bg. ja/nein
            If Worksheets("SSchneiden").Range("C113") = 1 Then 'Umschlag immer schneiden
                Worksheets("Schneiden").CheckBox2.Value = True
                Worksheets("Schneiden").Range("D10").FormulaLocal = "=SVERWEIS(SSchneiden!C104;SSchneiden!I13:K30;3)"
            Else
                If Worksheets("SSchneiden").Range("C112") = 1 And Worksheets("SSchneiden").Range("C104") > 1 Then 'Schneiden wenn Inhalt ja und Nutzen > 1
                    Worksheets("Schneiden").CheckBox2.Value = True
                    Worksheets("Schneiden").Range("D10").FormulaLocal = "=SVERWEIS(SSchneiden!C104;SSchneiden!I13:K30;3)"
                End If
                'Schneiden wenn Inhalt und Nutzen = 1, manuelles deaktivieren der CheckBox möglich
                If Worksheets("SSchneiden").Range("C112") = 1 And Worksheets("SSchneiden").Range("C104") = 1 Then
                    Worksheets("Schneiden").Range("D10").FormulaLocal = "=SVERWEIS(SSchneiden!C104;SSchneiden!I13:K30;3)"
                End If
            End If
        Else
            Worksheets("Schneiden").Range("D10") = "0" 'Anzahl der Schnitte
            Worksheets("Schneiden").CheckBox2.Value = False
        End If
        'Bg. C
        If Worksheets("Eingabe").CheckBox2.Value = True Then 'Bg. ja/nein
            If Worksheets("SSchneiden").Range("D113") = 1 Then 'Umschlag immer schneiden
                Worksheets("Schneiden").CheckBox3.Value = True
                Worksheets("Schneiden").Range("D17").FormulaLocal = "=SVERWEIS(SSchneiden!D104;SSchneiden!I13:K30;3)"
            Else
                If Worksheets("SSchneiden").Range("D112") = 1 And Worksheets("SSchneiden").Range("D104") > 1 Then 'Schneiden wenn Inhalt ja und Nutzen >= 1
                    Worksheets("Schneiden").CheckBox3.Value = True
                    Worksheets("Schneiden").Range("D17").FormulaLocal = "=SVERWEIS(SSchneiden!D104;SSchneiden!I13:K30;3)"
                End If
                'Schneiden wenn Inhalt und Nutzen = 1, manuelles deaktivieren der CheckBox möglich
                If Worksheets("SSchneiden").Range("D112") = 1 And Worksheets("SSchneiden").Range("D104") = 1 Then
                    Worksheets("Schneiden").Range("D17").FormulaLocal = "=SVERWEIS(SSchneiden!D104;SSchneiden!I13:K30;3)"
                End If
            End If
        Else
            Worksheets("Schneiden").Range("D17") = "0" 'Anzahl der Schnitte
            Worksheets("Schneiden").CheckBox3.Value = False
        End If
        'Bg. D
        If Worksheets("Eingabe").CheckBox3.Value = True Then 'Bg. ja/nein
            If Worksheets("SSchneiden").Range("E113") = 1 Then 'Umschlag immer schneiden
                Worksheets("Schneiden").CheckBox4.Value = True
                Worksheets("Schneiden").Range("D24").FormulaLocal = "=SVERWEIS(SSchneiden!E104;SSchneiden!I13:K30;3)"
            Else
                If Worksheets("SSchneiden").Range("E112") = 1 And Worksheets("SSchneiden").Range("E104") > 1 Then 'Schneiden wenn Inhalt und Nutzen > 1
                    Worksheets("Schneiden").CheckBox4.Value = True
                    Worksheets("Schneiden").Range("D24").FormulaLocal = "=SVERWEIS(SSchneiden!E104;SSchneiden!I13:K30;3)"
                End If
                'Schneiden wenn Inhalt und Nutzen = 1, manuelles deaktivieren der CheckBox möglich
                If Worksheets("SSchneiden").Range("E112") = 1 And Worksheets("SSchneiden").Range("E104") = 1 Then
                    Worksheets("Schneiden").Range("D24").FormulaLocal = "=SVERWEIS(SSchneiden!E104;SSchneiden!I13:K30;3)"
                End If
            End If
        Else
            Worksheets("Schneiden").Range("D24") = "0" 'Anzahl der Schnitte
            Worksheets("Schneiden").CheckBox4.Value = False
        End If
    Else
        Worksheets("Schneiden").Range("D3") = "0"
        Worksheets("Schneiden").CheckBox1.Value = False
        Worksheets("Schneiden").Range("D10") = "0"
        Worksheets("Schneiden").CheckBox2.Value = False
        Worksheets("Schneiden").Range("D17") = "0"
        Worksheets("Schneiden").CheckBox3.Value = False
        Worksheets("Schneiden").Range("D24") = "0"
        Worksheets("Schneiden").CheckBox4.Value = False
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


