VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Thickness()
    'Eingabe alternative Pappst‰rke Buchdecke
    Application.ScreenUpdating = False
    ' Auswahl alternative Materialst‰rke
        If Worksheets("Eingabe").ComboBox3.ListIndex = 6 Then
            Worksheets("Eingabe").Label3.Visible = True
        Else
            Worksheets("Eingabe").Label3.Visible = False
        End If
        If Worksheets("Eingabe").ComboBox3.ListIndex <> 6 And Worksheets("SEingabe").Range("H26") = 1 Then
            MsgBox ("Standardformate nur in 2mm Pappst‰rke vorr‰tig!")
            Worksheets("Eingabe").Range("F9") = "Sonderformat"
        End If
        If Worksheets("Eingabe").ComboBox3.ListIndex = 6 And Worksheets("SEingabe").Range("H26") = 1 Then
            Worksheets("Eingabe").Range("F9") = ""
        End If
    Application.ScreenUpdating = True
End Sub
Sub Beschriftung()
    'Beschriftung der Skizze
    '20151127
    Dim intRuecken, intPappeB, intPappeL, intBezugB, intBezugL, intDeckeB, intDeckeL, intFalz As String
    intRuecken = Worksheets("SBuchdecke").Range("K3")
    intPappeB = Worksheets("SBuchdecke").Range("K4")
    intPappeL = Worksheets("SBuchdecke").Range("M4")
    intBezugB = Worksheets("SBuchdecke").Range("K5")
    intBezugL = Worksheets("SBuchdecke").Range("M5")
    intDeckeB = Worksheets("SBuchdecke").Range("K6")
    intDeckeL = Worksheets("SBuchdecke").Range("M6")
    intFalz = Worksheets("SBuchdecke").Range("K7")
    Worksheets("Buchdecke").Label1.Caption = intRuecken & Chr(10) & "mm"
    Worksheets("Buchdecke").Label2.Caption = intPappeB & "mm"
    Worksheets("Buchdecke").Label3.Caption = intPappeL & "mm"
    Worksheets("Buchdecke").Label4.Caption = intBezugB & "mm"
    Worksheets("Buchdecke").Label5.Caption = intBezugL & Chr(10) & "mm"
    Worksheets("Buchdecke").Label6.Caption = intDeckeB & "mm"
    Worksheets("Buchdecke").Label7.Caption = intDeckeL & Chr(10) & "mm"
    Worksheets("Buchdecke").Label8.Caption = intFalz & Chr(10) & "mm"
    Worksheets("Buchdecke").Label9.Caption = intFalz & Chr(10) & "mm"
End Sub
Sub HardcoverVisual()
    '20150504
    '
    Dim intTop, intLeft, intUberzugB, intUberzugH, intRuckenB, intRuckenH, intPappeB, intPappeH, intOffeneDeckeB, intOffeneDeckeH As Integer
    intTop = Round(Worksheets("SBuchdecke").Range("N14"), 0)
    intLeft = Round(Worksheets("SBuchdecke").Range("N15"), 0)
    intUberzugB = Round(Worksheets("SBuchdecke").Range("N5"), 0)
    intUberzugH = Round(Worksheets("SBuchdecke").Range("O5"), 0)
    intRuckenB = Round(Worksheets("SBuchdecke").Range("N3"), 0)
    intRuckenH = Round(Worksheets("SBuchdecke").Range("O3"), 0)
    intPappeB = Round(Worksheets("SBuchdecke").Range("N4"), 0)
    intPappeH = Round(Worksheets("SBuchdecke").Range("O4"), 0)
    intOffeneDeckeB = Round(Worksheets("SBuchdecke").Range("N6"), 0)
    intOffeneDeckeH = Round(Worksheets("SBuchdecke").Range("O6"), 0)
    intEinbrennfalz = Round(Worksheets("SBuchdecke").Range("N7"), 0)
    Worksheets("Buchdecke").Shapes("Rectangle 40").Select '‹berzug
    With Selection
        .Top = intTop
        .Left = intLeft
        .Height = intUberzugH
        .Width = intUberzugB
    End With
    Worksheets("Buchdecke").Shapes("Rectangle 43").Select 'R¸cken
    With Selection
        .Top = intTop + ((intUberzugH - intRuckenH) / 2)
        .Left = (intUberzugB / 2 + intLeft) - intRuckenB / 2
        .Height = intRuckenH
        .Width = intRuckenB
    End With
    Worksheets("Buchdecke").Shapes("Rectangle 41").Select 'Pappe links
    With Selection
        .Top = intTop + ((intUberzugH - intRuckenH) / 2)
        .Left = (intUberzugB - intOffeneDeckeB) / 2 + intLeft
        .Height = intPappeH
        .Width = intPappeB
    End With
    Worksheets("Buchdecke").Shapes("Rectangle 42").Select 'Pappe rechts
    With Selection
        .Top = intTop + ((intUberzugH - intRuckenH) / 2)
        .Left = intLeft + intUberzugB / 2 + intRuckenB / 2 + intEinbrennfalz
        .Height = intPappeH
        .Width = intPappeB
    End With
    Worksheets("Buchdecke").Shapes("Line 49").Select '‹berzuglinie H
    With Selection
        .Top = intTop - 9
        .Left = intLeft
        .Width = intUberzugB
    End With
     Worksheets("Buchdecke").Shapes("Label4").Select '‹berzugmaﬂ H
    With Selection
        .Top = intTop - 20
        .Left = (intUberzugB / 2 + intLeft) - intRuckenB / 2
    End With
    Worksheets("Buchdecke").Shapes("Line 50").Select '‹berzuglinie V
    With Selection
        .Top = intTop
        .Left = intLeft + intUberzugB + 10
        .Height = intUberzugH
    End With
     Worksheets("Buchdecke").Shapes("Label5").Select '‹berzugmaﬂ V
    With Selection
        .Top = intTop + intUberzugH / 2
        .Left = intLeft + intUberzugB + 15
    End With
   Worksheets("Buchdecke").Shapes("Line 48").Select 'R¸ckenlinie
    With Selection
        .Top = intTop + ((intUberzugH - intRuckenH) / 2) + 20
        .Left = (intUberzugB / 2 + intLeft) - intRuckenB / 2
        .Width = intRuckenB
    End With
    Worksheets("Buchdecke").Shapes("Label1").Select 'R¸ckenmaﬂ
    With Selection
        .Top = intTop + ((intUberzugH - intRuckenH) / 2) + 25
        .Left = (intUberzugB / 2 + intLeft) - intRuckenB / 2 + 1
    End With
    Worksheets("Buchdecke").Shapes("Label8").Select 'Einbrennfalz
    With Selection
        .Top = intTop + ((intUberzugH - intRuckenH) / 2) + 100
        .Left = (intUberzugB / 2 + intLeft) - intRuckenB - Einbrennfalz
    End With
    Worksheets("Buchdecke").Shapes("Label9").Select 'Einbrennfalz
    With Selection
        .Top = intTop + ((intUberzugH - intRuckenH) / 2) + 100
        .Left = (intUberzugB / 2 + intLeft) + intRuckenB * 0.5 + 1
    End With
   Worksheets("Buchdecke").Shapes("Line 57").Select 'offene Deckenlinie H
    With Selection
        .Top = intTop + ((intUberzugH - intRuckenH) / 2) + intOffeneDeckeH - 20
        .Left = (intUberzugB - intOffeneDeckeB) / 2 + intLeft
        .Width = intOffeneDeckeB
    End With
    Worksheets("Buchdecke").Shapes("Label6").Select 'offenes Deckennmaﬂ H
    With Selection
        .Top = intTop + ((intUberzugH - intRuckenH) / 2) + intOffeneDeckeH - 33
        .Left = intLeft + intOffeneDeckeB * 0.8
    End With
   Worksheets("Buchdecke").Shapes("Line 44").Select 'offene Deckenlinie V
    With Selection
        .Top = intTop + ((intUberzugH - intRuckenH) / 2)
        .Left = intLeft + intOffeneDeckeB * 0.9
        .Height = intOffeneDeckeH
    End With
    Worksheets("Buchdecke").Shapes("Label7").Select 'offenes Deckennmaﬂ V
    With Selection
        .Top = intTop + intUberzugH * 0.8
        .Left = intLeft + intOffeneDeckeB * 0.9 + 1
    End With
   Worksheets("Buchdecke").Shapes("Line 45").Select 'Papplinie H
    With Selection
        .Top = intTop + intUberzugH * 0.7
        .Left = (intUberzugB - intOffeneDeckeB) / 2 + intLeft
        .Width = intPappeB
    End With
    Worksheets("Buchdecke").Shapes("Label2").Select 'Pappmaﬂ H
    With Selection
        .Top = intTop + intUberzugH * 0.7 - 11
        .Left = intLeft + intPappeB * 0.4
    End With
   Worksheets("Buchdecke").Shapes("Line 46").Select 'Papplinie V
    With Selection
        .Top = intTop + (intUberzugH - intOffeneDeckeH) / 2
        .Left = intLeft + (intUberzugB - intOffeneDeckeB) / 2 + intPappeB * 0.5
        .Height = intPappeH
    End With
    Worksheets("Buchdecke").Shapes("Label3").Select 'Pappmaﬂ V
    With Selection
        .Top = intTop + intUberzugH * 0.62
        .Left = intLeft + intPappeB * 0.5 + 30
    End With
    Worksheets("Buchdecke").Range("D37").Select
End Sub
