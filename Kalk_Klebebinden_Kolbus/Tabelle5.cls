VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Sub Produktformat_Kleben()
    '
    ' Formatprüfung Makro
    ' Makro von Enrico Dargel
    On Error Resume Next
    Dim Dicke As String
    If Range("SKleben!E2") = FALSCH Then
        MsgBox ("Fehlerhafte Formateingabe! Bitte nur Werte mit Komma eingeben." & vbCrLf & vbCrLf & _
        "Beispiel: 21 x 29,7")
        'FFormat = "Fehlerhafte Formateingabe! Bitte nur Werte mit Komma eingeben."
        Else
        'FFormat = ""
    End If
End Sub '
Sub kleben()
    Application.ScreenUpdating = False
    If Worksheets("SEingabe").Range("B34") > 0 Then
        Worksheets("Kleben").Visible = True
        Worksheets("Kleben").CheckBox1.Value = True
        Else
        Worksheets("Kleben").Visible = False
        Worksheets("Kleben").CheckBox1.Value = False
        Worksheets("Kleben").CheckBox2.Value = False

    End If
    Call KlebenAnzeige
    Application.ScreenUpdating = True
End Sub
Sub KlebenAnzeige()
    'Anzeige d. Klebearten auf der Eingabeseite
    Dim blVN, blBgankl, blKleben As Boolean
    Dim strVN, strBgankl, anzeige As String
    
    blVN = Worksheets("Kleben").CheckBox1.Value
    blBgankl = Worksheets("Kleben").CheckBox2.Value
    blKleben = Worksheets("Eingabe").CheckBox9.Value
    
    If blKleben = True Then
        If blVN = True Then
            strVN = "  V/N kleb., "
            Else: strVN = "  "
        End If
        If blBgankl = True Then
            strBgankl = "Bg. ankl."
            Else: strBgankl = " "
        End If
        anzeige = strVN & strBgankl
        Else: anzeige = " "
    End If
    Worksheets("SKleben").Range("B3") = anzeige
End Sub
Sub Bogen_erstellen_brutto()
'unused
'volle Ausnutzung des Bogens
On Error Resume Next
Dim intHline, intVline, intHnumber, intVnumber As Integer 'Anzahl d. Linien u. Zählervar.
Dim intBgH, intBgW, intBgHoffset, intBgVoffset As Integer 'Bogenmaße u. Offsets
Dim intHpixel, intVpixel, intHpixelSum, intVpixelSum As Integer 'Versatzwert, Versatzsumme
Dim rngBereich As Range, shpBogen As Shape ' Bereichauswahl f. Gruppierung
intBgW = Worksheets("SKleben").Range("J49")
intBgH = Worksheets("SKleben").Range("J50")
intBgHoffset = Worksheets("SKleben").Range("J51")
intBgVoffset = Worksheets("SKleben").Range("J52")
' Anzahl d. Linien ermitteln
If Worksheets("SKleben").Range("F49") = 1 Then
    intHline = Worksheets("SKleben").Range("F51") - 1
    intVline = Worksheets("SKleben").Range("F52") - 1
    intHpixel = intBgW / Worksheets("SKleben").Range("F51")
    intVpixel = intBgH / Worksheets("SKleben").Range("F52")
    Else
    intHline = Worksheets("SKleben").Range("G51") - 1
    intVline = Worksheets("SKleben").Range("G52") - 1
    intHpixel = intBgW / Worksheets("SKleben").Range("51")
    intVpixel = intBgH / Worksheets("SKleben").Range("G52")
End If
' Bogen zeichnen
    Worksheets("Kleben").Shapes.AddShape Type:=msoShapeRectangle, Left:=intBgHoffset, _
    Top:=intBgVoffset, Width:=intBgW, Height:=intBgH
If intHline > 0 Then
    intHpixelSum = intHpixel
    For intHnumber = 1 To intHline
        Worksheets("Kleben").Shapes.AddShape Type:=msoShapeRectangle, Left:=(intBgHoffset + intHpixelSum), _
        Top:=intBgVoffset, Width:=0.5, Height:=intBgH
        intHpixelSum = intHpixelSum + intHpixel
    Next
End If
If intVline > 0 Then
    intVpixelSum = intVpixel
    For intVnumber = 1 To intVline
        Worksheets("Kleben").Shapes.AddShape Type:=msoShapeRectangle, Left:=intBgHoffset, _
        Top:=(intBgVoffset + intVpixelSum), Width:=intBgW, Height:=0.5
        intVpixelSum = intVpixelSum + intVpixel
    Next
End If
End Sub
Sub Bogen_erstellen_netto()
'Originalnutzenverteilung
On Error Resume Next
Dim intHline, intVline, intHnumber, intVnumber As Integer 'Anzahl d. Linien u. Zählervar.
Dim intBgH, intBgW, intBgHoffset, intBgVoffset As Integer 'Bogenmaße u. Offsets
Dim intHpixel, intVpixel, intHpixelSum, intVpixelSum As Integer 'Versatzwert, Versatzsumme
' alle Autoformen löschen
'If Worksheets("Kleben").Shapes(intIndex).Type = msoShapeRectangle Then
'    Worksheets("Kleben").Shapes(intIndex).Delete
'End If
If Worksheets("Kleben").CheckBox1.Value = True And Worksheets("SKleben").Range("C33") = 1 Then
    Worksheets("Kleben").Unprotect "bw"
    Dim shpShape As Shape
        For Each shpShape In ActiveSheet.Shapes
            If shpShape.Type = msoAutoShape Then
                shpShape.Delete
            End If
        Next shpShape
    intBgW = Worksheets("SKleben").Range("J49")
    intBgH = Worksheets("SKleben").Range("J50")
    intBgHoffset = Worksheets("SKleben").Range("J51")
    intBgVoffset = Worksheets("SKleben").Range("J52")
    If Worksheets("SKleben").Range("G46") = 0 Then 'alternative Laufrichtung prüfen
            ' Anzahl d. Linien ermitteln bei optimaler Nutzenverteilung
            If Worksheets("SKleben").Range("F49") = 1 Then
                intHline = Worksheets("SKleben").Range("F51")
                intVline = Worksheets("SKleben").Range("F52")
                intHpixel = Worksheets("SKleben").Range("K56")
                intVpixel = Worksheets("SKleben").Range("K57")
                Else
                intHline = Worksheets("SKleben").Range("G51")
                intVline = Worksheets("SKleben").Range("G52")
                intHpixel = Worksheets("SKleben").Range("K57")
                intVpixel = Worksheets("SKleben").Range("K56")
            End If
        Else
        ' Anzahl d. Linien ermitteln bei alternativer Laufrichtung
            If Worksheets("SKleben").Range("F93") = 1 Then
                intHline = Worksheets("SKleben").Range("G51")
                intVline = Worksheets("SKleben").Range("G52")
                intHpixel = Worksheets("SKleben").Range("K57")
                intVpixel = Worksheets("SKleben").Range("K56")
                Else
                intHline = Worksheets("SKleben").Range("F51")
                intVline = Worksheets("SKleben").Range("F52")
                intHpixel = Worksheets("SKleben").Range("K56")
                intVpixel = Worksheets("SKleben").Range("K57")
            End If
    End If
    ' Bogen zeichnen
        Set shpShape = Worksheets("Kleben").Shapes.AddShape(Type:=msoShapeRectangle, Left:=intBgHoffset, _
        Top:=intBgVoffset, Width:=intBgW, Height:=intBgH)
        With shpShape.Fill
            .ForeColor.RGB = RGB(255, 153, 0)
            .Solid
        End With
        With shpShape.Shadow
            .Type = msoShadow2
        End With
    'Nutzen zeichnen
    If intHline > 0 Then
        intHpixelSum = intHpixel
        For intHnumber = 1 To intHline
            Worksheets("Kleben").Shapes.AddShape Type:=msoShapeRectangle, Left:=(intBgHoffset + intHpixelSum), _
            Top:=intBgVoffset, Width:=0.2, Height:=(intVpixel * intVline)
            intHpixelSum = intHpixelSum + intHpixel
        Next
        Else
        MsgBox "Achtung, das Bogenformat f. Vor-, Nachsatz ist zu klein!"
    End If
    If intVline > 0 Then
        intVpixelSum = intVpixel
        For intVnumber = 1 To intVline
            Worksheets("Kleben").Shapes.AddShape Type:=msoShapeRectangle, Left:=intBgHoffset, _
            Top:=(intBgVoffset + intVpixelSum), Width:=(intHpixel * intHline), Height:=0.2
            intVpixelSum = intVpixelSum + intVpixel
        Next
        Else
        MsgBox "Achtung, das Bogenformat f. Vor-, Nachsatz ist zu klein!"
    End If
    'Auswahl aller Shapes im Bereich u. Gruppierung
    '   Set rngBereich = Range("H14:L27")
    '   For Each shpBogen In ActiveSheet.Shapes
    '      If shpBogen.Top > rngBereich.Top And shpBogen.Height < rngBereich.Height Then
    '         If shpBogen.Left > rngBereich.Left And shpBogen.Width < rngBereich.Width Then
    '            shpBogen.Group
    '         End If
    '      End If
    '   Next
    Worksheets("Kleben").Protect "bw"
    Else
End If
End Sub

