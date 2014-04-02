VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'13.01.2009
Option Explicit
Sub draws_delete()
    'alle Zeichnungen löschen
    On Error Resume Next
    Worksheets("Verpacken").Unprotect "bw"
    Dim shpShape As Shape
        For Each shpShape In ActiveSheet.Shapes
            If shpShape.Type = msoAutoShape Then
                shpShape.Delete
            End If
        Next shpShape
End Sub
Sub draw_all()
    'Alles zeichnen
    '20090820
    On Error Resume Next
    Call draws_delete
    Call VE_erstellen
    Call VE_einschweissen
    Call Palettenlage_erstellen
    Worksheets("Verpacken").Protect "bw"
End Sub
Sub Palettenlage_erstellen()
    'Zeichnen der optimalen Exemplarverteilung auf d. Palette
    On Error Resume Next
    Dim intHline, intVline, intHnumber, intVnumber As Integer 'Anzahl d. Linien u. Zählervar.
    Dim intPlL, intPlB, intPlHoffset, intPlVoffset As Integer 'Palettenmaße u. Offsets
    Dim intHpixel, intVpixel, intHpixelSum, intVpixelSum As Integer 'Versatzwert, Versatzsumme
    Dim intLReiheH, intLReiheQ, intLSpalteH, intLSpalteQ, a As Integer 'Anzahl Ex. Letzte Spalte, Reihe
    Dim intLeRHoffset, intLeRVoffset, intLeSHoffset, intLeSVoffset As Integer 'Offsets Letzte Reihe u. Spalte
    Dim shpShape As Shape
    ' alle Autoformen löschen
    'If Worksheets("Verpacken").Shapes(intIndex).Type = msoShapeRectangle Then
    '    Worksheets("Verpacken").Shapes(intIndex).Delete
    'End If
    If Worksheets("Verpacken").CheckBox1.Value = True Then
        intPlL = Worksheets("SVerpacken").Range("B31")
        intPlB = Worksheets("SVerpacken").Range("B30")
        intPlHoffset = Worksheets("SVerpacken").Range("B32")
        intPlVoffset = Worksheets("SVerpacken").Range("B33")
        ' Anzahl d. Linien ermitteln bei optimaler Exemplarverteilung
        If Worksheets("SVerpacken").Range("B16") = 1 Then
            'Ex. Hochformat
            intHline = Worksheets("SVerpacken").Range("B18")
            intVline = Worksheets("SVerpacken").Range("B17")
            intHpixel = Worksheets("SVerpacken").Range("C37")
            intVpixel = Worksheets("SVerpacken").Range("C36")
            Else
            'Ex. Querformat
            intHline = Worksheets("SVerpacken").Range("C18")
            intVline = Worksheets("SVerpacken").Range("C17")
            intHpixel = Worksheets("SVerpacken").Range("C36")
            intVpixel = Worksheets("SVerpacken").Range("C37")
        End If
        ' Palettenlage zeichnen
            Set shpShape = Worksheets("Verpacken").Shapes.AddShape(Type:=msoShapeRectangle, Left:=intPlHoffset, _
            Top:=intPlVoffset, Width:=intPlB, Height:=intPlL)
            With shpShape.Fill
                .ForeColor.RGB = RGB(255, 153, 0)
                .Solid
            End With
            With shpShape.Line
                .Weight = 2.5
                .ForeColor.RGB = RGB(128, 64, 0)
            End With
            With shpShape.Shadow
                .Type = msoShadow5
            End With
        'Hauptnutzen zeichnen
        If intHline And intVline > 0 Then
            intHpixelSum = intVpixel
            For intHnumber = 1 To intHline
                Worksheets("Verpacken").Shapes.AddShape Type:=msoShapeRectangle, Left:=intPlHoffset, _
                Top:=(intPlVoffset + intHpixelSum), Width:=(intHpixel * intVline), Height:=0.2
                intHpixelSum = intHpixelSum + intVpixel
            Next
            intVpixelSum = intHpixel
            For intVnumber = 1 To intVline
                Worksheets("Verpacken").Shapes.AddShape Type:=msoShapeRectangle, Left:=(intPlHoffset + intVpixelSum), _
                Top:=intPlVoffset, Width:=0.2, Height:=(intVpixel * intHline)
                intVpixelSum = intVpixelSum + intHpixel
            Next
            'Letzte Reihe u. Spalte zeichnen
            intLReiheH = Worksheets("SVerpacken").Range("B22")
            intLReiheQ = Worksheets("SVerpacken").Range("C22")
            intLSpalteH = Worksheets("SVerpacken").Range("B23")
            intLSpalteQ = Worksheets("SVerpacken").Range("C23")
                'Letzte Reihe zeichnen
                If Worksheets("SVerpacken").Range("B20") And Worksheets("SVerpacken").Range("B16") = 1 Then
                intLeRHoffset = intPlHoffset
                intLeRVoffset = (intPlVoffset + intHline * intVpixel)
                    For a = 0 To intLSpalteH
                        Worksheets("Verpacken").Shapes.AddShape Type:=msoShapeRectangle, Left:=intLeRHoffset, _
                        Top:=(intLeRVoffset + a * intHpixel), Width:=(intVpixel * intLReiheH), Height:=0.2
                    Next
                    For a = 1 To intLReiheH
                        Worksheets("Verpacken").Shapes.AddShape Type:=msoShapeRectangle, Left:=(intLeRHoffset + a * intVpixel), _
                        Top:=intLeRVoffset, Width:=0.2, Height:=(intHpixel * intLSpalteH)
                     Next
                End If
                'Letzte Spalte zeichnen
                If Worksheets("SVerpacken").Range("C21") And Worksheets("SVerpacken").Range("C16") = 1 Then
                intLeSHoffset = (intPlHoffset + intVline * intHpixel)
                intLeSVoffset = intPlVoffset
                    For a = 0 To intLSpalteQ
                        Worksheets("Verpacken").Shapes.AddShape Type:=msoShapeRectangle, Left:=intLeSHoffset, _
                        Top:=(intLeSVoffset + a * intHpixel), Width:=(intVpixel * intLReiheQ), Height:=0.2
                    Next
                    For a = 1 To intLReiheQ
                        Worksheets("Verpacken").Shapes.AddShape Type:=msoShapeRectangle, Left:=(intLeSHoffset + a * intVpixel), _
                        Top:=intLeSVoffset, Width:=0.2, Height:=(intHpixel * intLSpalteQ)
                     Next
                End If
            Else
            MsgBox "Achtung, Palette zu klein!"
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
    End If
End Sub
Sub VE_erstellen()
    'Zeichnen der Verpackungseinheit mit Banderole
    '20090819
    On Error Resume Next
    Dim intHline, intBline, intVline, intHnumber, intVnumber As Integer 'Anzahl d. Linien (vertikal, horizontal, diagonal) u. Zählervar.
    Dim intVEB, intVEH, intVEHoffset, intVEVoffset As Integer 'VE-Maße u. Offsets
    Dim shpShape As Shape
    ' alle Autoformen löschen
    'If Worksheets("Verpacken").Shapes(intIndex).Type = msoShapeRectangle Then
    '    Worksheets("Verpacken").Shapes(intIndex).Delete
    'End If
    Worksheets("Verpacken").Unprotect "bw"
    intVEB = Worksheets("SVerpacken").Range("Q46")
    intVEH = Worksheets("SVerpacken").Range("Q47")
    intVEHoffset = Worksheets("SVerpacken").Range("Q49")
    intVEVoffset = Worksheets("SVerpacken").Range("Q50")
    ' Anzahl d. Linien d. Banderole ermitteln
    intHline = Worksheets("SVerpacken").Range("Q52")
    intVline = Worksheets("SVerpacken").Range("Q53")
    ' VE zeichnen
        Set shpShape = Worksheets("Verpacken").Shapes.AddShape(Type:=msoShapeRectangle, Left:=intVEHoffset, _
        Top:=intVEVoffset, Width:=intVEB, Height:=intVEH)
        With shpShape.Fill
            .ForeColor.RGB = RGB(255, 153, 0)
            .Solid
        End With
        With shpShape.Shadow
            .Type = msoShadow5
        End With
    'Banderole zeichnen
    'über breite Seite
    If intHline > 0 Then
        For intHnumber = 1 To intHline
             Set shpShape = Worksheets("Verpacken").Shapes.AddShape(Type:=msoShapeRectangle, Left:=intVEHoffset, _
            Top:=(intVEVoffset + intHnumber * (intVEH / (intHline + 1))), Width:=intVEB, Height:=3)
            shpShape.Fill.ForeColor.RGB = RGB(0, 128, 192)
            shpShape.Line.ForeColor.RGB = RGB(0, 128, 192)
        Next
    End If
    If intVline > 0 Then
    'über hohe Seite
        For intVnumber = 1 To intVline
            Set shpShape = Worksheets("Verpacken").Shapes.AddShape(Type:=msoShapeRectangle, Left:=(intVEHoffset + intVnumber * (intVEB / (intVline + 1))), _
            Top:=intVEVoffset, Width:=3, Height:=intVEH)
            shpShape.Fill.ForeColor.RGB = RGB(0, 128, 192)
            shpShape.Line.ForeColor.RGB = RGB(0, 128, 192)
        Next
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
End Sub
Sub VE_einschweissen()
    'Zeichnen der VE Lage beim Einschweißen
    '20090820
    On Error Resume Next
    Dim intHline, intBline, intVline, intHnumber, intVnumber As Integer 'Anzahl d. Linien (vertikal, horizontal, diagonal) u. Zählervar.
    Dim intVEB, intVEH, intVEHoffset, intVEVoffset, intFB, intFBist, intFBsoll, intFBoffset As Integer 'VE- und Folienmaße u. Offsets
    Dim shpShape As Shape
    ' alle Autoformen löschen
    'If Worksheets("Verpacken").Shapes(intIndex).Type = msoShapeRectangle Then
    '    Worksheets("Verpacken").Shapes(intIndex).Delete
    'End If
    Worksheets("Verpacken").Unprotect "bw"
    If Worksheets("SVerpacken").Range("B71") > 1 Then
        If Worksheets("SVerpacken").Range("B61") = 1 Then
            'Hoch
            intVEB = Worksheets("SVerpacken").Range("B85")
            intVEH = Worksheets("SVerpacken").Range("B86")
            intHline = Worksheets("SVerpacken").Range("Q52")
            intVline = Worksheets("SVerpacken").Range("Q53")
        Else
            'Quer
            intVEB = Worksheets("SVerpacken").Range("B86")
            intVEH = Worksheets("SVerpacken").Range("B85")
            intHline = Worksheets("SVerpacken").Range("Q53")
            intVline = Worksheets("SVerpacken").Range("Q52")
        End If
        intVEHoffset = Worksheets("SVerpacken").Range("B88")
        intVEVoffset = Worksheets("SVerpacken").Range("B89")
        intFB = Worksheets("SVerpacken").Range("B90")
        intFBist = Worksheets("SVerpacken").Range("C90")
        intFBsoll = Worksheets("SVerpacken").Range("D90")
        ' Folie zeichnen und Maße prüfen
        intFBoffset = intVEVoffset + intVEH / 2 - intFB / 2
            Set shpShape = Worksheets("Verpacken").Shapes.AddShape(Type:=msoShapeRectangle, Left:=intVEHoffset - 55, _
            Top:=intFBoffset, Width:=intVEB + 100, Height:=intFB)
            If intFBsoll < intFBist Then
                With shpShape.Fill
                    .ForeColor.RGB = RGB(198, 198, 198)
                    .Solid
                End With
            Else
                With shpShape.Fill
                    .ForeColor.RGB = RGB(255, 0, 0)
                    .Solid
                End With
                MsgBox "Achtung, zu geringe Folienbreite!"
            End If
        ' VE zeichnen
            Set shpShape = Worksheets("Verpacken").Shapes.AddShape(Type:=msoShapeRectangle, Left:=intVEHoffset, _
            Top:=intVEVoffset, Width:=intVEB, Height:=intVEH)
            With shpShape.Fill
                .ForeColor.RGB = RGB(255, 153, 0)
                .Solid
            End With
        'Banderole zeichnen
        'über breite Seite
        If intHline > 0 Then
            For intHnumber = 1 To intHline
                 Set shpShape = Worksheets("Verpacken").Shapes.AddShape(Type:=msoShapeRectangle, Left:=intVEHoffset, _
                Top:=(intVEVoffset + intHnumber * (intVEH / (intHline + 1))), Width:=intVEB, Height:=3)
                shpShape.Fill.ForeColor.RGB = RGB(0, 128, 192)
                shpShape.Line.ForeColor.RGB = RGB(0, 128, 192)
            Next
        End If
        If intVline > 0 Then
        'über hohe Seite
            For intVnumber = 1 To intVline
                Set shpShape = Worksheets("Verpacken").Shapes.AddShape(Type:=msoShapeRectangle, Left:=(intVEHoffset + intVnumber * (intVEB / (intVline + 1))), _
                Top:=intVEVoffset, Width:=3, Height:=intVEH)
                shpShape.Fill.ForeColor.RGB = RGB(0, 128, 192)
                shpShape.Line.ForeColor.RGB = RGB(0, 128, 192)
            Next
        End If
        'Laufrichtung zeichnen
            Set shpShape = Worksheets("Verpacken").Shapes.AddShape(Type:=msoShapeLeftArrow, Left:=(intVEHoffset - 50), _
            Top:=(intVEVoffset + intVEH / 2 - 15), Width:=40, Height:=30)
            shpShape.Fill.ForeColor.RGB = RGB(0, 0, 0)
    End If
End Sub

