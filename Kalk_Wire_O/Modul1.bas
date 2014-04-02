Attribute VB_Name = "Modul1"
Public FFormat, FSchlaufe, FSchlaufeS, FTeilung, FSchaftMin, FSchaftMax, FStanzen As String 'Fehlervariabeln
Public v As Integer 'Versionsnummer
Function Interpolation(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X0 As Double) As Variant
    Dim Y0 As Double
    If X1 <> X2 Then
        Y0 = (Y2 - Y1) / (X2 - X1) * (X0 - X1) + Y1
        Interpolation = Y0
        If (X0 < X1 And X0 < X2) Or (X0 > X1 And X0 > X2) Then
            'MsgBox "X0 liegt nicht zwischen X1 und X2", vbInformation, "Trendberechnung"
        End If
    Else
        'MsgBox "Tabellenwerte X1 und X2 dürfen nicht übereinstimmen", vbCritical, "Achtung"
    End If
End Function
Sub Schlaufe()
Attribute Schlaufe.VB_Description = "Makro am 27.07.2005 von Enrico Dargel aufgezeichnet"
Attribute Schlaufe.VB_ProcData.VB_Invoke_Func = " \r14"
    '
    ' Schlaufenpruefung
    '
    On Error Resume Next
    Dim Teilung As String
    Dim Schlaufe As String
    Teilung = Range("Steuerung!B37")
    Schlaufe = Range("Steuerung!B38")
    Range("Material_Binden!D104").FormulaLocal = "=SVERWEIS(Eingabe!C22;Material_Binden!A105:B117;2;FALSCH)"
    If Range("Material_Binden!D104") = 0 Then
        MsgBox ("Die Schlaufengroeße " & Schlaufe & " ist nicht moeglich." & vbCrLf & vbCrLf & "Das entsprechende Werkzeug fehlt!")
        FSchlaufe = "Hinweis: Die Schlaufengröße " & Schlaufe & " kann nicht verarbeitet werden!"
        Else
            If Range("Material_Binden!D104") = 1 Then
            FSchlaufe = ""
                If Range("Eingabe!E22") = 0 Then
                    MsgBox ("Fuer die Teilung " & Teilung & " gibt es keine Schlaufe " & Schlaufe & " !" _
                     & vbCrLf & vbCrLf & "Bitte waehlen Sie eine andere Teilung oder Schlaufengroeße.")
                    FTeilung = "Hinweis: Die gewaehlte Schlaufengröße oder Teilung ist falsch!"
                Else
                    FTeilung = ""
                End If
            End If
    End If
End Sub
Sub SizeSchlaufe()
Attribute SizeSchlaufe.VB_Description = "Makro am 27.07.2005 von Enrico Dargel aufgezeichnet"
Attribute SizeSchlaufe.VB_ProcData.VB_Invoke_Func = " \r14"
    '
    ' Button
    '
    On Error Resume Next
        Dim Teilung As String
        Dim Schlaufe As String
        Teilung = Range("Steuerung!B37")
        Schlaufe = Range("Steuerung!B38")
        If Range("Eingabe!E22") = 0 Then
                MsgBox ("Fuer die Teilung " & Teilung & " gibt es keine Schlaufe " & Schlaufe & " !" _
                & vbCrLf & vbCrLf & "Bitte waehlen Sie eine andere Teilung oder Schlaufengroeße.")
                FSchlaufeS = "Hinweis: Die gewaehlte Schlaufengroeße oder Teilung ist falsch!"
            Else
                FSchlaufeS = ""
        End If
End Sub
Sub SizeSchaft()
    '
    ' Pruefung der Schaftlaenge
    '
    On Error Resume Next
        Dim Schaft_min As Integer
        Dim Schaft_max As Integer
        Schaft_min = Range("Material_Binden!H26")
        Schaft_max = Range("Material_Binden!H27")
        Schaft_min_wert = Range("Material_Binden!I26")
        Schaft_max_wert = Range("Material_Binden!I27")
        Schaft_Ist = Range("Material_Binden!I28")
        If Range("Eingabe!C19") < Schaft_min Then
            MsgBox ("Es sind nur Schaftlaengen zwischen " _
            & vbCrLf & vbCrLf & Schaft_min_wert & " mm und " & Schaft_max_wert & " mm moeglich !" & vbCrLf & vbCrLf & "Bitte waehlen Sie eine andere Schaftlaenge.")
            FSchaftMin = "Hinweis: Die gewaehlte Schaftlaenge von " & Schaft_Ist & " mm ist nicht verfuegbar! (mind. " & Schaft_min_wert & " mm)"
            Else
            FSchaftMin = ""
            If Range("Eingabe!C19") > Schaft_max Then
                MsgBox ("Es sind nur Schaftlaengen zwischen " _
                & vbCrLf & vbCrLf & Schaft_min_wert & " mm und " & Schaft_max_wert & " mm moeglich !" & vbCrLf & vbCrLf & "Bitte waehlen Sie eine andere Schaftlaenge.")
                FSchaftMax = "Hinweis: Die gewaehlte Schaftlaenge von " & Schaft_Ist & " mm ist nicht verfuegbar! (max. " & Schaft_max_wert & " mm)"
                Else
                FSchaftMax = ""
            End If
        End If
End Sub
Sub Bogenformat_Pappe() 'CommandButton1
    'ok 7.5.08
    'Bogenformat d. alternativen Pappe
    '
    If Worksheets("Steuerung").Range("H49").Value = 1 Or Worksheets("Steuerung").Range("H49").Value = 3 Then
        If Worksheets("Steuerung").Range("H70").Value = 1 Then
            Dim strBgLa, strBgLb As String 'Bogen Länge a, b
                Do
                strBgLa = InputBox("Bitte Bogenlänge in cm eingeben:")
                Worksheets("SBinden").Range("C21") = strBgLa
                strBgLb = InputBox("Bitte Bogenbreite in cm eingeben:")
                Worksheets("SBinden").Range("D21") = strBgLb
                If Worksheets("SBinden").Range("E21") = 0 Then
                    Answer = MsgBox("Ist Ihre Eingabe richtig?" & vbLf & vbLf & strBgLb & " cm x " & strBgLa & " cm", vbYesNo + 256 + vbQuestion, "Nachfrage")
                    Worksheets("SBinden").Range("G21") = "(" & strBgLb & " x " & strBgLa & "cm)"
                    Else
                    MsgBox ("Achtung! Bitte nur Ganze Zahlen eingeben!")
                    Worksheets("SBinden").Range("G21") = "Formatfehler!"
                End If
                Loop Until Answer = 6
                Worksheets("Eingabe").CommandButton1.Visible = True
            End If
        Else: Worksheets("Eingabe").CommandButton1.Visible = False
        End If
        Call NutzenCheck_Pappe
End Sub
Sub Bogenformat_Folie() 'CommandButton5
    'ok 7.5.08
    'Bogenformat d. Folie
    '
        If Worksheets("Steuerung").Range("D49").Value = 1 Or Worksheets("Steuerung").Range("D49").Value = 3 Then
            If Worksheets("Steuerung").Range("D70").Value = 1 Then
                Dim strBgLa, strBgLb As String 'Bogen Länge a, b
                    Do
                    strBgLa = InputBox("Bitte Bogenlänge in cm eingeben:")
                    Worksheets("SBinden").Range("C20") = strBgLa
                    strBgLb = InputBox("Bitte Bogenbreite in cm eingeben:")
                    Worksheets("SBinden").Range("D20") = strBgLb
                    If Worksheets("SBinden").Range("E20") = 0 Then
                        Answer = MsgBox("Ist Ihre Eingabe richtig?" & vbLf & vbLf & strBgLb & " cm x " & strBgLa & " cm", vbYesNo + 256 + vbQuestion, "Nachfrage")
                        Worksheets("SBinden").Range("G20") = "(" & strBgLb & " x " & strBgLa & "cm)"
                        Else
                        MsgBox ("Achtung! Bitte nur Ganze Zahlen eingeben!")
                        Worksheets("SBinden").Range("G20") = "Formatfehler!"
                    End If
                    Loop Until Answer = 6
                    Worksheets("Eingabe").CommandButton5.Visible = True
            End If
        Else: Worksheets("Eingabe").CommandButton5.Visible = False
        End If
        Call NutzenCheck_Folie
End Sub
Sub Rueckpappe()
    'Rückpappenvarianten
    If Worksheets("Eingabe").Range("C15").Value > 0 Then
    Dim auswahl As Integer

     auswahl = Worksheets("Steuerung").Range("H49").Value
     Select Case auswahl
        Case Is = 1 'v. Bubi, schneiden
            Worksheets("Eingabe").ComboBox12.Visible = True 'Nutzen
            Worksheets("Eingabe").ComboBox2.ListFillRange = "Steuerung!E72:F75" 'Auswahl Standard Grammatur + Alternative
                'Damit nicht die Alternative zuerst angezeigt und berechnet wird
                Worksheets("Eingabe").ComboBox2.ListIndex = 1 'Zweiten Wert d. Auswahl autom. anzeigen
                Worksheets("Steuerung").Range("H70") = 2 'Erste Standardgrammatur
            Call NutzenCheck_Pappe
            Worksheets("Eingabe").CommandButton1.Visible = False 'Formatänderung Alternativ-Bogen

        Case Is = 2 'formatig eingekauft
            Worksheets("Eingabe").ComboBox12.Visible = False
            Worksheets("Eingabe").ComboBox2.ListFillRange = "Steuerung!E72:F72" 'Auswahl alternativer Grammatur Rückpappe
            'Worksheets("Eingabe").ComboBox2.ListIndex = 0
            Worksheets("Steuerung").Range("H56").Value = 1 '1 Nutzen da formatig
            Worksheets("Steuerung").Range("H70").Value = 1 '1 alternative Grammatur
            Worksheets("Eingabe").CommandButton1.Visible = False 'kein alternatives Format
        
        Case Is = 3 'geliefert
            Worksheets("Eingabe").ComboBox12.Visible = True
            Worksheets("Eingabe").ComboBox2.ListFillRange = "Steuerung!E72:F72" 'Auswahl alternativer Grammatur Rückpappe
            Worksheets("Steuerung").Range("H70").Value = 1 '1 alternative Grammatur
            Worksheets("Eingabe").CommandButton1.Visible = True 'Formatänderung Alternativ-Bogen
            Call NutzenCheck_Pappe
        
        Case Is = 4 'geliefert u. geschnitten
            Worksheets("Eingabe").ComboBox12.Visible = False
            Worksheets("Eingabe").ComboBox2.ListFillRange = "Steuerung!E72:F72" 'Auswahl alternativer Grammatur Rückpappe
            Worksheets("Steuerung").Range("H56").Value = 1 '1 Nutzen da formatig
            Worksheets("Steuerung").Range("H70").Value = 1 '1 alternative Grammatur
            Worksheets("Eingabe").CommandButton1.Visible = False 'kein alternatives Format
     End Select
    End If
End Sub
Sub Alternativfolie()
    ' ok 7.5.08
    'On Error Resume Next
    Dim auswahl As Integer
    
    If Worksheets("Eingabe").Range("C11").Value > 0 Then
        auswahl = Worksheets("Steuerung").Range("D49").Value
         Select Case auswahl
            Case Is = 1 'v. Bubi, schneiden
                Worksheets("Eingabe").ComboBox8.Visible = True 'Nutzen
                Worksheets("Eingabe").ComboBox1.ListFillRange = "Steuerung!B72:C74" 'Auswahl Standard Grammatur + Alternative
                    'Damit nicht die Alternative zuerst angezeigt und berechnet wird
                    Worksheets("Eingabe").ComboBox1.ListIndex = 1 'Zweiten Wert d. Auswahl autom. anzeigen
                    Worksheets("Steuerung").Range("D70") = 2 'Erste Standardgrammatur
                Call Mindestmenge_Folie
                Call NutzenCheck_Folie
                Worksheets("Eingabe").CommandButton5.Visible = False 'Formatänderung Alternativ-Bogen
    
            Case Is = 2 'formatig eingekauft
                Worksheets("Eingabe").ComboBox8.Visible = False
                Worksheets("Eingabe").ComboBox1.ListFillRange = "Steuerung!B72:C72" 'Auswahl alternativer Grammatur Rückpappe
                Worksheets("Steuerung").Range("D56").Value = 1 '1 Nutzen da formatig
                Worksheets("Steuerung").Range("D70").Value = 1 '1 alternative Grammatur
                Worksheets("Eingabe").CommandButton5.Visible = False 'kein alternatives Format
            
            Case Is = 3 'geliefert
                Worksheets("Eingabe").ComboBox8.Visible = True
                Worksheets("Eingabe").ComboBox1.ListFillRange = "Steuerung!B72:C72" 'Auswahl alternativer Grammatur Rückpappe
                Worksheets("Steuerung").Range("D70").Value = 1 '1 alternative Grammatur
                Worksheets("Eingabe").CommandButton5.Visible = True 'Formatänderung Bogen
                Call NutzenCheck_Folie
            
            Case Is = 4 'geliefert u. geschnitten
                Worksheets("Eingabe").ComboBox8.Visible = False
                Worksheets("Eingabe").ComboBox1.ListFillRange = "Steuerung!B72:C72" 'Auswahl alternativer Grammatur Rückpappe
                Worksheets("Steuerung").Range("D56").Value = 1 '1 Nutzen da formatig
                Worksheets("Steuerung").Range("D70").Value = 1 '1 alternative Grammatur
                Worksheets("Eingabe").CommandButton5.Visible = False 'kein alternatives Format
         End Select
    End If
End Sub
Sub Mindestmenge_Folie()
    '
    'Prüfung der Mindestmenge f. Einkauf
    '
    Dim intGewichtA, intGewichtB, intGewichtC, intGewichtMin As Integer
    intGewichtA = Worksheets("Binden").Range("D10")
    intGewichtB = Worksheets("Binden").Range("F10")
    intGewichtC = Worksheets("Binden").Range("G10")
    intGewichtMin = Worksheets("Material_Binden").Range("E5")
    
    If Worksheets("Eingabe").Range("H11") < 3 Then
        If intGewichtA Or intGewichtB Or intGewichtC > intGewichtMin Then
            MsgBox "Hinweis: Das Foliengewicht einer Auflage überschreitet die Mindestbestellmenge (" & intGewichtMin & _
            " kg) für formatig eingekaufte Folie." & vbLf & vbLf & "Folie formatig Einkaufen?"
        End If
    End If
End Sub
Sub Materialkommentar()
    '
    ' Einfuegen der einzelnen Materialstaerken in den Kommentar
    '
    On Error Resume Next
    Dim Folie, Deckblatt, Inhalt, Rueckblatt, Rueckpappe, Summe, Schlaufe As String
    
    Worksheets("Eingabe").Unprotect "bw"
        
        Folie = Range("Stanzen!K15")
        Deckblatt = Range("Stanzen!K16")
        Inhalt = Range("Stanzen!K17")
        Rueckblatt = Range("Stanzen!K18")
        Rueckpappe = Range("Stanzen!K19")
        Summe = Range("Stanzen!K20")
        Schlaufe = Range("Stanzen!K21")
        Range("Eingabe!C32").ClearComments
        Range("Eingabe!C32").AddComment
        Range("Eingabe!C32").Comment.Visible = False
        Range("Eingabe!C32").Comment.Text Text:="Einzelstärken:" & vbLf & "=================" & vbLf & Folie & " mm Folie" & vbLf & Deckblatt & " mm Deckblatt" & vbLf & Inhalt & " mm Inhalt" _
        & vbLf & Rueckblatt & " mm Rückblatt" & vbLf & Rueckpappe & " mm Rückpappe" & vbLf & "=================" & vbLf & Summe & " mm Summe" & vbLf & Schlaufe & " mm " & Chr(248) & " Schlaufe"
        Range("Eingabe!C32").Comment.Shape.TextFrame.AutoSize = True
End Sub
Sub Bindemaschine()
    ' ok 6.5.08
    ' Maschinenpruefung
    ' Alle Schlaufen >1" koennen nur auf CLS verarbeitet werden
    '
    On Error Resume Next
    Dim Schlaufe As String
    
    Schlaufe = Range("Steuerung!B38")
    If Range("Eingabe!C22") > 11 And Range("SBinden!B4") < 4 Then
            MsgBox ("Achtung! " & vbCrLf & vbCrLf & "Die Schlaufe " & Schlaufe & _
            " kann nur auf der CLS verarbeitet werden." & vbCrLf & vbCrLf & _
            "Bitte beim Binden richtige Maschine auswählen.")
    End If
End Sub
Sub Produkt()
    '
    ' Anzeigen d. Produktangaben bei "Verpacken"
    '
    On Error Resume Next
    Dim format, Gewicht, Dicke As String
    
    Worksheets("Verpacken").Unprotect "bw"
    
        format = Worksheets("Eingabe").CommandButton2.Caption
        Dicke = Range("Eingabe!C32")
        Gewicht = Range("Eingabe!C34")
        Worksheets("Verpacken").Label1.Caption = "Produkt:" & vbLf & "======" & vbLf & vbLf & "Format: " & vbLf & format _
         & vbLf & vbLf & "Stärke: " & vbLf & Dicke & " mm" & vbLf & vbLf & "Gewicht: " & vbLf & Gewicht & " g"
    Worksheets("Verpacken").Protect "bw"
End Sub
Sub NutzenCheck_Folie()
    ' ok 7.5.08
    ' Überprüfung der Nutzenanzahl
    '
    If Worksheets("Steuerung").Range("D49").Value = 1 Or Worksheets("Steuerung").Range("D49").Value = 3 Then
    'Nur Prüfen wenn "von Bubi schneiden" oder "geliefert"
        Dim NutzenMax As Integer
        
        If Worksheets("Steuerung").Range("D70").Value = 1 Then
            NutzenMax = Worksheets("SBinden").Range("C50")
            Else
            NutzenMax = Worksheets("SBinden").Range("F50")
        End If
        
        If Worksheets("Steuerung").Range("D56") > NutzenMax Then
            MsgBox "Achtung!" & vbLf & vbLf & "Ihre Nutzenanzahl der Folie ist zu hoch." _
            & vbLf & vbLf & "(Maximal: " & NutzenMax & " Nutzen)."
            Worksheets("Eingabe").ComboBox8.BackColor = &HFF 'rot &H000000FF
            
        End If
        If Worksheets("Steuerung").Range("D56") < NutzenMax Then
            MsgBox "Hinweis:" & vbLf & vbLf & "Ihre Nutzenanzahl der Folie ist zu gering." _
            & vbLf & vbLf & "(Mindesten: " & NutzenMax & " Nutzen)."
            Worksheets("Eingabe").ComboBox8.BackColor = &HFFFF00 'blau
        End If
        If Worksheets("Steuerung").Range("D56") = NutzenMax Then
            Worksheets("Eingabe").ComboBox8.BackColor = &HFFFFFF 'weiß
        End If
    Else: Worksheets("Eingabe").ComboBox8.BackColor = &HFFFFFF 'weiß
    End If
End Sub
Sub NutzenCheck_Pappe()
    ' ok 7.5.08
    ' Überprüfung der Nutzenanzahl
    '
    If Worksheets("Steuerung").Range("H49").Value = 1 Or Worksheets("Steuerung").Range("H49").Value = 3 Then
    'Nur Prüfen wenn "von Bubi schneiden" oder "geliefert"
        Dim NutzenMax As Integer
        
        If Worksheets("Steuerung").Range("H70").Value = 1 Then
            NutzenMax = Worksheets("SBinden").Range("I50")
            Else
            NutzenMax = Worksheets("SBinden").Range("L50")
        End If
        
        If Worksheets("Steuerung").Range("H56") > NutzenMax Then
            MsgBox "Achtung!" & vbLf & vbLf & "Ihre Nutzenanzahl der Rückpappe ist zu hoch." _
            & vbLf & vbLf & "(Maximal: " & NutzenMax & " Nutzen)."
            Worksheets("Eingabe").ComboBox12.BackColor = &HFF 'rot &H000000FF
            
        End If
        If Worksheets("Steuerung").Range("H56") < NutzenMax Then
            MsgBox "Hinweis:" & vbLf & vbLf & "Ihre Nutzenanzahl der Rückpappe ist zu gering." _
            & vbLf & vbLf & "(Mindestens: " & NutzenMax & " Nutzen)."
            Worksheets("Eingabe").ComboBox12.BackColor = &HFFFF00 'blau
        End If
        If Worksheets("Steuerung").Range("H56") = NutzenMax Then
            Worksheets("Eingabe").ComboBox12.BackColor = &HFFFFFF 'weiß
        End If
    Else: Worksheets("Eingabe").ComboBox12.BackColor = &HFFFFFF 'weiß
    End If
End Sub
Sub Stanzen_Size()
    'Überprüfung der Stärke
    If Worksheets("Stanzen").Range("H12") = True Then
        Dim intSize As Double
        intSize = Worksheets("Stanzen").Range("F12").Value
        Debug.Print intSize
            FStanzen = "Hinweis: Die Stanzstärke weicht um " & intSize & " mm ab!"
            MsgBox (FStanzen)
        Else: FStanzen = ""
    End If
End Sub
