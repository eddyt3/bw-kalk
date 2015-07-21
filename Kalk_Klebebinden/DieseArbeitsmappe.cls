VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'Copyright by E.Dargel ed@dissenter.de
'
Public Sub Workbook_BeforePrint(Cancel As Boolean)
    'Druckbereich nur auf Eingabeseite festlegen
    '20150421
    Dim Letzte As Long 'Letzte Zeile
    Dim a As Integer
    Letzte = Worksheets("Eingabe").Cells(Rows.Count, 1).End(xlUp).Row
    ThisWorkbook.Worksheets("Eingabe").PageSetup.PrintArea = "A1:L" & Letzte
    Application.ScreenUpdating = False
    Dim wks As Worksheet
    Dim v As String
    v = Worksheets("Steuerung").Range("B178") & "/" & Worksheets("Steuerung").Range("A178")
    ThisWorkbook.Worksheets("Eingabe").PageSetup.LeftFooter = "&""Verdana""&06" & Application.UserName & " " & format("&D", "dd.mm.yy") & " " & "&T"
    ThisWorkbook.Worksheets("Eingabe").PageSetup.CenterFooter = "&""Verdana""&06" & Chr(10) & ActiveWorkbook.Path & "\" & ActiveWorkbook.Name
    ThisWorkbook.Worksheets("Eingabe").PageSetup.RightFooter = "&""Verdana""&06" & " (v" & v & ")"
    Call Dokumenteigenschaften_Soll
    Application.ScreenUpdating = True
End Sub
Private Sub Workbook_Open()
On Error Resume Next
fntSize = 10
With Worksheets("Eingabe").CheckBox1   '1
   .Height = 18.75: .Width = 57.75: .Top = 239.25: .Left = 9.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox2   '2
   .Height = 18.75: .Width = 57.75: .Top = 351.75: .Left = 9.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox3   '3
   .Height = 18.75: .Width = 57.75: .Top = 465.75: .Left = 9.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox1   '4
   .Height = 21.75: .Width = 114: .Top = 66: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox4   '5
   .Height = 12: .Width = 12: .Top = 792.75: .Left = 595.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox5   '6
   .Height = 12: .Width = 12: .Top = 818.25: .Left = 595.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox6   '7
   .Height = 12: .Width = 12: .Top = 843.75: .Left = 595.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox7   '8
   .Height = 12: .Width = 12: .Top = 894.75: .Left = 595.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox8   '9
   .Height = 12: .Width = 12: .Top = 920.25: .Left = 595.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox9   '10
   .Height = 12: .Width = 12: .Top = 945.75: .Left = 595.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox10   '11
   .Height = 12: .Width = 12: .Top = 971.25: .Left = 595.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox11   '12
   .Height = 12: .Width = 12: .Top = 1022.25: .Left = 595.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox2   '13
   .Height = 19.5: .Width = 86.25: .Top = 738: .Left = 3: .FontSize = fntSize
End With
With Worksheets("Eingabe").CommandButton2   '14
   .Height = 23.25: .Width = 73.5: .Top = 33.75: .Left = 18: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox3   '15
   .Height = 19.5: .Width = 113.25: .Top = 122.25: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox4   '16
   .Height = 19.5: .Width = 113.25: .Top = 275.25: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox5   '17
   .Height = 19.5: .Width = 113.25: .Top = 388.5: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox6   '18
   .Height = 19.5: .Width = 113.25: .Top = 502.5: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox7   '19
   .Height = 19.5: .Width = 113.25: .Top = 104.25: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox8   '20
   .Height = 19.5: .Width = 113.25: .Top = 257.25: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox9   '21
   .Height = 19.5: .Width = 113.25: .Top = 370.5: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox10   '22
   .Height = 19.5: .Width = 113.25: .Top = 484.5: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox12   '23
   .Height = 12: .Width = 12: .Top = 996.75: .Left = 595.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox11   '24
   .Height = 19.5: .Width = 124.5: .Top = 738.75: .Left = 147: .FontSize = fntSize
End With
With Worksheets("Eingabe").Bemerkung   '25
   .Height = 149.25: .Width = 213.75: .Top = 1044: .Left = 570: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox12   '26
   .Height = 19.5: .Width = 96: .Top = 123: .Left = 276: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox1   '27
   .Height = 19.5: .Width = 42: .Top = 123: .Left = 375.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox2   '28
   .Height = 19.5: .Width = 42: .Top = 123: .Left = 429: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox1   '29
   .Height = 15.75: .Width = 65.25: .Top = 54: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox2   '30
   .Height = 15.75: .Width = 65.25: .Top = 137.25: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox3   '31
   .Height = 15.75: .Width = 65.25: .Top = 221.25: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox4   '32
   .Height = 15.75: .Width = 65.25: .Top = 304.5: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").ComboBox1   '33
   .Height = 18.75: .Width = 237: .Top = 21.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Schneiden").TextBox1   '34
   .Height = 129.75: .Width = 578.25: .Top = 677.25: .Left = 3.75: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox1   '35
   .Height = 17.25: .Width = 174.75: .Top = 48.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox2   '36
   .Height = 17.25: .Width = 174.75: .Top = 296.25: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox3   '37
   .Height = 17.25: .Width = 174.75: .Top = 527.25: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox4   '38
   .Height = 17.25: .Width = 174.75: .Top = 758.25: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox1   '39
   .Height = 18: .Width = 60.75: .Top = 88.5: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox2   '40
   .Height = 18: .Width = 60.75: .Top = 323.25: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox3   '41
   .Height = 18: .Width = 60.75: .Top = 554.25: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox4   '42
   .Height = 18: .Width = 60.75: .Top = 785.25: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox5   '43
   .Height = 17.25: .Width = 174.75: .Top = 66: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox6   '44
   .Height = 17.25: .Width = 174.75: .Top = 314.25: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox7   '45
   .Height = 17.25: .Width = 174.75: .Top = 544.5: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox8   '46
   .Height = 17.25: .Width = 174.75: .Top = 775.5: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").TextBox1   '47
   .Height = 198: .Width = 156.75: .Top = 1044.75: .Left = 612: .FontSize = fntSize
End With
With Worksheets("Kleben").CheckBox1   '48
   .Height = 17.25: .Width = 72.75: .Top = 32.25: .Left = 362.25: .FontSize = fntSize
End With
With Worksheets("Kleben").ComboBox1   '49
   .Height = 17.25: .Width = 108.75: .Top = 53.25: .Left = 102: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox1   '50
   .Height = 72: .Width = 200.25: .Top = 53.25: .Left = 744: .FontSize = fntSize
End With
With Worksheets("Kleben").CheckBox2   '51
   .Height = 17.25: .Width = 72.75: .Top = 63: .Left = 362.25: .FontSize = fntSize
End With
With Worksheets("Kleben").CheckBox3   '52
   .Height = 19.5: .Width = 155.25: .Top = 53.25: .Left = 447: .FontSize = fntSize
End With
With Worksheets("Kleben").ComboBox2   '53
   .Height = 17.25: .Width = 203.25: .Top = 33: .Left = 104.25: .FontSize = fntSize
End With
With Worksheets("Kleben").ComboBox3   '54
   .Height = 17.25: .Width = 174.75: .Top = 53.25: .Left = 35.25: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox2   '55
   .Height = 106.5: .Width = 627.75: .Top = 381: .Left = 7.5: .FontSize = fntSize
End With
With Worksheets("Kleben").CommandButton21   '56
   .Height = 19.5: .Width = 39: .Top = 53.25: .Left = 342: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox8   '57
   .Height = 19.5: .Width = 27.75: .Top = 53.25: .Left = 213.75: .FontSize = fntSize
End With
With Worksheets("Zusammentragen").ComboBox2   '58
   .Height = 17.25: .Width = 203.25: .Top = 62.25: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Zusammentragen").CheckBox1   '59
   .Height = 14.25: .Width = 90: .Top = 86.25: .Left = 4.5: .FontSize = fntSize
End With
With Worksheets("Zusammentragen").TextBox1   '60
   .Height = 103.5: .Width = 630: .Top = 719.25: .Left = 6: .FontSize = fntSize
End With
With Worksheets("Zusammentragen").CheckBox2   '61
   .Height = 17.25: .Width = 72.75: .Top = 348: .Left = 123.75: .FontSize = fntSize
End With
With Worksheets("Fadenheften").TextBox1   '62
   .Height = 92.25: .Width = 590.25: .Top = 558.75: .Left = 5.25: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox1   '63
   .Height = 15.75: .Width = 107.25: .Top = 44.25: .Left = 291: .FontSize = fntSize
End With
With Worksheets("Binden").CheckBox1   '64
   .Height = 15.75: .Width = 38.25: .Top = 28.5: .Left = 293.25: .FontSize = fntSize
End With
With Worksheets("Binden").CheckBox2   '65
   .Height = 15.75: .Width = 57: .Top = 28.5: .Left = 337.5: .FontSize = fntSize
End With
With Worksheets("Binden").CheckBox3   '66
   .Height = 15.75: .Width = 73.5: .Top = 45: .Left = 76.5: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox2   '67
   .Height = 15.75: .Width = 107.25: .Top = 60: .Left = 291: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox3   '68
   .Height = 15.75: .Width = 107.25: .Top = 75.75: .Left = 291: .FontSize = fntSize
End With
With Worksheets("Binden").TextBox1   '69
   .Height = 131.25: .Width = 599.25: .Top = 588: .Left = 5.25: .FontSize = fntSize
End With
With Worksheets("3Seitenbeschnitt").CheckBox1   '70
   .Height = 17.25: .Width = 72.75: .Top = 39.75: .Left = 112.5: .FontSize = fntSize
End With
With Worksheets("3Seitenbeschnitt").CheckBox2   '71
   .Height = 17.25: .Width = 74.25: .Top = 252.75: .Left = 111.75: .FontSize = fntSize
End With
With Worksheets("3Seitenbeschnitt").TextBox1   '72
   .Height = 65.25: .Width = 620.25: .Top = 722.25: .Left = 3: .FontSize = fntSize
End With
With Worksheets("Bohren").ComboBox1   '73
   .Height = 18.75: .Width = 252.75: .Top = 21.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Bohren").CheckBox2   '74
   .Height = 15.75: .Width = 65.25: .Top = 121.5: .Left = 585.75: .FontSize = fntSize
End With
With Worksheets("Bohren").CheckBox4   '75
   .Height = 15.75: .Width = 65.25: .Top = 54.75: .Left = 107.25: .FontSize = fntSize
End With
With Worksheets("Bohren").TextBox1   '76
   .Height = 116.25: .Width = 567.75: .Top = 489: .Left = 3: .FontSize = fntSize
End With
With Worksheets("Verpacken").Label1   '77
   .Height = 146.25: .Width = 74.25: .Top = 87: .Left = 14.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox1   '78
   .Height = 21: .Width = 296.25: .Top = 302.25: .Left = 324: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox2   '79
   .Height = 19.5: .Width = 296.25: .Top = 61.5: .Left = 324: .FontSize = fntSize
End With
With Worksheets("Verpacken").CheckBox1   '80
   .Height = 18: .Width = 55.5: .Top = 282: .Left = 333: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox3   '81
   .Height = 21: .Width = 296.25: .Top = 348.75: .Left = 324: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox4   '82
   .Height = 21: .Width = 296.25: .Top = 663.75: .Left = 324: .FontSize = fntSize
End With
With Worksheets("Verpacken").CommandButton1   '83
   .Height = 19.5: .Width = 39: .Top = 78.75: .Left = 624: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox5   '84
   .Height = 21: .Width = 200.25: .Top = 615.75: .Left = 324: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox6   '85
   .Height = 21: .Width = 94.5: .Top = 615.75: .Left = 525.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").CheckBox4   '86
   .Height = 18: .Width = 111: .Top = 457.5: .Left = 204.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox7   '87
   .Height = 21: .Width = 296.25: .Top = 851.25: .Left = 324: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox8   '88
   .Height = 21: .Width = 296.25: .Top = 382.5: .Left = 324: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox9   '89
   .Height = 21: .Width = 193.5: .Top = 424.5: .Left = 427.5: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox10   '90
   .Height = 21: .Width = 48: .Top = 425.25: .Left = 324: .FontSize = fntSize
End With
With Worksheets("Verpacken").ToggleButton1   '91
   .Height = 36.75: .Width = 93.75: .Top = 567: .Left = 114: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox11   '92
   .Height = 21: .Width = 48: .Top = 425.25: .Left = 372: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox12   '93
   .Height = 21: .Width = 296.25: .Top = 403.5: .Left = 324: .FontSize = fntSize
End With
With Worksheets("Verpacken").CommandButton2   '94
   .Height = 35.25: .Width = 83.25: .Top = 81.75: .Left = 114: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox13   '95
   .Height = 21: .Width = 48: .Top = 145.5: .Left = 324: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox14   '96
   .Height = 21: .Width = 48: .Top = 145.5: .Left = 372.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox15   '97
   .Height = 21: .Width = 193.5: .Top = 145.5: .Left = 427.5: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox16   '98
   .Height = 21: .Width = 96.75: .Top = 911.25: .Left = 324: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox17   '99
   .Height = 21: .Width = 97.5: .Top = 971.25: .Left = 324: .FontSize = fntSize
End With
With Worksheets("Verpacken").TextBox1   '100
   .Height = 102.75: .Width = 610.5: .Top = 1378.5: .Left = 3.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").Label2   '101
   .Height = 64.5: .Width = 76.5: .Top = 499.5: .Left = 123.75: .FontSize = fntSize
End With
With Worksheets("Produktionsdaten").ComboBox1   '102
   .Height = 18: .Width = 78.75: .Top = 1110: .Left = 112.5: .FontSize = fntSize
End With
With Worksheets("Produktionsdaten").ComboBox2   '103
   .Height = 18: .Width = 78.75: .Top = 1110: .Left = 193.5: .FontSize = fntSize
End With
With Worksheets("Material").CommandButton1   '104
   .Height = 39.75: .Width = 86.25: .Top = 917.25: .Left = 3: .FontSize = fntSize
End With
End Sub
