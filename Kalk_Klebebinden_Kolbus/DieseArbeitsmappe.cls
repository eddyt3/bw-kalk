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
    Worksheets("Eingabe").PageSetup.PrintArea = "A1:L" & Letzte
    Application.ScreenUpdating = False
    Dim wks As Worksheet
    Dim v As String
    v = Worksheets("Steuerung").Range("B178") & "/" & Worksheets("Steuerung").Range("A178")
    ThisWorkbook.Worksheets("Eingabe").PageSetup.LeftFooter = "&""Verdana""&06" & Application.UserName & " " & format("&D", "dd.mm.yy") & " " & "&T"
    ThisWorkbook.Worksheets("Eingabe").PageSetup.CenterFooter = "&""Verdana""&06" & Chr(10) & ThisWorkbook.Path & "\" & ThisWorkbook.Name
    ThisWorkbook.Worksheets("Eingabe").PageSetup.RightFooter = "&""Verdana""&06" & " (v" & v & ")"
    Call Dokumenteigenschaften_Soll
    Application.ScreenUpdating = True
End Sub
Private Sub Workbook_Open()
On Error Resume Next
fntSize = 10
With Worksheets("Eingabe").CheckBox1   '1
   .Height = 18.75: .Width = 57.75: .Top = 282: .Left = 9.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox2   '2
   .Height = 18.75: .Width = 57.75: .Top = 394.5: .Left = 9.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox3   '3
   .Height = 18.75: .Width = 57.75: .Top = 508.5: .Left = 9.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox4   '4
   .Height = 12: .Width = 12: .Top = 826.5: .Left = 568.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox5   '5
   .Height = 12: .Width = 12: .Top = 852: .Left = 568.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox6   '6
   .Height = 12: .Width = 12: .Top = 979.5: .Left = 568.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox7   '7
   .Height = 12: .Width = 12: .Top = 1030.5: .Left = 568.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox2   '8
   .Height = 19.5: .Width = 86.25: .Top = 81.75: .Left = 3.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").CommandButton2   '9
   .Height = 23.25: .Width = 73.5: .Top = 33.75: .Left = 18: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox3   '10
   .Height = 19.5: .Width = 123: .Top = 81.75: .Left = 117: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox5   '11
   .Height = 19.5: .Width = 86.25: .Top = 81: .Left = 246.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox1   '12
   .Height = 19.5: .Width = 113.25: .Top = 164.25: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox7   '13
   .Height = 19.5: .Width = 113.25: .Top = 318: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox8   '14
   .Height = 19.5: .Width = 113.25: .Top = 431.25: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox6   '15
   .Height = 19.5: .Width = 113.25: .Top = 544.5: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox4   '16
   .Height = 19.5: .Width = 113.25: .Top = 146.25: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox9   '17
   .Height = 19.5: .Width = 113.25: .Top = 299.25: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox10   '18
   .Height = 19.5: .Width = 113.25: .Top = 412.5: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox11   '19
   .Height = 19.5: .Width = 113.25: .Top = 526.5: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox8   '20
   .Height = 12: .Width = 12: .Top = 1005: .Left = 568.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox9   '21
   .Height = 12: .Width = 12: .Top = 903: .Left = 568.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox10   '22
   .Height = 12: .Width = 12: .Top = 954: .Left = 568.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CommandButton4   '23
   .Height = 35.25: .Width = 83.25: .Top = 69: .Left = 634.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox12   '24
   .Height = 19.5: .Width = 92.25: .Top = 81.75: .Left = 345: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox2   '25
   .Height = 149.25: .Width = 282: .Top = 1060.5: .Left = 539.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox13   '26
   .Height = 25.5: .Width = 111.75: .Top = 108: .Left = 3.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox14   '27
   .Height = 19.5: .Width = 160.5: .Top = 82.5: .Left = 442.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox11   '28
   .Height = 12: .Width = 12: .Top = 928.5: .Left = 568.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox15   '29
   .Height = 19.5: .Width = 96: .Top = 165: .Left = 245.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox1   '30
   .Height = 19.5: .Width = 42: .Top = 165.75: .Left = 345: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox3   '31
   .Height = 19.5: .Width = 42: .Top = 165: .Left = 399: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox1   '32
   .Height = 15.75: .Width = 65.25: .Top = 54: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox2   '33
   .Height = 17.25: .Width = 65.25: .Top = 137.25: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox3   '34
   .Height = 17.25: .Width = 65.25: .Top = 221.25: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox4   '35
   .Height = 17.25: .Width = 65.25: .Top = 304.5: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").ComboBox1   '36
   .Height = 18.75: .Width = 237: .Top = 21.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Schneiden").TextBox1   '37
   .Height = 129.75: .Width = 578.25: .Top = 678: .Left = 5.25: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox1   '38
   .Height = 17.25: .Width = 174.75: .Top = 48.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox2   '39
   .Height = 17.25: .Width = 174.75: .Top = 295.5: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox3   '40
   .Height = 17.25: .Width = 174.75: .Top = 526.5: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox4   '41
   .Height = 17.25: .Width = 174.75: .Top = 757.5: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox1   '42
   .Height = 18: .Width = 60.75: .Top = 88.5: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox2   '43
   .Height = 18: .Width = 60.75: .Top = 322.5: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox3   '44
   .Height = 18: .Width = 60.75: .Top = 553.5: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox4   '45
   .Height = 18: .Width = 60.75: .Top = 784.5: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox5   '46
   .Height = 17.25: .Width = 174.75: .Top = 66: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox6   '47
   .Height = 17.25: .Width = 174.75: .Top = 313.5: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox7   '48
   .Height = 17.25: .Width = 174.75: .Top = 543.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox8   '49
   .Height = 17.25: .Width = 174.75: .Top = 774.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").TextBox1   '50
   .Height = 129.75: .Width = 588.75: .Top = 1254.75: .Left = 5.25: .FontSize = fntSize
End With
With Worksheets("Kleben").CheckBox1   '51
   .Height = 17.25: .Width = 72.75: .Top = 32.25: .Left = 362.25: .FontSize = fntSize
End With
With Worksheets("Kleben").ComboBox1   '52
   .Height = 17.25: .Width = 108.75: .Top = 496.5: .Left = 102: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox1   '53
   .Height = 72: .Width = 200: .Top = 58: .Left = 744: .FontSize = fntSize
End With
With Worksheets("Kleben").CheckBox2   '54
   .Height = 17.25: .Width = 72.75: .Top = 915: .Left = 362.25: .FontSize = fntSize
End With
With Worksheets("Kleben").CheckBox3   '55
   .Height = 19.5: .Width = 155.25: .Top = 508.5: .Left = 447: .FontSize = fntSize
End With
With Worksheets("Kleben").ComboBox2   '56
   .Height = 17.25: .Width = 203.25: .Top = 33: .Left = 104.25: .FontSize = fntSize
End With
With Worksheets("Kleben").ComboBox3   '57
   .Height = 17.25: .Width = 174.75: .Top = 935.25: .Left = 3: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox2   '58
   .Height = 129.75: .Width = 636: .Top = 1230.75: .Left = 3.75: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox8   '59
   .Height = 19.5: .Width = 27.75: .Top = 508.5: .Left = 213.75: .FontSize = fntSize
End With
With Worksheets("Kleben").CommandButton21   '60
   .Height = 19.5: .Width = 39: .Top = 578.25: .Left = 341.25: .FontSize = fntSize
End With
With Worksheets("Fadenheften").TextBox1   '61
   .Height = 129.75: .Width = 590.25: .Top = 558.75: .Left = 3.75: .FontSize = fntSize
End With
With Worksheets("Basis_Kolbus").TextBox1   '62
   .Height = 129.75: .Width = 623.25: .Top = 974.25: .Left = 3.75: .FontSize = fntSize
End With
With Worksheets("Zusammentragen").ComboBox2   '63
   .Height = 17.25: .Width = 203.25: .Top = 62.25: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Zusammentragen").CheckBox1   '64
   .Height = 14.25: .Width = 90: .Top = 86.25: .Left = 4.5: .FontSize = fntSize
End With
With Worksheets("Zusammentragen").TextBox1   '65
   .Height = 103.5: .Width = 630: .Top = 719.25: .Left = 6: .FontSize = fntSize
End With
With Worksheets("Zusammentragen").CheckBox2   '66
   .Height = 17.25: .Width = 72.75: .Top = 348: .Left = 123.75: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox1   '67
   .Height = 15.75: .Width = 107.25: .Top = 44.25: .Left = 291: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox2   '68
   .Height = 15.75: .Width = 107.25: .Top = 60: .Left = 291: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox3   '69
   .Height = 15.75: .Width = 107.25: .Top = 75.75: .Left = 291: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox4   '70
   .Height = 15.75: .Width = 107.25: .Top = 237.75: .Left = 291: .FontSize = fntSize
End With
With Worksheets("Binden").CheckBox3   '71
   .Height = 14.25: .Width = 168.75: .Top = 332.25: .Left = 663.75: .FontSize = fntSize
End With
With Worksheets("Binden").CommandButton1   '72
   .Height = 19.5: .Width = 39: .Top = 105: .Left = 249: .FontSize = fntSize
End With
With Worksheets("Binden").TextBox1   '73
   .Height = 129.75: .Width = 601.5: .Top = 607.5: .Left = 5.25: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox5   '74
   .Height = 15.75: .Width = 107.25: .Top = 28.5: .Left = 291.75: .FontSize = fntSize
End With
With Worksheets("Bohren").ComboBox1   '75
   .Height = 18.75: .Width = 252.75: .Top = 21.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Bohren").CheckBox2   '76
   .Height = 15.75: .Width = 65.25: .Top = 121.5: .Left = 585.75: .FontSize = fntSize
End With
With Worksheets("Bohren").CheckBox4   '77
   .Height = 15.75: .Width = 65.25: .Top = 54.75: .Left = 107.25: .FontSize = fntSize
End With
With Worksheets("Bohren").TextBox1   '78
   .Height = 129.75: .Width = 573.75: .Top = 489: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").Label1   '79
   .Height = 146.25: .Width = 74.25: .Top = 87: .Left = 14.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox1   '80
   .Height = 21: .Width = 296.25: .Top = 303: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox2   '81
   .Height = 19.5: .Width = 296.25: .Top = 61.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").CheckBox1   '82
   .Height = 18: .Width = 55.5: .Top = 282.75: .Left = 345: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox3   '83
   .Height = 21: .Width = 296.25: .Top = 349.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox4   '84
   .Height = 21: .Width = 296.25: .Top = 664.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").CommandButton1   '85
   .Height = 19.5: .Width = 39: .Top = 78.75: .Left = 636: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox5   '86
   .Height = 21: .Width = 200.25: .Top = 616.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox6   '87
   .Height = 21: .Width = 94.5: .Top = 616.5: .Left = 537.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").CheckBox4   '88
   .Height = 18: .Width = 123: .Top = 459: .Left = 204.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox7   '89
   .Height = 21: .Width = 296.25: .Top = 852: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox8   '90
   .Height = 21: .Width = 296.25: .Top = 383.25: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox9   '91
   .Height = 21: .Width = 193.5: .Top = 425.25: .Left = 439.5: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox10   '92
   .Height = 21: .Width = 48: .Top = 426: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ToggleButton1   '93
   .Height = 36.75: .Width = 93.75: .Top = 567: .Left = 117: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox11   '94
   .Height = 21: .Width = 48: .Top = 426: .Left = 384: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox12   '95
   .Height = 21: .Width = 296.25: .Top = 404.25: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").CommandButton2   '96
   .Height = 35.25: .Width = 83.25: .Top = 81.75: .Left = 114: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox13   '97
   .Height = 21: .Width = 48: .Top = 145.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox14   '98
   .Height = 21: .Width = 48: .Top = 145.5: .Left = 384.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox15   '99
   .Height = 21: .Width = 193.5: .Top = 145.5: .Left = 439.5: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox16   '100
   .Height = 21: .Width = 96.75: .Top = 910.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox17   '101
   .Height = 21: .Width = 200.25: .Top = 970.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").CheckBox2   '102
   .Height = 18: .Width = 93.75: .Top = 1017: .Left = 338.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").TextBox1   '103
   .Height = 129.75: .Width = 619.5: .Top = 1419.75: .Left = 3.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").Label2   '104
   .Height = 64.5: .Width = 76.5: .Top = 499.5: .Left = 128.25: .FontSize = fntSize
End With
With Worksheets("Produktionsdaten").ComboBox1   '105
   .Height = 18: .Width = 78.75: .Top = 974.25: .Left = 112.5: .FontSize = fntSize
End With
With Worksheets("Produktionsdaten").ComboBox2   '106
   .Height = 18: .Width = 78.75: .Top = 974.25: .Left = 193.5: .FontSize = fntSize
End With
With Worksheets("Material").CommandButton1   '107
   .Height = 25.5: .Width = 145.5: .Top = 0.75: .Left = 110.25: .FontSize = fntSize
End With
End Sub
