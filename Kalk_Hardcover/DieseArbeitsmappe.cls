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
    '20180328
    Dim Letzte As Long 'Letzte Zeile
    Dim i As Integer
    Letzte = Worksheets("Eingabe").Cells(Rows.Count, 1).End(xlUp).Row
    Worksheets("Eingabe").PageSetup.PrintArea = "A1:L" & Letzte
    Application.ScreenUpdating = False
    Dim wks As Worksheet
    Dim v As String
    v = Worksheets("Steuerung").Range("B178") & "/" & Worksheets("Steuerung").Range("A178")
    For i = 1 To Sheets.Count
        ActiveWorkbook.Sheets(i).PageSetup.LeftFooter = "&""Verdana""&06" & Application.UserName & " " & format("&D", "dd.mm.yy") & " " & "&T"
        ActiveWorkbook.Sheets(i).PageSetup.CenterFooter = "&""Verdana""&06" & Chr(10) & ThisWorkbook.Path & "\" & ThisWorkbook.Name
        ActiveWorkbook.Sheets(i).PageSetup.RightFooter = "&""Verdana""&06" & " (v" & v & ")"
    Next
    'MsgBox "Alle Footer gesetzt."
    Call Dokumenteigenschaften_Soll
    Application.ScreenUpdating = True
End Sub
Private Sub Workbook_Open()
ActiveWorkbook.CheckCompatibility = False 'Unterdrückung d. Fehlermeldung beim Speichern in alte Excel Version
' Buttonformat zurücksetzen wg. Problem der unterschiedlichen Größen bei unterschiedlichen Bildschirmauflösungen
On Error Resume Next
fntSize = 10
With Worksheets("Eingabe").CheckBox1   '1
   .Height = 18.75: .Width = 57.75: .Top = 316.5: .Left = 9.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox2   '2
   .Height = 18.75: .Width = 57.75: .Top = 429: .Left = 9.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox3   '3
   .Height = 18.75: .Width = 57.75: .Top = 543: .Left = 9.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox4   '4
   .Height = 12: .Width = 12: .Top = 867: .Left = 565.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox5   '5
   .Height = 12: .Width = 12: .Top = 892.5: .Left = 565.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox6   '6
   .Height = 12: .Width = 12: .Top = 994.5: .Left = 565.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox7   '7
   .Height = 12: .Width = 12: .Top = 1071: .Left = 565.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox2   '8
   .Height = 19.5: .Width = 86.25: .Top = 81.75: .Left = 3.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").CommandButton2   '9
   .Height = 23.25: .Width = 73.5: .Top = 33.75: .Left = 18: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox3   '10
   .Height = 19.5: .Width = 45: .Top = 81.75: .Left = 117: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox5   '11
   .Height = 19.5: .Width = 86.25: .Top = 81: .Left = 246.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox1   '12
   .Height = 19.5: .Width = 113.25: .Top = 198.75: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox7   '13
   .Height = 19.5: .Width = 113.25: .Top = 352.5: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox8   '14
   .Height = 19.5: .Width = 113.25: .Top = 465.75: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox6   '15
   .Height = 19.5: .Width = 113.25: .Top = 579: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox4   '16
   .Height = 19.5: .Width = 113.25: .Top = 180.75: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox9   '17
   .Height = 19.5: .Width = 113.25: .Top = 333.75: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox10   '18
   .Height = 19.5: .Width = 113.25: .Top = 447: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox11   '19
   .Height = 19.5: .Width = 113.25: .Top = 561: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox8   '20
   .Height = 12: .Width = 12: .Top = 1020: .Left = 565.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox9   '21
   .Height = 12: .Width = 12: .Top = 918: .Left = 565.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox10   '22
   .Height = 12: .Width = 12: .Top = 969: .Left = 565.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CommandButton4   '23
   .Height = 35.25: .Width = 83.25: .Top = 63: .Left = 558: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox2   '24
   .Height = 149.25: .Width = 282: .Top = 1107.75: .Left = 538.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox13   '25
   .Height = 25.5: .Width = 111.75: .Top = 142.5: .Left = 3.75
End With
With Worksheets("Eingabe").ComboBox14   '26
   .Height = 19.5: .Width = 172.5: .Top = 80.25: .Left = 355.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox11   '27
   .Height = 12: .Width = 12: .Top = 943.5: .Left = 565.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox12   '28
   .Height = 12: .Width = 12: .Top = 1045.5: .Left = 565.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox4   '29
   .Height = 19.5: .Width = 42: .Top = 200.25: .Left = 344.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox5   '30
   .Height = 19.5: .Width = 42: .Top = 200.25: .Left = 396.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").Label1   '31
   .Height = 13.5: .Width = 95.25: .Top = 105: .Left = 117.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox15   '32
   .Height = 19.5: .Width = 93.75: .Top = 117: .Left = 117: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox16   '33
   .Height = 19.5: .Width = 174.75: .Top = 116.25: .Left = 246: .FontSize = fntSize
End With
With Worksheets("Eingabe").Label2   '34
   .Height = 15.75: .Width = 83.25: .Top = 105: .Left = 246: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox6   '35
   .Height = 19.5: .Width = 24: .Top = 117: .Left = 213.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox21   '36
   .Height = 87: .Width = 195.75: .Top = 648: .Left = 553.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox21   '37
   .Height = 19.5: .Width = 95.25: .Top = 201: .Left = 246: .FontSize = fntSize
End With
With Worksheets("Eingabe").Label3   '38
   .Height = 13.5: .Width = 54: .Top = 84: .Left = 162.75: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox1   '39
   .Height = 15.75: .Width = 65.25: .Top = 54: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox2   '40
   .Height = 17.25: .Width = 65.25: .Top = 137.25: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox3   '41
   .Height = 17.25: .Width = 65.25: .Top = 221.25: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox4   '42
   .Height = 17.25: .Width = 65.25: .Top = 304.5: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").ComboBox1   '43
   .Height = 18.75: .Width = 237: .Top = 21.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Schneiden").TextBox1   '44
   .Height = 129.75: .Width = 578.25: .Top = 678: .Left = 5.25: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox1   '45
   .Height = 17.25: .Width = 174.75: .Top = 48.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox2   '46
   .Height = 17.25: .Width = 174.75: .Top = 295.5: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox3   '47
   .Height = 17.25: .Width = 174.75: .Top = 526.5: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox4   '48
   .Height = 17.25: .Width = 174.75: .Top = 757.5: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox1   '49
   .Height = 18: .Width = 60.75: .Top = 88.5: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox2   '50
   .Height = 18: .Width = 60.75: .Top = 322.5: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox3   '51
   .Height = 18: .Width = 60.75: .Top = 553.5: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox4   '52
   .Height = 18: .Width = 60.75: .Top = 784.5: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox5   '53
   .Height = 17.25: .Width = 174.75: .Top = 66: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox6   '54
   .Height = 17.25: .Width = 174.75: .Top = 313.5: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox7   '55
   .Height = 17.25: .Width = 174.75: .Top = 543.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox8   '56
   .Height = 17.25: .Width = 174.75: .Top = 774.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").TextBox1   '57
   .Height = 129.75: .Width = 588.75: .Top = 1254.75: .Left = 5.25: .FontSize = fntSize
End With
With Worksheets("Kleben").CheckBox1   '58
   .Height = 17.25: .Width = 72.75: .Top = 32.25: .Left = 362.25: .FontSize = fntSize
End With
With Worksheets("Kleben").ComboBox1   '59
   .Height = 17.25: .Width = 108.75: .Top = 67.5: .Left = 102.75: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox1   '60
   .Height = 72: .Width = 200: .Top = 66.75: .Left = 768: .FontSize = fntSize
End With
With Worksheets("Kleben").CheckBox2   '61
   .Height = 17.25: .Width = 72.75: .Top = 492: .Left = 360: .FontSize = fntSize
End With
With Worksheets("Kleben").ComboBox2   '62
   .Height = 17.25: .Width = 203.25: .Top = 33: .Left = 104.25: .FontSize = fntSize
End With
With Worksheets("Kleben").ComboBox3   '63
   .Height = 17.25: .Width = 174.75: .Top = 234.75: .Left = 33.75: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox2   '64
   .Height = 129.75: .Width = 636: .Top = 903: .Left = 4.5: .FontSize = fntSize
End With
With Worksheets("Kleben").ComboBox4   '65
   .Height = 17.25: .Width = 286.5: .Top = 66.75: .Left = 213: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox3   '66
   .Height = 19.5: .Width = 42: .Top = 85.5: .Left = 342: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox4   '67
   .Height = 19.5: .Width = 42: .Top = 85.5: .Left = 396: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox5   '68
   .Height = 19.5: .Width = 42: .Top = 85.5: .Left = 447: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox6   '69
   .Height = 19.5: .Width = 62.25: .Top = 85.5: .Left = 524.25: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox7   '70
   .Height = 19.5: .Width = 37.5: .Top = 65.25: .Left = 504: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox8   '71
   .Height = 19.5: .Width = 27.75: .Top = 53.25: .Left = 213.75: .FontSize = fntSize
End With
With Worksheets("Kleben").CommandButton1   '72
   .Height = 35.25: .Width = 83.25: .Top = 30.75: .Left = 664.5: .FontSize = fntSize
End With
With Worksheets("Kleben").CommandButton21   '73
   .Height = 19.5: .Width = 39: .Top = 156.75: .Left = 343.5: .FontSize = fntSize
End With
With Worksheets("Zusammentragen").ComboBox2   '74
   .Height = 17.25: .Width = 203.25: .Top = 62.25: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Zusammentragen").CheckBox1   '75
   .Height = 14.25: .Width = 90: .Top = 86.25: .Left = 4.5: .FontSize = fntSize
End With
With Worksheets("Zusammentragen").TextBox1   '76
   .Height = 103.5: .Width = 630: .Top = 719.25: .Left = 6: .FontSize = fntSize
End With
With Worksheets("Zusammentragen").CheckBox2   '77
   .Height = 17.25: .Width = 72.75: .Top = 348: .Left = 123.75: .FontSize = fntSize
End With
With Worksheets("Fadenheften").TextBox1   '78
   .Height = 129.75: .Width = 590.25: .Top = 581.75: .Left = 3.75: .FontSize = fntSize
End With
With Worksheets("Basis_Kolbus").TextBox1   '79
   .Height = 129.75: .Width = 623.25: .Top = 985.5: .Left = 3.75: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox1   '80
   .Height = 15.75: .Width = 107.25: .Top = 44.25: .Left = 291: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox2   '81
   .Height = 15.75: .Width = 107.25: .Top = 60: .Left = 291: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox3   '82
   .Height = 15.75: .Width = 107.25: .Top = 75.75: .Left = 291: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox4   '83
   .Height = 15.75: .Width = 107.25: .Top = 237.75: .Left = 291: .FontSize = fntSize
End With
With Worksheets("Binden").CheckBox3   '84
   .Height = 14.25: .Width = 168.75: .Top = 280.25: .Left = 663.75: .FontSize = fntSize
End With
With Worksheets("Binden").CommandButton1   '85
   .Height = 19.5: .Width = 39: .Top = 105: .Left = 249: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox5   '87
   .Height = 15.75: .Width = 107.25: .Top = 28.5: .Left = 291.75: .FontSize = fntSize
End With
With Worksheets("Buchdecke").ComboBox1   '88
   .Height = 19.5: .Width = 123: .Top = 83.25: .Left = 348.75: .FontSize = fntSize
End With
With Worksheets("Buchdecke").Label1   '89
   .Height = 24: .Width = 18.75: .Top = 500.25: .Left = 295.5: .FontSize = fntSize
End With
With Worksheets("Buchdecke").Label2   '90
   .Height = 12: .Width = 57: .Top = 505.5: .Left = 153: .FontSize = fntSize
End With
With Worksheets("Buchdecke").Label3   '91
   .Height = 11.25: .Width = 51: .Top = 543.75: .Left = 109.5: .FontSize = fntSize
End With
With Worksheets("Buchdecke").Label4   '92
   .Height = 12: .Width = 59.25: .Top = 432.75: .Left = 287.25: .FontSize = fntSize
End With
With Worksheets("Buchdecke").Label5   '93
   .Height = 24: .Width = 36: .Top = 534.75: .Left = 582: .FontSize = fntSize
End With
With Worksheets("Buchdecke").Label6   '94
   .Height = 12: .Width = 57.75: .Top = 577.5: .Left = 425.25: .FontSize = fntSize
End With
With Worksheets("Buchdecke").Label7   '95
   .Height = 24: .Width = 36: .Top = 524.25: .Left = 503.25: .FontSize = fntSize
End With
With Worksheets("Buchdecke").Label8   '96
   .Height = 24: .Width = 18.75: .Top = 539.25: .Left = 261.75: .FontSize = fntSize
End With
With Worksheets("Buchdecke").Label9   '97
   .Height = 24: .Width = 18.75: .Top = 541.5: .Left = 329.25: .FontSize = fntSize
End With
With Worksheets("Buchdecke").CommandButton1   '98
   .Height = 35.25: .Width = 83.25: .Top = 63: .Left = 10.5: .FontSize = fntSize
End With
With Worksheets("Buchdecke").TextBox21   '99
   .Height = 55.5: .Width = 165: .Top = 84: .Left = 475.5: .FontSize = fntSize
End With
With Worksheets("Einhängen").CheckBox1   '100
   .Height = 17.25: .Width = 82.5: .Top = 118.5: .Left = 99: .FontSize = fntSize
End With
With Worksheets("Einhängen").CheckBox2   '101
   .Height = 14.25: .Width = 82.5: .Top = 186: .Left = 97.5: .FontSize = fntSize
End With
With Worksheets("Einhängen").CheckBox3   '102
   .Height = 14.25: .Width = 129: .Top = 156: .Left = 188.25: .FontSize = fntSize
End With
With Worksheets("Einhängen").CheckBox4   '103
   .Height = 14.25: .Width = 82.5: .Top = 29.25: .Left = 99: .FontSize = fntSize
End With
With Worksheets("Einhängen").CommandButton1   '104
   .Height = 35.25: .Width = 83.25: .Top = 65.25: .Left = 6.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").Label1   '105
   .Height = 146.25: .Width = 89.25: .Top = 87: .Left = 7.5: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox1   '106
   .Height = 21: .Width = 296.25: .Top = 303: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox2   '107
   .Height = 19.5: .Width = 296.25: .Top = 61.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").CheckBox1   '108
   .Height = 18: .Width = 55.5: .Top = 282.75: .Left = 345: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox3   '109
   .Height = 21: .Width = 296.25: .Top = 349.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox4   '110
   .Height = 21: .Width = 296.25: .Top = 664.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").CommandButton1   '111
   .Height = 19.5: .Width = 39: .Top = 78.75: .Left = 636: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox5   '112
   .Height = 21: .Width = 200.25: .Top = 616.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox6   '113
   .Height = 21: .Width = 94.5: .Top = 616.5: .Left = 537.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").CheckBox4   '114
   .Height = 18: .Width = 123: .Top = 459: .Left = 204.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox7   '115
   .Height = 21: .Width = 296.25: .Top = 852: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox8   '116
   .Height = 21: .Width = 296.25: .Top = 383.25: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox9   '117
   .Height = 21: .Width = 193.5: .Top = 425.25: .Left = 439.5: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox10   '118
   .Height = 21: .Width = 48: .Top = 426: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ToggleButton1   '119
   .Height = 36.75: .Width = 93.75: .Top = 567: .Left = 117: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox11   '120
   .Height = 21: .Width = 48: .Top = 426: .Left = 384: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox12   '121
   .Height = 21: .Width = 296.25: .Top = 404.25: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").CommandButton2   '122
   .Height = 35.25: .Width = 83.25: .Top = 81.75: .Left = 114: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox13   '123
   .Height = 21: .Width = 48: .Top = 145.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox14   '124
   .Height = 21: .Width = 48: .Top = 145.5: .Left = 384.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox15   '125
   .Height = 21: .Width = 193.5: .Top = 145.5: .Left = 439.5: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox16   '126
   .Height = 21: .Width = 96.75: .Top = 910.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox17   '127
   .Height = 21: .Width = 200.25: .Top = 970.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").CheckBox2   '128
   .Height = 18: .Width = 93.75: .Top = 1017: .Left = 338.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").TextBox1   '129
   .Height = 129.75: .Width = 619.5: .Top = 1445.75: .Left = 3.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").Label2   '130
   .Height = 64.5: .Width = 76.5: .Top = 499.5: .Left = 128.25: .FontSize = fntSize
End With
With Worksheets("Produktionsdaten").ComboBox1   '131
   .Height = 18: .Width = 78.75: .Top = 974.25: .Left = 112.5: .FontSize = fntSize
End With
With Worksheets("Produktionsdaten").ComboBox2   '132
   .Height = 18: .Width = 78.75: .Top = 974.25: .Left = 193.5: .FontSize = fntSize
End With
With Worksheets("Material").CommandButton1   '133
   .Height = 25.5: .Width = 145.5: .Top = 0.75: .Left = 110.25: .FontSize = fntSize
End With
End Sub
