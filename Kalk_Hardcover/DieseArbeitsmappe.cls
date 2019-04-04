VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub Workbook_BeforePrint(Cancel As Boolean)
    Dim Letzte As Long
    Dim i As Integer
    Letzte = Worksheets("Eingabe").Cells(Rows.Count, 1).End(xlUp).Row
    Worksheets("Eingabe").PageSetup.PrintArea = "A1:L" & Letzte
    Application.ScreenUpdating = False
    Dim wks As Worksheet
    Dim v As String
    v = Worksheets("Steuerung").Range("B178") & "/" & Worksheets("Steuerung").Range("A178")
    For i = 1 To Sheets.Count
    ActiveWorkbook.Sheets(i).PageSetup.LeftFooter = "&""Calibri""&06" & Application.UserName & " " & format("&D", "dd.mm.yy") & " " & "&T"
    ActiveWorkbook.Sheets(i).PageSetup.CenterFooter = "&""Calibri""&06" & Chr(10) & ThisWorkbook.Path & "\" & ThisWorkbook.Name
    ActiveWorkbook.Sheets(i).PageSetup.RightFooter = "&""Calibri""&06" & " (v" & v & ")"
    Next
    Call Dokumenteigenschaften_Soll
    Application.ScreenUpdating = True
End Sub
Private Sub Workbook_Open()
ActiveWorkbook.CheckCompatibility = False
On Error Resume Next
fntSize = 10
With Worksheets("Eingabe").CheckBox1   '1
   .Height = 18.75: .Width = 57.75: .Top = 317.25: .Left = 9.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox2   '2
   .Height = 18.75: .Width = 57.75: .Top = 429.75: .Left = 9.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox3   '3
   .Height = 18.75: .Width = 57.75: .Top = 543.75: .Left = 9.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox4   '4
   .Height = 12: .Width = 12: .Top = 865.5: .Left = 565.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox5   '5
   .Height = 12: .Width = 12: .Top = 891: .Left = 565.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox6   '6
   .Height = 12: .Width = 12: .Top = 993: .Left = 565.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox7   '7
   .Height = 12: .Width = 12: .Top = 1069.5: .Left = 565.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox2   '8
   .Height = 19.5: .Width = 86.25: .Top = 82.5: .Left = 3.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").CommandButton2   '9
   .Height = 23.25: .Width = 73.5: .Top = 45.75: .Left = 535.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox3   '10
   .Height = 19.5: .Width = 45: .Top = 82.5: .Left = 117: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox5   '11
   .Height = 19.5: .Width = 86.25: .Top = 81.75: .Left = 246.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox1   '12
   .Height = 19.5: .Width = 113.25: .Top = 199.5: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox7   '13
   .Height = 19.5: .Width = 113.25: .Top = 353.25: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox8   '14
   .Height = 19.5: .Width = 113.25: .Top = 466.5: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox6   '15
   .Height = 19.5: .Width = 113.25: .Top = 579.75: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox4   '16
   .Height = 19.5: .Width = 113.25: .Top = 181.5: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox9   '17
   .Height = 19.5: .Width = 113.25: .Top = 334.5: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox10   '18
   .Height = 19.5: .Width = 113.25: .Top = 447.75: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox11   '19
   .Height = 19.5: .Width = 113.25: .Top = 561.75: .Left = 2.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox8   '20
   .Height = 12: .Width = 12: .Top = 1018.5: .Left = 565.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox9   '21
   .Height = 12: .Width = 12: .Top = 916.5: .Left = 565.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox10   '22
   .Height = 12: .Width = 12: .Top = 967.5: .Left = 565.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CommandButton4   '23
   .Height = 35.25: .Width = 83.25: .Top = 150: .Left = 628.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox2   '24
   .Height = 149.25: .Width = 282: .Top = 1106.25: .Left = 538.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox13   '25
   .Height = 25.5: .Width = 111.75: .Top = 143.25: .Left = 3.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox14   '26
   .Height = 19.5: .Width = 172.5: .Top = 81: .Left = 355.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox11   '27
   .Height = 12: .Width = 12: .Top = 942: .Left = 565.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox12   '28
   .Height = 12: .Width = 12: .Top = 1044: .Left = 565.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox4   '29
   .Height = 19.5: .Width = 42: .Top = 201: .Left = 344.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox5   '30
   .Height = 19.5: .Width = 42: .Top = 201: .Left = 396.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").Label1   '31
   .Height = 13.5: .Width = 95.25: .Top = 105.75: .Left = 117.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox15   '32
   .Height = 19.5: .Width = 93.75: .Top = 117.75: .Left = 117: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox16   '33
   .Height = 19.5: .Width = 174.75: .Top = 117: .Left = 246: .FontSize = fntSize
End With
With Worksheets("Eingabe").Label2   '34
   .Height = 15.75: .Width = 83.25: .Top = 105.75: .Left = 246: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox6   '35
   .Height = 19.5: .Width = 24: .Top = 117.75: .Left = 213.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox21   '36
   .Height = 87: .Width = 195.75: .Top = 648.75: .Left = 553.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox21   '37
   .Height = 19.5: .Width = 95.25: .Top = 201.75: .Left = 246: .FontSize = fntSize
End With
With Worksheets("Eingabe").Label3   '38
   .Height = 13.5: .Width = 54: .Top = 84.75: .Left = 162.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").CommandButton1   '39
   .Height = 23.25: .Width = 73.5: .Top = 73.5: .Left = 535.5: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox1   '40
   .Height = 15.75: .Width = 65.25: .Top = 54: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox2   '41
   .Height = 17.25: .Width = 65.25: .Top = 137.25: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox3   '42
   .Height = 17.25: .Width = 65.25: .Top = 221.25: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox4   '43
   .Height = 17.25: .Width = 65.25: .Top = 304.5: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").ComboBox1   '44
   .Height = 18.75: .Width = 237: .Top = 21.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Schneiden").TextBox1   '45
   .Height = 129.75: .Width = 578.25: .Top = 678: .Left = 5.25: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox1   '46
   .Height = 17.25: .Width = 174.75: .Top = 48.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox2   '47
   .Height = 17.25: .Width = 174.75: .Top = 295.5: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox3   '48
   .Height = 17.25: .Width = 174.75: .Top = 526.5: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox4   '49
   .Height = 17.25: .Width = 174.75: .Top = 757.5: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox1   '50
   .Height = 18: .Width = 60.75: .Top = 88.5: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox2   '51
   .Height = 18: .Width = 60.75: .Top = 322.5: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox3   '52
   .Height = 18: .Width = 60.75: .Top = 553.5: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox4   '53
   .Height = 18: .Width = 60.75: .Top = 784.5: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox5   '54
   .Height = 17.25: .Width = 174.75: .Top = 66: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox6   '55
   .Height = 17.25: .Width = 174.75: .Top = 313.5: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox7   '56
   .Height = 17.25: .Width = 174.75: .Top = 543.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox8   '57
   .Height = 17.25: .Width = 174.75: .Top = 774.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").TextBox1   '58
   .Height = 129.75: .Width = 588.75: .Top = 1254.75: .Left = 5.25: .FontSize = fntSize
End With
With Worksheets("Kleben").CheckBox1   '59
   .Height = 17.25: .Width = 72.75: .Top = 33: .Left = 362.25: .FontSize = fntSize
End With
With Worksheets("Kleben").ComboBox1   '60
   .Height = 17.25: .Width = 108.75: .Top = 54: .Left = 102.75: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox1   '61
   .Height = 72: .Width = 200.25: .Top = 67.5: .Left = 768: .FontSize = fntSize
End With
With Worksheets("Kleben").CheckBox2   '62
   .Height = 17.25: .Width = 72.75: .Top = 408.75: .Left = 360: .FontSize = fntSize
End With
With Worksheets("Kleben").ComboBox2   '63
   .Height = 17.25: .Width = 203.25: .Top = 33.75: .Left = 104.25: .FontSize = fntSize
End With
With Worksheets("Kleben").ComboBox3   '64
   .Height = 17.25: .Width = 174.75: .Top = 151.5: .Left = 33.75: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox2   '65
   .Height = 129.75: .Width = 636: .Top = 903.75: .Left = 4.5: .FontSize = fntSize
End With
With Worksheets("Kleben").ComboBox4   '66
   .Height = 17.25: .Width = 286.5: .Top = 67.5: .Left = 213: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox3   '67
   .Height = 19.5: .Width = 42: .Top = 86.25: .Left = 342: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox4   '68
   .Height = 19.5: .Width = 42: .Top = 86.25: .Left = 396: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox5   '69
   .Height = 19.5: .Width = 42: .Top = 86.25: .Left = 447: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox6   '70
   .Height = 19.5: .Width = 62.25: .Top = 86.25: .Left = 524.25: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox7   '71
   .Height = 19.5: .Width = 37.5: .Top = 66: .Left = 504: .FontSize = fntSize
End With
With Worksheets("Kleben").TextBox8   '72
   .Height = 19.5: .Width = 27.75: .Top = 53.25: .Left = 213.75: .FontSize = fntSize
End With
With Worksheets("Kleben").CommandButton1   '73
   .Height = 35.25: .Width = 83.25: .Top = 31.5: .Left = 664.5: .FontSize = fntSize
End With
With Worksheets("Kleben").CommandButton21   '74
   .Height = 19.5: .Width = 39: .Top = 157.5: .Left = 343.5: .FontSize = fntSize
End With
With Worksheets("Zusammentragen").ComboBox2   '75
   .Height = 17.25: .Width = 203.25: .Top = 62.25: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Zusammentragen").CheckBox1   '76
   .Height = 14.25: .Width = 90: .Top = 86.25: .Left = 4.5: .FontSize = fntSize
End With
With Worksheets("Zusammentragen").TextBox1   '77
   .Height = 103.5: .Width = 630: .Top = 719.25: .Left = 6: .FontSize = fntSize
End With
With Worksheets("Zusammentragen").CheckBox2   '78
   .Height = 17.25: .Width = 72.75: .Top = 348: .Left = 123.75: .FontSize = fntSize
End With
With Worksheets("Fadenheften").TextBox1   '79
   .Height = 129.75: .Width = 590.25: .Top = 582: .Left = 3.75: .FontSize = fntSize
End With
With Worksheets("Fadenheften").CheckBox1   '80
   .Height = 19.5: .Width = 60.75: .Top = 141: .Left = 409.5: .FontSize = fntSize
End With
With Worksheets("Basis_Kolbus").TextBox1   '81
   .Height = 129.75: .Width = 623.25: .Top = 986.25: .Left = 3.75: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox1   '82
   .Height = 15.75: .Width = 107.25: .Top = 45: .Left = 291: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox2   '83
   .Height = 15.75: .Width = 107.25: .Top = 60.75: .Left = 291: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox3   '84
   .Height = 15.75: .Width = 107.25: .Top = 76.5: .Left = 291: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox4   '85
   .Height = 15.75: .Width = 107.25: .Top = 238.5: .Left = 291: .FontSize = fntSize
End With
With Worksheets("Binden").CheckBox3   '86
   .Height = 14.25: .Width = 168.75: .Top = 281.25: .Left = 663.75: .FontSize = fntSize
End With
With Worksheets("Binden").CommandButton1   '87
   .Height = 19.5: .Width = 39: .Top = 105.75: .Left = 249: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox5   '88
   .Height = 15.75: .Width = 107.25: .Top = 29.25: .Left = 291.75: .FontSize = fntSize
End With
With Worksheets("Buchdecke").ComboBox1   '89
   .Height = 19.5: .Width = 123: .Top = 84: .Left = 348.75: .FontSize = fntSize
End With
With Worksheets("Buchdecke").Label1   '90
   .Height = 24: .Width = 18.75: .Top = 893.25: .Left = 306.75: .FontSize = fntSize
End With
With Worksheets("Buchdecke").Label2   '91
   .Height = 12: .Width = 57: .Top = 1134.75: .Left = 120: .FontSize = fntSize
End With
With Worksheets("Buchdecke").Label3   '92
   .Height = 11.25: .Width = 51: .Top = 1111.5: .Left = 176.25: .FontSize = fntSize
End With
With Worksheets("Buchdecke").Label4   '93
   .Height = 12: .Width = 59.25: .Top = 831: .Left = 306: .FontSize = fntSize
End With
With Worksheets("Buchdecke").Label5   '94
   .Height = 24: .Width = 36: .Top = 1061.25: .Left = 630: .FontSize = fntSize
End With
With Worksheets("Buchdecke").Label6   '95
   .Height = 12: .Width = 57.75: .Top = 1221: .Left = 466.5: .FontSize = fntSize
End With
With Worksheets("Buchdecke").Label7   '96
   .Height = 24: .Width = 36: .Top = 1187.25: .Left = 523.5: .FontSize = fntSize
End With
With Worksheets("Buchdecke").Label8   '97
   .Height = 24: .Width = 18.75: .Top = 968.25: .Left = 297: .FontSize = fntSize
End With
With Worksheets("Buchdecke").Label9   '98
   .Height = 24: .Width = 18.75: .Top = 968.25: .Left = 324.75: .FontSize = fntSize
End With
With Worksheets("Buchdecke").CommandButton1   '99
   .Height = 35.25: .Width = 83.25: .Top = 63.75: .Left = 10.5: .FontSize = fntSize
End With
With Worksheets("Buchdecke").TextBox21   '100
   .Height = 55.5: .Width = 165: .Top = 84.75: .Left = 475.5: .FontSize = fntSize
End With
With Worksheets("Einhängen").CheckBox1   '101
   .Height = 17.25: .Width = 82.5: .Top = 119.25: .Left = 99: .FontSize = fntSize
End With
With Worksheets("Einhängen").CheckBox2   '102
   .Height = 14.25: .Width = 82.5: .Top = 186.75: .Left = 97.5: .FontSize = fntSize
End With
With Worksheets("Einhängen").CheckBox3   '103
   .Height = 14.25: .Width = 129: .Top = 156.75: .Left = 188.25: .FontSize = fntSize
End With
With Worksheets("Einhängen").CheckBox4   '104
   .Height = 14.25: .Width = 82.5: .Top = 30: .Left = 99: .FontSize = fntSize
End With
With Worksheets("Einhängen").CommandButton1   '105
   .Height = 35.25: .Width = 83.25: .Top = 66: .Left = 6.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").Label1   '106
   .Height = 146.25: .Width = 89.25: .Top = 87: .Left = 7.5: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox1   '107
   .Height = 21: .Width = 296.25: .Top = 303: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox2   '108
   .Height = 19.5: .Width = 296.25: .Top = 61.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").CheckBox1   '109
   .Height = 18: .Width = 55.5: .Top = 282.75: .Left = 345: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox3   '110
   .Height = 21: .Width = 296.25: .Top = 349.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox4   '111
   .Height = 21: .Width = 296.25: .Top = 664.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").CommandButton1   '112
   .Height = 19.5: .Width = 39: .Top = 78.75: .Left = 636: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox5   '113
   .Height = 21: .Width = 200.25: .Top = 616.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox6   '114
   .Height = 21: .Width = 94.5: .Top = 616.5: .Left = 537.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").CheckBox4   '115
   .Height = 18: .Width = 123: .Top = 459: .Left = 204.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox7   '116
   .Height = 21: .Width = 296.25: .Top = 852: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox8   '117
   .Height = 21: .Width = 296.25: .Top = 383.25: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox9   '118
   .Height = 21: .Width = 193.5: .Top = 425.25: .Left = 439.5: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox10   '119
   .Height = 21: .Width = 48: .Top = 426: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ToggleButton1   '120
   .Height = 36.75: .Width = 93.75: .Top = 567: .Left = 117: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox11   '121
   .Height = 21: .Width = 48: .Top = 426: .Left = 384: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox12   '122
   .Height = 21: .Width = 296.25: .Top = 404.25: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").CommandButton2   '123
   .Height = 35.25: .Width = 83.25: .Top = 81.75: .Left = 114: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox13   '124
   .Height = 21: .Width = 48: .Top = 145.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox14   '125
   .Height = 21: .Width = 48: .Top = 145.5: .Left = 384.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox15   '126
   .Height = 21: .Width = 193.5: .Top = 145.5: .Left = 439.5: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox16   '127
   .Height = 21: .Width = 96.75: .Top = 910.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox17   '128
   .Height = 21: .Width = 200.25: .Top = 970.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").CheckBox2   '129
   .Height = 18: .Width = 93.75: .Top = 1017: .Left = 338.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").TextBox1   '130
   .Height = 129.75: .Width = 619.5: .Top = 1446: .Left = 3.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").Label2   '131
   .Height = 64.5: .Width = 76.5: .Top = 499.5: .Left = 128.25: .FontSize = fntSize
End With
With Worksheets("Produktionsdaten").ComboBox1   '132
   .Height = 18: .Width = 78.75: .Top = 981: .Left = 112.5: .FontSize = fntSize
End With
With Worksheets("Produktionsdaten").ComboBox2   '133
   .Height = 18: .Width = 78.75: .Top = 981: .Left = 193.5: .FontSize = fntSize
End With
With Worksheets("Material").CommandButton1   '134
   .Height = 25.5: .Width = 145.5: .Top = 0.75: .Left = 110.25: .FontSize = fntSize
End With
'Leimart
Worksheets("Eingabe").ComboBox13.FontSize = 16
End Sub
