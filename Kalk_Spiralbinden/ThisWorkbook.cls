VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
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
    Letzte = Cells(Rows.Count, 1).End(xlUp).Row
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
fntSize = 10
With Worksheets("Eingabe").ComboBox1   '1
   .Height = 21: .Width = 73.5: .Top = 166.5: .Left = 570.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox2   '2
   .Height = 21: .Width = 73.5: .Top = 250.5: .Left = 570.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox3   '3
   .Height = 19.5: .Width = 128.25: .Top = 168: .Left = 374.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox4   '4
   .Height = 19.5: .Width = 128.25: .Top = 189: .Left = 374.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox5   '5
   .Height = 19.5: .Width = 128.25: .Top = 210: .Left = 374.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox6   '6
   .Height = 19.5: .Width = 128.25: .Top = 231: .Left = 374.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox7   '7
   .Height = 19.5: .Width = 128.25: .Top = 252: .Left = 374.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").CommandButton1   '8
   .Height = 19.5: .Width = 51: .Top = 251.25: .Left = 73.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CommandButton3   '9
   .Height = 23.25: .Width = 73.5: .Top = 33.75: .Left = 16.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox8   '10
   .Height = 17.25: .Width = 37.5: .Top = 168: .Left = 518.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox9   '11
   .Height = 17.25: .Width = 38.25: .Top = 189: .Left = 518.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox10   '12
   .Height = 17.25: .Width = 38.25: .Top = 210: .Left = 518.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox11   '13
   .Height = 17.25: .Width = 38.25: .Top = 231: .Left = 518.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox12   '14
   .Height = 17.25: .Width = 38.25: .Top = 252: .Left = 518.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").Label1   '15
   .Height = 15: .Width = 48: .Top = 150.75: .Left = 516: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox1   '16
   .Height = 12: .Width = 12: .Top = 705.75: .Left = 602.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox2   '17
   .Height = 12: .Width = 12: .Top = 735.75: .Left = 602.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox3   '18
   .Height = 12: .Width = 12: .Top = 765.75: .Left = 602.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox4   '19
   .Height = 12: .Width = 12: .Top = 795.75: .Left = 602.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox5   '20
   .Height = 12: .Width = 12: .Top = 825.75: .Left = 602.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox6   '21
   .Height = 12: .Width = 12: .Top = 885.75: .Left = 602.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").CommandButton5   '22
   .Height = 19.5: .Width = 51: .Top = 168: .Left = 72: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox7   '23
   .Height = 12: .Width = 12: .Top = 856.5: .Left = 602.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox13   '24
   .Height = 19.5: .Width = 96: .Top = 125.25: .Left = 264.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox1   '25
   .Height = 19.5: .Width = 42: .Top = 126: .Left = 374.25: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox2   '26
   .Height = 19.5: .Width = 42: .Top = 126: .Left = 427.5: .FontSize = fntSize
End With
With Worksheets("Schneiden").ComboBox1   '27
   .Height = 18.75: .Width = 237: .Top = 21: .Left = 3: .FontSize = fntSize
End With
With Worksheets("Zusammentragen").ComboBox2   '28
   .Height = 17.25: .Width = 203.25: .Top = 55.5: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Zusammentragen").CheckBox1   '29
   .Height = 15: .Width = 91.5: .Top = 79.5: .Left = 4.5: .FontSize = fntSize
End With
With Worksheets("3Seitenbeschnitt").CheckBox1   '30
   .Height = 17.25: .Width = 72.75: .Top = 39.75: .Left = 112.5: .FontSize = fntSize
End With
With Worksheets("3Seitenbeschnitt").CheckBox2   '31
   .Height = 17.25: .Width = 74.25: .Top = 252.75: .Left = 111.75: .FontSize = fntSize
End With
With Worksheets("Binden").ComboBox1   '32
   .Height = 16.5: .Width = 248.25: .Top = 22.5: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Bohren").CheckBox1   '33
   .Height = 15.75: .Width = 65.25: .Top = 180: .Left = 107.25: .FontSize = fntSize
End With
With Worksheets("Bohren").ComboBox1   '34
   .Height = 18.75: .Width = 252.75: .Top = 21.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Bohren").CheckBox2   '35
   .Height = 15.75: .Width = 65.25: .Top = 121.5: .Left = 585.75: .FontSize = fntSize
End With
With Worksheets("Bohren").CheckBox3   '36
   .Height = 15.75: .Width = 65.25: .Top = 180: .Left = 586.5: .FontSize = fntSize
End With
With Worksheets("Bohren").CheckBox4   '37
   .Height = 15.75: .Width = 65.25: .Top = 54.75: .Left = 107.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").Label1   '38
   .Height = 146.25: .Width = 74.25: .Top = 87: .Left = 14.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox1   '39
   .Height = 21: .Width = 296.25: .Top = 303: .Left = 335.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox2   '40
   .Height = 19.5: .Width = 296.25: .Top = 61.5: .Left = 335.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").CheckBox1   '41
   .Height = 18: .Width = 55.5: .Top = 282.75: .Left = 344.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox3   '42
   .Height = 21: .Width = 296.25: .Top = 349.5: .Left = 335.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox4   '43
   .Height = 21: .Width = 296.25: .Top = 664.5: .Left = 335.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").CommandButton1   '44
   .Height = 19.5: .Width = 39: .Top = 78.75: .Left = 635.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox5   '45
   .Height = 21: .Width = 200.25: .Top = 616.5: .Left = 335.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox6   '46
   .Height = 21: .Width = 94.5: .Top = 616.5: .Left = 537: .FontSize = fntSize
End With
With Worksheets("Verpacken").CheckBox4   '47
   .Height = 18: .Width = 111: .Top = 458.25: .Left = 204.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox7   '48
   .Height = 21: .Width = 296.25: .Top = 852: .Left = 335.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox8   '49
   .Height = 21: .Width = 296.25: .Top = 383.25: .Left = 335.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox9   '50
   .Height = 21: .Width = 193.5: .Top = 425.25: .Left = 438.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox10   '51
   .Height = 21: .Width = 48: .Top = 426: .Left = 335.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").ToggleButton1   '52
   .Height = 36.75: .Width = 93.75: .Top = 567: .Left = 117: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox11   '53
   .Height = 21: .Width = 48: .Top = 426: .Left = 383.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox12   '54
   .Height = 21: .Width = 296.25: .Top = 404.25: .Left = 335.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").CommandButton2   '55
   .Height = 35.25: .Width = 83.25: .Top = 81.75: .Left = 114: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox13   '56
   .Height = 21: .Width = 48: .Top = 145.5: .Left = 335.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox14   '57
   .Height = 21: .Width = 48: .Top = 145.5: .Left = 384: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox15   '58
   .Height = 21: .Width = 193.5: .Top = 145.5: .Left = 438.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox16   '59
   .Height = 21: .Width = 96.75: .Top = 912: .Left = 335.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox17   '60
   .Height = 21: .Width = 201: .Top = 972: .Left = 335.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").CheckBox2   '61
   .Height = 18: .Width = 93.75: .Top = 1017.75: .Left = 338.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").Label2   '62
   .Height = 64.5: .Width = 76.5: .Top = 500.25: .Left = 127.5: .FontSize = fntSize
End With
With Worksheets("Produktionsdaten").ComboBox1   '63
   .Height = 18: .Width = 78.75: .Top = 1029: .Left = 132: .FontSize = fntSize
End With
With Worksheets("Produktionsdaten").ComboBox2   '64
   .Height = 18: .Width = 78.75: .Top = 1029: .Left = 213: .FontSize = fntSize
End With
With Worksheets("Material").CommandButton1   '65
   .Height = 39.75: .Width = 86.25: .Top = 789.75: .Left = 2.25: .FontSize = fntSize
End With

End Sub
