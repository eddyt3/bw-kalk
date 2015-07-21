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
   .Height = 12: .Width = 12: .Top = 754.5: .Left = 574.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox2   '2
   .Height = 12: .Width = 12: .Top = 780: .Left = 574.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox3   '3
   .Height = 12: .Width = 12: .Top = 805.5: .Left = 574.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox4   '4
   .Height = 12: .Width = 12: .Top = 856.5: .Left = 574.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox5   '5
   .Height = 18.75: .Width = 57.75: .Top = 234: .Left = 9.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox6   '6
   .Height = 18.75: .Width = 57.75: .Top = 359.25: .Left = 9.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox7   '7
   .Height = 18.75: .Width = 57.75: .Top = 485.25: .Left = 9.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").CommandButton3   '8
   .Height = 23.25: .Width = 73.5: .Top = 33.75: .Left = 16.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox1   '9
   .Height = 19.5: .Width = 113.25: .Top = 105: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox2   '10
   .Height = 19.5: .Width = 113.25: .Top = 251.25: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox3   '11
   .Height = 19.5: .Width = 113.25: .Top = 378: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox4   '12
   .Height = 19.5: .Width = 113.25: .Top = 504: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").CheckBox8   '13
   .Height = 12: .Width = 12: .Top = 831: .Left = 574.5: .FontSize = fntSize
End With
With Worksheets("Eingabe").ComboBox5   '14
   .Height = 19.5: .Width = 97.5: .Top = 118.5: .Left = 255: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox2   '15
   .Height = 19.5: .Width = 42: .Top = 117.75: .Left = 354.75: .FontSize = fntSize
End With
With Worksheets("Eingabe").TextBox3   '16
   .Height = 19.5: .Width = 42: .Top = 117.75: .Left = 409.5: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox1   '17
   .Height = 15.75: .Width = 65.25: .Top = 54: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox2   '18
   .Height = 15.75: .Width = 65.25: .Top = 137.25: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox3   '19
   .Height = 15.75: .Width = 65.25: .Top = 221.25: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").CheckBox4   '20
   .Height = 15.75: .Width = 65.25: .Top = 304.5: .Left = 89.25: .FontSize = fntSize
End With
With Worksheets("Schneiden").ComboBox1   '21
   .Height = 18.75: .Width = 237: .Top = 21.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox1   '22
   .Height = 17.25: .Width = 174.75: .Top = 48.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox2   '23
   .Height = 17.25: .Width = 174.75: .Top = 297: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox3   '24
   .Height = 17.25: .Width = 174.75: .Top = 528: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox4   '25
   .Height = 17.25: .Width = 174.75: .Top = 759: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox1   '26
   .Height = 18: .Width = 60.75: .Top = 88.5: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox2   '27
   .Height = 18: .Width = 60.75: .Top = 324: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox3   '28
   .Height = 18: .Width = 60.75: .Top = 555: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").CheckBox4   '29
   .Height = 18: .Width = 60.75: .Top = 786: .Left = 618: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox5   '30
   .Height = 17.25: .Width = 174.75: .Top = 66: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox6   '31
   .Height = 17.25: .Width = 174.75: .Top = 315: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox7   '32
   .Height = 17.25: .Width = 174.75: .Top = 545.25: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Falzen").ComboBox8   '33
   .Height = 17.25: .Width = 174.75: .Top = 776.25: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Sammelheften").ComboBox1   '34
   .Height = 17.25: .Width = 69.75: .Top = 74.25: .Left = 20.25: .FontSize = fntSize
End With
With Worksheets("Sammelheften").ComboBox2   '35
   .Height = 17.25: .Width = 97.5: .Top = 59.25: .Left = 318: .FontSize = fntSize
End With
With Worksheets("Sammelheften").CommandButton1   '36
   .Height = 19.5: .Width = 39: .Top = 43.5: .Left = 616.5: .FontSize = fntSize
End With
With Worksheets("Bohren").ComboBox1   '37
   .Height = 18.75: .Width = 252.75: .Top = 21.75: .Left = 1.5: .FontSize = fntSize
End With
With Worksheets("Bohren").CheckBox2   '38
   .Height = 15.75: .Width = 65.25: .Top = 121.5: .Left = 585.75: .FontSize = fntSize
End With
With Worksheets("Bohren").CheckBox4   '39
   .Height = 15.75: .Width = 65.25: .Top = 54.75: .Left = 107.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").Label1   '40
   .Height = 146.25: .Width = 74.25: .Top = 87: .Left = 14.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox1   '41
   .Height = 21: .Width = 296.25: .Top = 302.25: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox2   '42
   .Height = 19.5: .Width = 296.25: .Top = 61.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").CheckBox1   '43
   .Height = 18: .Width = 55.5: .Top = 282: .Left = 345: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox3   '44
   .Height = 21: .Width = 296.25: .Top = 348.75: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox4   '45
   .Height = 21: .Width = 296.25: .Top = 663.75: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").CommandButton1   '46
   .Height = 19.5: .Width = 39: .Top = 78.75: .Left = 636: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox5   '47
   .Height = 21: .Width = 200.25: .Top = 615.75: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox6   '48
   .Height = 21: .Width = 94.5: .Top = 615.75: .Left = 537.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").CheckBox4   '49
   .Height = 18: .Width = 111: .Top = 457.5: .Left = 204.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox7   '50
   .Height = 21: .Width = 296.25: .Top = 851.25: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox8   '51
   .Height = 21: .Width = 296.25: .Top = 382.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox9   '52
   .Height = 21: .Width = 193.5: .Top = 424.5: .Left = 439.5: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox10   '53
   .Height = 21: .Width = 48: .Top = 425.25: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ToggleButton1   '54
   .Height = 36.75: .Width = 93.75: .Top = 566.25: .Left = 117: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox11   '55
   .Height = 21: .Width = 48: .Top = 425.25: .Left = 384: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox12   '56
   .Height = 21: .Width = 296.25: .Top = 403.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").CommandButton2   '57
   .Height = 35.25: .Width = 83.25: .Top = 81.75: .Left = 114: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox13   '58
   .Height = 21: .Width = 48: .Top = 145.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox14   '59
   .Height = 21: .Width = 48: .Top = 145.5: .Left = 384.75: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox15   '60
   .Height = 21: .Width = 193.5: .Top = 145.5: .Left = 439.5: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox16   '61
   .Height = 21: .Width = 96.75: .Top = 910.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").ComboBox17   '62
   .Height = 21: .Width = 199.5: .Top = 970.5: .Left = 336: .FontSize = fntSize
End With
With Worksheets("Verpacken").CheckBox2   '63
   .Height = 18: .Width = 93.75: .Top = 1016.25: .Left = 338.25: .FontSize = fntSize
End With
With Worksheets("Verpacken").Label2   '64
   .Height = 64.5: .Width = 76.5: .Top = 492: .Left = 129: .FontSize = fntSize
End With
With Worksheets("Produktionsdaten").ComboBox1   '65
   .Height = 18: .Width = 78.75: .Top = 882.75: .Left = 112.5: .FontSize = fntSize
End With
With Worksheets("Produktionsdaten").ComboBox2   '66
   .Height = 18: .Width = 78.75: .Top = 882.75: .Left = 193.5: .FontSize = fntSize
End With
With Worksheets("Material").CommandButton1   '67
   .Height = 39.75: .Width = 86.25: .Top = 917.25: .Left = 3: .FontSize = fntSize
End With
End Sub
