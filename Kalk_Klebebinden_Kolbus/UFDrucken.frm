VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFDrucken 
   Caption         =   "Drucken"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2400
   OleObjectBlob   =   "UFDrucken.frx":0000
End
Attribute VB_Name = "UFDrucken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AbbrechenButton_Click()
Unload Me
End Sub

Private Sub DruckenButton_Click()
'Blattnamen jeweils anpassen!
Dim a As Integer
Dim cb As CheckBox
If CheckBox1 Then Sheets("Eingabe").PrintOut
If CheckBox2 Then Sheets("Schneiden").PrintOut
If CheckBox3 Then Sheets("Zusammentragen").PrintOut
If CheckBox4 Then Sheets("Falzen").PrintOut
If CheckBox5 Then Sheets("Kleben").PrintOut
If CheckBox6 Then Sheets("Fadenheften").PrintOut
If CheckBox7 Then Sheets("Binden").PrintOut
If CheckBox9 Then Sheets("Verpacken").PrintOut
If CheckBox10 Then Sheets("Produktionsdaten").PrintOut
If CheckBox11 Then Sheets(Array("Eingabe", "Schneiden", "Zusammentragen", "Falzen", "Kleben", "Fadenheften", _
    "Binden", "Verpacken", "Produktiondaten")).PrintOut
End Sub

Private Sub UserForm_Initialize()
'Voreinstellungen:
CheckBox1 = True
CheckBox10 = True
End Sub
