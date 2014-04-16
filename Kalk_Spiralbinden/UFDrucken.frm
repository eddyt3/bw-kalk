VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFDrucken 
   Caption         =   "Drucken"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2400
   OleObjectBlob   =   "UFDrucken.frx":0000
   StartUpPosition =   1  'Fenstermitte
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
If CheckBox2 Then Sheets("Beschneiden").PrintOut
If CheckBox3 Then Sheets("Zusammentragen").PrintOut
If CheckBox4 Then Sheets("Stanzen").PrintOut
If CheckBox5 Then Sheets("Binden").PrintOut
If CheckBox6 Then Sheets("Verpacken").PrintOut
If CheckBox7 Then Sheets("Produktionsdaten").PrintOut
If CheckBox11 Then Sheets(Array("Eingabe", "Beschneiden", "Zusammentragen", "Stanzen", "Binden", _
    "Verpacken", "Produktiondaten")).PrintOut
End Sub
Private Sub UserForm_Initialize()
'Voreinstellungen:
CheckBox1 = True
CheckBox7 = True
End Sub
