Attribute VB_Name = "Modul3"
Sub Makro1()
Attribute Makro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro1 Makro
'

'
    Range("F6:G6,D6").Select
    Range("D6").Activate
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Locked = True
    Selection.FormulaHidden = False
    Range("D6").AddComment
    Range("D6").Comment.Visible = False
    Range("D6").Comment.Text Text:="Enrico Dargel:" & Chr(10) & "x"
    Range("F6").Select
    Range("F6").AddComment
    Range("F6").Comment.Visible = False
    Range("F6").Comment.Text Text:="Enrico Dargel:" & Chr(10) & "x1"
    Range("G6").Select
    Range("G6").AddComment
    Range("G6").Comment.Visible = False
    Range("G6").Comment.Text Text:="Enrico Dargel:" & Chr(10) & "x2"
    Range("H20").Select
End Sub
Sub Makro2()
Attribute Makro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro2 Makro
'

'
    Range("F6:G6,D6").Select
    Range("D6").Activate
    Selection.ClearComments
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Locked = False
    Selection.FormulaHidden = False
    Range("H8").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    ActiveWindow.LargeScroll Down:=0
    ActiveWindow.Zoom = 82
    ActiveWindow.Zoom = 86
    ActiveWindow.Zoom = 87
    ActiveWindow.Zoom = 89
    ActiveWindow.Zoom = 100
    ActiveWindow.LargeScroll Down:=-1
    Range("F10").Select
    Sheets("Plantafel").Select
    ActiveWindow.SmallScroll Down:=18
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    Sheets("Verpacken").Select
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "16"
    Range("G6").Select
    ActiveWorkbook.Save
    Range("H9").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    Range("D6").Select
    Range("D6").Comment.Text Text:= _
        "E.Dargel:" & Chr(10) & "Einfaches Abstapeln ohne Folie od. Karton." & Chr(10) & "Bei Bedarf Zwischenlagen festlegen."
    Selection.ShapeRange.ScaleWidth 1.63, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 1.86, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleWidth 1.18, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.86, msoFalse, msoScaleFromTopLeft
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "16"
    Range("D6").Select
    ActiveCell.FormulaR1C1 = "16"
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("F7").Select
    ActiveWorkbook.Save
    Range("D6").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("F8").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    Range("H11").Select
    ActiveWorkbook.Save
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "45"
    ActiveCell.Next.Activate
    ActiveCell.FormulaR1C1 = "45"
    Range("F7").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    Range("H9").Select
    ActiveWorkbook.Save
    Range("C6").Select
    ActiveCell.FormulaR1C1 = "Exemplare/VE"
    Range("H8").Select
    ActiveWindow.SmallScroll ToRight:=-1
    Sheets("Steuerung").Select
    ActiveWorkbook.Save
    ActiveWindow.LargeScroll Down:=-4
End Sub
