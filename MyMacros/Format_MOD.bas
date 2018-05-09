Attribute VB_Name = "Format_MOD"
Sub FmtDate()
    useCol = ActiveCell.Column
    botRow = ColumnLastRow(useCol)
    Set aRange = Range(Cells(2, useCol), Cells(botRow, useCol))
    aRange.NumberFormat = "m/d/yyyy"
End Sub

Sub FmtTime()
    useCol = ActiveCell.Column
    botRow = ColumnLastRow(useCol)
    Set aRange = Range(Cells(2, useCol), Cells(botRow, useCol))
    aRange.NumberFormat = "[$-F400]h:mm:ss AM/PM"
End Sub

Sub FmtGenrl()
    useCol = ActiveCell.Column
    botRow = ColumnLastRow(useCol)
    Set aRange = Range(Cells(2, useCol), Cells(botRow, useCol))
    aRange.NumberFormat = "General"
End Sub

Sub FmtNumCommas()
    Selection.NumberFormat = "#,##0"
End Sub

Sub FmtDtTm()
    useCol = ActiveCell.Column
    botRow = ColumnLastRow(useCol)
    Set aRange = Range(Cells(2, useCol), Cells(botRow, useCol))
    aRange.NumberFormat = "m/d/yy h:mm;@"
End Sub

Sub WhatisTheFmt()
    For Each c In Cells.SpecialCells(xlCellTypeConstants)
        MsgBox c.Interior.ColorIndex
    Next c
End Sub
