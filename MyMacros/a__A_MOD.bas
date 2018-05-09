Attribute VB_Name = "a__A_MOD"
'
' This MODULE is for testing out functions
'

Sub CurrentRegion()
    Set aRange = ActiveCell.CurrentRegion
    aRange.Select
End Sub

Function NaCount(Optional useRange)
Dim useRow As Long, useCol As Integer
Dim aRange As Range

    If IsMissing(useRange) Then
        useCol = ActiveCell.Column
        botRow = ColumnLastRow(useCol)
        Set useRange = Range(Cells(2, useCol), Cells(botRow, useCol))
    End If
    
    t = useRange.SpecialCells(xlCellTypeVisible, xlErrors).count
    
    t = bRange.count
    
End Function
