Attribute VB_Name = "NoDups_MOD"
Sub RemoveColumnDups(Optional aRange)
Dim useRange As Range

    If IsMissing(aRange) Then Set aRange = ActiveCell
    useCol = aRange.Column
    topRow = 1
    If HasHeader Then topRow = 2
    
    Set useRange = Range(Cells(topRow, useCol), Cells(LastRow, useCol))
    totalCount = WorksheetFunction.CountA(useRange)
    
    Debug.Print useRange.Address
    useRange.RemoveDuplicates Columns:=1, Header:=xlNo
    uniqueCount = WorksheetFunction.CountA(useRange)
    If (uniqueCount = useRange.count) Then
        MsgBox "All Unique!"
    Else
        MsgBox WorksheetFunction.CountA(useRange) & " unique values from " & useRange.count
    End If
    retCode = MsgBox("Insert Unique Count?", vbYesNo)
    If (retCode = vbYes) Then
        Cells(2, useCol).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Cells(2, useCol).Value = WorksheetFunction.CountA(useRange) & " of " & totalCount
    End If
 End Sub

Sub UniqueValues()
    useCol = ActiveCell.Column
    botRow = LastRow()
    Set useRange = Range(Cells(1, useCol), Cells(botRow, useCol))

End Sub

Sub ShowDupRow()

    SHUse = ActiveSheet.Name
    useCol = ActiveCell.Column
    useRow = 2

    Call ClearClipboard
    Cells(useRow, useCol).Offset(0, 1).FormulaR1C1 = "1"
    Cells(useRow, useCol).Offset(1, 1).FormulaR1C1 = "2"
    Cells(useRow, useCol).Offset(2, 1).FormulaR1C1 = "3"
    Set inRange = Range(Cells(useRow, useCol).Offset(0, 1), Cells(useRow, useCol).Offset(2, 1))
    botRow = Cells(Rows.count, useCol).End(xlUp).Row
    Set outRange = Range(Cells(useRow, useCol).Offset(0, 1), Cells(botRow, useCol).Offset(0, 1))
    inRange.Copy
    outRange.Copy
    inRange.AutoFill Destination:=outRange
    
    'Columns("G:H").Select
    ActiveWorkbook.Worksheets(SHUse).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(SHUse).Sort.SortFields.Add _
        Key:=Range(Cells(2, useCol), Cells(botRow, useCol)), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
    
    With ActiveWorkbook.Worksheets(SHUse).Sort
        .SetRange Range(Cells(1, useCol), Cells(botRow, useCol + 1))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Cells(useRow, useCol).Offset(1, 2).FormulaR1C1 = "=IF(R[-1]C[-2]=RC[-2],1,0)"
    Set fillRange = Range(Cells(useRow, useCol).Offset(1, 2), Cells(botRow, useCol).Offset(0, 2))
    Cells(useRow, useCol).Offset(1, 2).AutoFill Destination:=fillRange

    Debug.Print fillRange.Address
    Debug.Print "=SUM(" & fillRange.Address & ")"
    Cells(1, useCol).Offset(0, 2).Formula = "=SUM(" & fillRange.Address & ")"
    If Cells(1, useCol).Offset(0, 2).Value > 0 Then
        MsgBox "There are duplicates"
        
        Set fRange = FindInRange(1, fillRange, xlWhole)
        Debug.Print fRange.count
        Call ColorRange(fRange, "yellow")
    End If
End Sub
