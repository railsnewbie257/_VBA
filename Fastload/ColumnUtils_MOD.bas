Attribute VB_Name = "ColumnUtils_MOD"
' Routines that deal with columns
'
'
' Append daat from one column to another column
'

Function ColumnInsertRight(useCol, Optional SHUse, Optional WBUse)
    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    Call ClearClipboard
    Call CalculationOff
    Workbooks(WBUse).Worksheets(SHUse).Columns(useCol + 1).Insert Shift:=xlToRight
    Columns(useCol + 1).NumberFormat = "General"
    ColumnInsertRight = useCol + 1
    Call CalculationOn
End Function

Function ColumnInsertLeft(useCol, Optional SHUse, Optional WBUse)
    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    Call ClearClipboard
    Call CalculationOff
    Workbooks(WBUse).Worksheets(SHUse).Columns(useCol).Insert Shift:=xlToRight
    Columns(useCol).NumberFormat = "General"
    ColumnInsertLeft = useCol
    Call CalculationOn
End Function

Function ColumnLastRow(Optional useCol, Optional SHUse, Optional WBUse) ' problem ?

10  On Error GoTo gotError

20  If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
30  If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
40  If IsMissing(useCol) Then useCol = 1
50  k = LastRow(SHUse, WBUse)
    
60  With Workbooks(WBUse).Worksheets(SHUse)
70      ColumnLastRow = .Cells(k + 1, useCol).End(xlUp).Row
80      If IsEmpty(.Cells(ColumnLastRow, useCol)) Then ColumnLastRow = 0
90  End With

100 Exit Function

gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="ColumnLastRow"
    Stop
    Resume Next
End Function

Function ColumnCountA_del(useCol)
    Set aRange = range(Cells(2, useCol), Cells(ColumnLastRow(useCol), useCol))
    ColumnCountA = WorksheetFunction.CountA(aRange)
    ColumnCountA = ActiveSheet.Columns(useCol).Cells.SpecialCells(xlCellTypeConstants).count
End Function

Function LastColumn(Optional SHUse, Optional WBUse)
On Error GoTo Err1:

    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    
    LastColumn = Workbooks(WBUse).Sheets(SHUse).Cells.Find(What:="*", _
                    After:=Workbooks(WBUse).Worksheets(SHUse).range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Column
    Exit Function
Err1:
    LastColumn = 0
End Function

