Attribute VB_Name = "ColumnUtils_MOD"
' Routines that deal with columns
'
'
' Append daat from one column to another column
'
Sub ColumnAppend(Optional fromRange, Optional toRange)
Dim aRange As range
Dim botRowFrom As Long, nextRowTo As Long
Dim fromCol As Integer, toCol As Integer

    On Error Resume Next
    If IsMissing(fromRange) Then
        Set fromRange = Nothing
        Set fromRange = Application.InputBox("Select Column To Append", Default:=Selection.Address, Type:=8)
        If fromRange Is Nothing Then Exit Sub
    End If
    
    If IsMissing(toRange) Then
        Set toRange = Nothing
        Set toRange = Application.InputBox("Select Column To Append To", Default:=Selection.Address, Type:=8)
        If toRange Is Nothing Then Exit Sub
    End If
    
    
    WBFrom = fromRange.Parent.Parent.Name
    SHFrom = fromRange.Parent.Name
    fromCol = fromRange.Column
    
    WBTo = toRange.Parent.Parent.Name
    SHTo = toRange.Parent.Name
    toCol = toRange.Column
    
    botRowFrom = ColumnLastRow(fromRange.Column, SHFrom, WBFrom)
    nextRowTo = ColumnNextRow(toRange.Column, SHTo, WBTo)
    
    Set aRange = range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(2, fromCol), _
                  Workbooks(WBFrom).Worksheets(SHFrom).Cells(botRowFrom, fromCol))
    aRange.Copy Destination:=Workbooks(WBTo).Worksheets(SHTo).Cells(nextRowTo, toCol)
    
End Sub

Sub ColumnsDefaultWidth()
    Call ColumnWidthAutoMax(maxWidth:=8.43)
    MsgBox "Finished."
End Sub

Function ColumnWidthAutoMax(Optional SHUse, Optional useCol, Optional maxWidth)

    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(maxWidth) Then maxWidth = 30
    
    If IsMissing(useCol) Then ' assume it's for the entire sheet
        Worksheets(SHUse).Columns.AutoFit
        rightCol = LastColumn()
        For i = 1 To rightCol
            qw = Worksheets(SHUse).Columns(i).columnWidth
            If qw > maxWidth Then Worksheets(SHUse).Columns(i).columnWidth = maxWidth
        Next i
    Else
        Worksheets(SHUse).Columns(useCol).AutoFit
        qw = Worksheets(SHUse).Columns(useCol).columnWidth
        If qw > maxWidth Then Worksheets(SHUse).Columns(useCol).columnWidth = maxWidth
    End If
End Function

Sub ColumnAlign(headerName, Optional SHUse, Optional useAlign)
    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(useAlign) Then useAlign = xlRight
    
    useCol = FindColumnHeader(headerName)
    Worksheets(SHUse).Columns(useCol).HorizontalAlignment = useAlign
End Sub

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

Function ColumnNextRow(Optional useCol, Optional SHUse, Optional WBUse)
    ColumnNextRow = ColumnLastRow(useCol, SHUse, WBUse) + 1
End Function

Function ColumnDataRange(useCol, Optional SHUse, Optional WBUse) As range
Dim botRow As Long

    Call DefaultWorkbookAndSheet(SHUse, WBUse)
    
    botRow = ColumnLastRow(useCol, SHUse, WBUse)
    Set ColumnDataRange = range(Cells(2, useCol), Cells(botRow, useCol))
End Function

Function ColumnCountA(useCol)
    Set aRange = range(Cells(2, useCol), Cells(ColumnLastRow(useCol), useCol))
    ColumnCountA = WorksheetFunction.CountA(aRange)
    ColumnCountA = ActiveSheet.Columns(useCol).Cells.SpecialCells(xlCellTypeConstants).count
End Function
'
' similar to filterMultipleWorkOrders
'
Sub ColumnCountValues(Optional useCol)
Dim colRange As range
Dim botRow As Long, i As Long, k As Long

    If IsMissing(useCol) Then useCol = ActiveCell.Column
    Set colRange = range(Cells(1, useCol), Cells(1, useCol))
    
    Call SortSheetUp(colRange.Column)
    
    countCol = ColumnInsertLeft(colRange.Column)
    Cells(1, countCol) = "Count of " & Cells(1, colRange.Column)
    
    botRow = ColumnLastRow(colRange.Column) + 1
    
    i = 2
    k = 1
    While i < botRow
        While (Cells(i, colRange.Column) = Cells(i + 1, colRange.Column))
            k = k + 1
            Rows(i + 1).Delete
            botRow = botRow - 1
        Wend
        Cells(i, countCol) = k
        
        i = i + 1
        k = 1
    Wend

End Sub

' assumes there is a header row at top
Function ExtractColumnToSheet(SHTo, colName, Optional SHFrom)

    If IsMissing(SHFrom) Then SHFrom = ActiveSheet.Name
    
    ' find column
    ''Set headerRange = Worksheets(SHFrom).Rows(1) ' for finding columns by name
    
    ''Set aRange = FindInRangeExact(colName, headerRange)
    fromCol = HeaderToColumnNum(colName, SHFrom)
    
    toCol = NextColumn(SHTo)
    Set fromRange = Worksheets(SHFrom).Columns(fromCol)
    fromRange.Copy
    
    Set toRange = Worksheets(SHTo).Cells(1, toCol)
    toRange.PasteSpecial Paste:=xlAll
    
End Function

Function HeaderToColumnNum(useHeader, Optional SHUse)
On Error GoTo None:

    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    Set headerRange = Worksheets(SHUse).Rows(1) ' for finding columns by name
    Set aRange = FindInRangeExact(useHeader, headerRange)
    HeaderToColumnNum = aRange.Column
    Exit Function
    
None:
    HeaderToColumnNum = -1
    
End Function

Function ColumnNumToLetter(iCol) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   
   t = Cells(1, iCol).Address(False, False, xlA1)
   ColumnNumToLetter = left(t, Len(t) - 1)

End Function

Function IsColumnEmpty(colNum, Optional SHUse, Optional WBUse) As Boolean

    On Error GoTo gotError
    
10  If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
20  If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
30  IsColumnEmpty = False
40  If ColumnLastRow(colNum, SHUse, WBUse) <= 1 Then IsColumnEmpty = True

    Exit Function
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="IsColumnEmpty"
    Stop
    Resume Next
End Function

Function NextColumn(Optional SHUse, Optional WBUse)
On Error GoTo Err1:
    NextColumn = LastColumn(SHUse, WBUse) + 1
    Exit Function
Err1:
    LastColumn = 0
    
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

Function ThisColumnRange()
Dim colRange As range

    useCol = ActiveCell.Column
    Set colRange = range(Cells(1, useCol), Cells(ColumnLastRow, useCol))
    Set ThisColumnRange = colRange
End Function

Sub ColumnConCat(Optional colRange, Optional toRange)
Dim aRange As range
Dim i As Integer, j As Integer
Dim fff As String
Dim botRow As Long

    If IsMissing(colRange) Then
        On Error Resume Next
        Set colRange = Nothing
        Set colRange = Application.InputBox("Select Columns To Concatenate", Default:=Selection.Address, Title:="ColConCat", Type:=8)
        If colRange Is Nothing Then Exit Sub
    End If
    
    If IsMissing(toRange) Then
        Set toRange = Nothing
        Set toRange = Application.InputBox("Select Destination", Title:="ColConCat", Type:=8)
        If toRange Is Nothing Then Exit Sub
    End If
        
    botRow = LastRow()
    
    fff = "="
    For i = 1 To colRange.Areas.count
        For j = 1 To colRange.Areas(i).Columns.count
            fff = fff & left(colRange.Areas(i).Columns(j).End(xlUp).Address(False, False), 1) & "2" & "&"
        Next j
    Next i
    
    fff = left(fff, Len(fff) - 1)
    
    useCol = toRange.Column
    newCol = ColumnInsertLeft(toRange.Column)
    
    Set aRange = range(Cells(2, newCol), Cells(botRow, newCol))
    
    aRange.Formula = fff
    
    Call RangeToValues(aRange)
    
            
End Sub

