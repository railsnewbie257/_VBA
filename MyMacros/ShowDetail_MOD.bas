Attribute VB_Name = "ShowDetail_MOD"
Sub ShowDetails()

    'Call ScreenOff
    
    Set useSelection = Selection

    SHFrom = useSelection.Parent.Name
    SHTo = MakeSheet(False, "Show")
    
    useHeader = Cells(1, ActiveCell.Column)
    
    ' Copy the header names
    Worksheets(SHTo).Activate
    useRow = ActiveCell.Row
    Set headerRange = Range(Worksheets(SHFrom).Cells(1, 1), Worksheets(SHFrom).Cells(1, LastColumn(SHFrom)))
    headerRange.Copy
    Worksheets(SHTo).Cells(3, 1).PasteSpecial Paste:=xlPasteAll, Transpose:=True
    'Worksheets(SHTo).Activate
    Range(Worksheets(SHTo).Cells(3, 1), Worksheets(SHTo).Cells(LastRow(SHTo), 1)).Font.Bold = True
    With Worksheets(SHTo).Cells(2, 1)
        .Value = "Column Name"
        .Font.Bold = True
    End With
    Call ColorRange(Worksheets(SHTo).Cells(2, 1), YELLOW)
    Call ColumnWidthAutoMax(SHTo, 1, 30)
    
    newCol = ColumnInsertRight(0)
    nRows = ColumnLastRow(2) - 1
    For i = 1 To nRows
        Worksheets(SHTo).Cells(i + 2, newCol) = "ColNum=" & i
    Next i

    
    ' Copy the values
    Worksheets(SHFrom).Activate
    For Each useRange In useSelection
        useRow = useRange.Row
        Set copyRange = Worksheets(SHFrom).Range(Cells(useRow, 1), Cells(useRow, LastColumn(SHFrom)))
        copyRange.Copy
        toCol = NextColumn(SHTo)
        Worksheets(SHTo).Cells(1, toCol) = "RowNum=" & useRow
        Worksheets(SHTo).Cells(3, toCol).PasteSpecial Paste:=xlPasteAll, Transpose:=True
        Call ColumnWidthAutoMax(SHTo, toCol, 30)
        Worksheets(SHTo).Columns(toCol).HorizontalAlignment = xlLeft
    Next useRange
    With Worksheets(SHTo).Cells(2, 3)
        .Value = "Column Value"
        .Font.Bold = True
    End With
    Worksheets(SHTo).Activate
    
    Cells(1, 1) = "From=" & SHFrom

    ' Highlight where you were
    Set fRange = FindInRange(useHeader, Cells)
    fRange.Offset(0, 1).Select
    fRange.Offset(0, 1).Copy
    
    Call ScreenOn
End Sub
