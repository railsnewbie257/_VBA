Attribute VB_Name = "Show_MOD"
Option Explicit

Sub ShowDetails()
Dim useRow As Long
Dim toCol As Long
Dim SHTo As String, WBTo As String
Dim SHFrom As String, WBFrom As String
Dim useSelection As Range
Dim useHeader As String
Dim useRange As Range
Dim fRange As Range
Dim copyRange As Range
Dim headerRange As Range
Dim t As Long


    'Call ScreenOff
    
    Set useSelection = Selection

    SHFrom = useSelection.Parent.Name
    WBFrom = useSelection.Parent.Parent.Name

    Call DeleteSheet("ShowDetails")

    SHTo = MakeSheet(False, "ShowDetails")
    WBTo = ActiveWorkbook.Name
    
    useHeader = Cells(1, ActiveCell.Column)
    
    ' Copy the header names
    Worksheets(SHTo).Activate
    useRow = ActiveCell.Row
    Set headerRange = Range(Worksheets(SHFrom).Cells(1, 1), Worksheets(SHFrom).Cells(1, LastColumn(SHFrom)))
    headerRange.Copy
    Worksheets(SHTo).Cells(2, 1).PasteSpecial Paste:=xlPasteAll, Transpose:=True
    'Worksheets(SHTo).Activate
    Range(Worksheets(SHTo).Cells(2, 1), Worksheets(SHTo).Cells(ColumnLastRow(1), 1)).Font.Bold = True
    With Worksheets(SHTo).Cells(1, 1)
        .Value = "Column Names"
        .Font.Bold = True
    End With
    Call ColumnWidthAutoMax(SHTo, 1, 30)
    Call ColorRange(Worksheets(SHTo).Cells(1, 1), LIGHTBLUE)

    ' Copy the Values -------------------------------------------------------------------------------------------
    'Call ScreenOff
    Worksheets(SHFrom).Activate ' need to go back to original sheet to get the selection
    For Each useRange In Selection
        useRow = useRange.Row
        t = RowLastColumn(useRow, SHFrom, WBFrom)
        Set copyRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(useRow, 1), Workbooks(WBFrom).Worksheets(SHFrom).Cells(useRow, RowLastColumn(useRow, SHFrom, WBFrom)))

        'copyRange.Copy Destination:=Workbooks(WBTo).Worksheets(SHTo).Cells(20, 3)
        toCol = RowNextColumn(1, SHTo)
        copyRange.Copy
        Workbooks(WBTo).Worksheets(SHTo).Cells(2, toCol).PasteSpecial Paste:=xlPasteAll, Transpose:=True
        Workbooks(WBTo).Worksheets(SHTo).Cells(2, toCol).Copy
        Workbooks(WBTo).Worksheets(SHTo).Columns(toCol).HorizontalAlignment = xlLeft
    
        With Worksheets(SHTo).Cells(1, toCol)
            .Value = "Column Values"
            .Font.Bold = True
        End With
        Call ColumnWidthAutoMax(SHTo, toCol, 30)
        Call ColorRange(Worksheets(SHTo).Cells(1, toCol), LIGHTBLUE)
    Next useRange
    Call ScreenOn
    Worksheets(SHTo).Activate

    ' Highlight where you were
    Set fRange = FindInRange(useHeader, Cells)
    fRange.Offset(0, 1).Select
    fRange.Offset(0, 1).Copy
    
    Call ScreenOn
End Sub

Sub QueryToLines(Optional txtRange, Optional toRange)
Dim t As String
'Dim txtRange As Range
'Dim toRange As Range

    On Error Resume Next
    If IsMissing(txtRange) Then
        Set txtRange = Nothing
        Set txtRange = Application.InputBox("Query String Location", Title:="QueryToLine", Type:=8)
        If txtRange Is Nothing Then Exit Sub
    End If
    
    t = txtRange.Value
    t = Replace(t, ",", "," & Chr(10))
    Call CopyToClipboard(t)
    
    If IsMissing(toRange) Then
        Set toRange = Nothing
        Set toRange = Application.InputBox("Put Query Location:", Title:="QueryToLine", Type:=8)
        If toRange Is Nothing Then Exit Sub
    End If
    
    toRange.PasteSpecial
End Sub

Sub ShowTable_Callback()

    Columns(2).columnWidth = 30
    
    Call QueryToLines(Cells(2, 2), Cells(4, 1))
    
End Sub

Sub ShowVars()

    MsgBox GLBColumnNamesWB

End Sub
