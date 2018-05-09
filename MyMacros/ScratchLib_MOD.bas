Attribute VB_Name = "ScratchLib_MOD"
Function HeaderToSheet(SHFrom, SHTo)
    Worksheets(SHTo).Rows(1).Value = Worksheets(SHFrom).Rows(1).Value
End Function

Sub MakeScratchSheet(Optional copyHeader, Optional SHName)
    
    SHFrom = ActiveSheet.Name
    Call MakeSheet(False, "Scratch")
    Worksheets("Scratch").Activate
    Exit Sub
    
    SHCurrent = ActiveSheet.Name
    If IsMissing(SHName) Then SHName = "Scratch"
    nSheets = SheetExists(ActiveWorkbook.Name, SHName)
    If (nSheets > 0) Then
        retCode = MsgBox("Delete " & SHName & "?", vbYesNoCancel)
        If (retCode = vbCancel) Then Exit Sub
        If (retCode = vbYes) Then Call DeleteSheet(SHName)
        If (retCode = vbNo) Then SHName = SHName & "-" & (nSheets + 1)
            
    End If
    
    If Not SheetExists(ActiveWorkbook.Name, SHName) Then
        Set SHNew = ActiveWorkbook.Sheets.Add
        SHNew.Name = SHName
    End If
    Sheets(SHCurrent).Activate
    If Not IsMissing(copyHeader) Then
        If (copyHeader) Then Call HeaderToScratch(SHCurrent)
        
    Else
        retCode = MsgBox("Copy Headers To Scratch?", vbYesNoCancel)
        If (retCode = vbCancel) Then Exit Sub
        If (retCode = vbYes) Then Call HeaderToScratch(SHCurrent)
    End If
End Sub

Function ScratchNextColumn()
    ScratchNextColumn = LastColumn("Scratch") + 1
End Function

Sub CopyColumnToSheet(Optional SHFrom, Optional colNum)
Dim fromRange As Range
Dim i As Integer, j As Integer
Dim botRow As Long
    
    On Error Resume Next
    Set fromRange = Nothing
    Set fromRange = Application.InputBox("Select Columns To Copy", Default:=Selection.Address, Title:="CopyColumnsToScratch", Type:=8)
    If fromRange Is Nothing Then Exit Sub
    
    Set toRange = Nothing
    Set toRange = Application.InputBox("Select Destination Sheet" & vbNewLine & "(same sheet to use SCRATCH", Default:=Selection.Address, Title:="CopyColumnsToScratch", Type:=8)
    If toRange Is Nothing Then Exit Sub
    
    WBFrom = fromRange.Parent.Parent.Name
    SHFrom = fromRange.Parent.Name
    WBTo = toRange.Parent.Parent.Name
    SHTo = toRange.Parent.Name
    
    If SHTo = SHFrom Then SHTo = MakeSheet(False, "Scratch")
    
    
    For i = 1 To fromRange.Areas.count
        For j = 1 To fromRange.Areas(i).Columns.count
            botRow = LastRow(SHFrom, WBFrom)
            colNum = NextColumn(SHTo, WBTo)
            fromCol = fromRange.Areas(i).Columns(j).Column
            Set fromRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(1, fromCol), _
                                  Workbooks(WBFrom).Worksheets(SHFrom).Cells(botRow, fromCol))
            toCol = NextColumn(SHTo, WBTo)
            'Set toRange = Workbooks(WBTo).Worksheets(SHTo).Cells(1, toCol)
            'toRange.PasteSpecial Paste:=xlPasteAll
            fromRange.Copy Destination:=Workbooks(WBTo).Worksheets(SHTo).Cells(1, toCol)
        Next j
    Next i
    
    Worksheets(SHTo).Activate
    Call ColumnWidthAutoMax(SHTo, scratchCol, 30)
End Sub
