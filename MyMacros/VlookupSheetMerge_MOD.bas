Attribute VB_Name = "VlookupSheetMerge_MOD"
Sub AnchorAppendRow()
Dim fromRange As Variant
Dim toRange As Variant

    On Error Resume Next
    Set fromRange = Application.InputBox("Click on FROM INDEX column", Type:=8)
    If IsEmpty(fromRange) Then Exit Sub
    
    fromRange.Copy
    SHFrom = fromRange.Parent.Name
    WBFrom = fromRange.Parent.Parent.Name
    fromCol = fromRange.Column
    botRowFrom = ColumnLastRow(fromCol, SHFrom, WBFrom)
    Set fromIndexRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(2, fromCol), _
                               Workbooks(WBFrom).Worksheets(SHFrom).Cells(botRowFrom, fromCol))
    fromColumnCount = LastColumn(SHFrom, WBFrom)
    
    On Error Resume Next
    Set toRange = Application.InputBox("Click on TO INDEX column", Title:="VlookupSheetMerge", Type:=8)
    If IsEmpty(toRange) Then Exit Sub
    
    SHTo = toRange.Parent.Name
    WBTo = toRange.Parent.Parent.Name
    toCol = toRange.Column
    lastColumnTo = NextColumn(SHTo, WBTo)
    
    leftColumn = NextColumn(SHTo, WBTo)
    Set copyRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(1, fromCol), _
                        Workbooks(WBFrom).Worksheets(SHFrom).Cells(1, fromColumnCount))
    copyRange.Copy destinastion:=Workbooks(WBTo).Worksheets(SHTo).Cells(1, leftColumn)
    For i = 2 To botRowFrom
        Set copyRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(i, fromCol), _
              Workbooks(WBFrom).Worksheets(SHFrom).Cells(i, fromColumnCount))
        copyRange.Copy destinastion:=Workbooks(WBTo).Worksheets(SHTo).Cells(i, leftColumn)
    Next i
    
    
End Sub
