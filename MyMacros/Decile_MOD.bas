Attribute VB_Name = "Decile_MOD"
Sub Decilize()
Dim t As Range

    Call MakeScratchSheet
    Set colRange = Columns(ActiveCell.Column)
    colRange.Copy
    Worksheets("Scratch").Cells(1, 1).PasteSpecial Paste:=xlValues
    Worksheets("Scratch").Activate
    endRow = LastRow()
    idxCol = InsertColumnRight(ActiveCell.Column)
    Set t = Range(Cells(2, idxCol), Cells(endRow, idxCol))
    
    Cells(2, idxCol) = 1
    Cells(3, idxCol) = 2
    Cells(4, idxCol) = 3
    Set fromRange = Range(Cells(2, idxCol), Cells(4, idxCol))
    Set toRange = Range(Cells(2, idxCol), Cells(LastRow(), idxCol))
    fromRange.AutoFill Destination:=toRange

    factor = toRange.count / 10
    
    Set decRange = toRange.Offset(0, 1)
    
    decRange.Formula = "=MIN(INT((ROW()-1)/" & factor & ")+1,10)"
    

End Sub

Sub dec()
  factor = Selection.count / 10
  Selection.Formula = "=MIN(INT((ROW()-1)/" & factor & ")+1,10)"
End Sub

