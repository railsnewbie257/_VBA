Attribute VB_Name = "ConvertUTC_Mod"
Sub ExtractOUI()
    useCol = ActiveCell.Column
    Columns(useCol).EntireColumn.Offset(0, 1).Insert
End Sub

Sub ConvertUTCToSerialNumber()
    useCol = ActiveCell.Column
    newCol = InsertColumnRight(useCol)
    
    Cells(1, newCol).Formula = "DateSerialNumber"
    Set colRange = Range(Cells(2, newCol), Cells(LastRow, newCol))
    useAddr = Cells(2, useCol).Address(False, False, xlA1)
    colRange.FormulaLocal = "=Left(" & useAddr & ", 10) +  Mid(" & useAddr & ",12,8)"
End Sub


