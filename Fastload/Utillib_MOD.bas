Attribute VB_Name = "Utillib_MOD"

Function MacroTimestamp() As String
    MacroTimestamp = ThisWorkbook.Worksheets("Pallette").Cells(8, 1)
End Function

Function ClearClipboard()
    Application.CutCopyMode = False
End Function


Function NextRow_duplicate(Optional SHUse, Optional WBUse)
On Error GoTo Err1:
    NextRow = LastRow(SHUse, WBUse) + 1
    Exit Function
Err1:
    LastRow = 0
    
End Function

