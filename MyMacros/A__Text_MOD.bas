Attribute VB_Name = "A__Text_MOD"
Sub test_Setup_Box(Optional nRows, Optional nCols)
    If IsMissing(nRows) Then nRows = 10
    If IsMissing(nCols) Then nCols = 10

    For i = 1 To nRows
        For j = 1 To nCols
            Cells(i, j) = "Cells(" & i & "," & j & ")"
        Next j
    Next i
    
End Sub
