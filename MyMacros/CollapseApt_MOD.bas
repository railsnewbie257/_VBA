Attribute VB_Name = "CollapseApt_MOD"
'
' Collapses and Uncollapses multiple rows with the same meter
'

Sub CollapseApt()
    Set origActiveCell = ActiveCell
    
    bpCol = FindColumnHeader("bp_num")
    addressCol = FindColumnHeader("pos_address_line_1")
    Call SortSheetUp(bpCol, addressCol)
    
    origActiveCell.Select
    '
    ' figure out if you are in a collapse zone
    '
    useRow = ActiveCell.Row
    If Rows(useRow).Interior.Pattern = 4000 Then
        While Rows(useRow).Interior.Pattern = 4000 And useRow > 1
            useRow = useRow - 1
        Wend
        startRow = useRow
        useRow = useRow + 1
        While Rows(useRow).Interior.Pattern = 4000 And useRow > 1
            useRow = useRow + 1
        Wend
        botRow = useRow
    Else
        startRow = 2
        botRow = ColumnLastRow(bpCol)
        'Call SortSheetUp(bpCol, addressCol)
    End If
    
    oldBP = ""
    oldAddress = ""
    For i = startRow To botRow - 1
        If Cells(i, bpCol) = oldBP And Cells(i, addressCol) = oldAddress Then
            Call ColorRange(Rows(i), APTCOLLAPSE)
            Call ColorRange(Rows(i - 1), APTCOLLAPSE)
            Rows(i).Hidden = True
        Else
            oldBP = Cells(i, bpCol)
            oldAddress = Cells(i, addressCol)
        End If
    Next i
    
End Sub

Sub UncollapseApt()
Dim t As Integer

    ' bpCol = FindColumnHeader("BP_NUM")
    useRow = ActiveCell.Row

    While Rows(useRow).Interior.Pattern = 4000
        t = Rows(useRow).Interior.Pattern
        Rows(useRow).Hidden = False
        useRow = useRow + 1
    Wend
    
End Sub
