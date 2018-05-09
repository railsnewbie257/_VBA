Attribute VB_Name = "A_Dahl_MOD"
Sub setupDahl()

    SHFrom = ActiveSheet.Name ' setup source
    Call InitSheet("Dahl") 'setup destination
    
    Call ExtractColumnToSheet("Dahl", "event_time", SHFrom)
    Call ExtractColumnToSheet("Dahl", "src_util_id", SHFrom)
    Call ExtractColumnToSheet("Dahl", "src_device_type", SHFrom)
    Call ExtractColumnToSheet("Dahl", "src_admin_state", SHFrom)
    Call ExtractColumnToSheet("Dahl", "src_ops_state", SHFrom)
    Call ExtractColumnToSheet("Dahl", "src_addr_line1", SHFrom)
    Call ExtractColumnToSheet("Dahl", "src_addr_line1", SHFrom)
    Call ExtractColumnToSheet("Dahl", "src_city", SHFrom)
    Call ExtractColumnToSheet("Dahl", "src_postal_code", SHFrom)
    Call ExtractColumnToSheet("Dahl", "src_dist_net_transformer_util_id", SHFrom)
    
End Sub

Sub vilyColumns()

    SHFrom = ActiveSheet.Name ' setup source
    Call InitSheet("Dahl") 'setup destination
    
    Call ExtractColumnToSheet("Dahl", "event_log_id", SHFrom)
    Call ExtractColumnToSheet("Dahl", "event_time", SHFrom)
    Call ExtractColumnToSheet("Dahl", "src_util_id", SHFrom)
    Call ExtractColumnToSheet("Dahl", "src_device_type", SHFrom)
    Call ExtractColumnToSheet("Dahl", "src_admin_state", SHFrom)
    Call ExtractColumnToSheet("Dahl", "src_ops_state", SHFrom)
    Call ExtractColumnToSheet("Dahl", "src_addr_line1", SHFrom)
    Call ExtractColumnToSheet("Dahl", "src_city", SHFrom)
    Call ExtractColumnToSheet("Dahl", "src_postal_code", SHFrom)
    Call ExtractColumnToSheet("Dahl", "src_dist_net_transformer_util_id", SHFrom)
    
End Sub

Sub disconnects()

    stateCol = HeaderToColumnNum("src_admin_state")
    
    Set searchRange = Range(Cells(2, stateCol), Cells(LastRow, stateCol))
    Set fRange = FindInRangeExact("Disconnected", searchRange)
    Debug.Print fRange.Count
    
    
    For Each r In fRange
        Call ColorRange(Range(Cells(r.Row, 1), Cells(r.Row, 10)), 5)
    Next r
    
    Call ColorRange(Range(Cells(fRange(1).Row, 1), Cells(fRange(1).Row, 5)), 5)
    Call ColorRange(fRange(2).EntireRow, 4)
    
    fRange.EntireRow(1).Copy
    Range(fRange.EntireRow(2)).Copy
    Call ColorRange(fRange(1).EntireRow)
    fRange(1).Copy
    Set aRange = fRange(2).EntireRow
    aRange.Copy

    Call ColorRange(fRange(2).EntireRow)
End Sub


