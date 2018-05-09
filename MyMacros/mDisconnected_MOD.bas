Attribute VB_Name = "mDisconnected_MOD"
Sub GetDisconnects()
Dim fRange As range
Dim SHUse As String
Dim botRow As Long
    '
    ' Add RowNumbers to the Main Sheet
    '
    Worksheets("LastGasp").Activate
    
    col1 = FindColumnHeader("first_event_time")
    col2 = FindColumnHeader("METER_SERIAL_NUM")
    
    Call SortSheetUp(col1, col2)
    
    'Call AddRowNumbers(1)

    useCol = FindColumnHeader("src_ops_state")
    botRow = ColumnLastRow(useCol)
    
    Set sRange = range(Cells(2, useCol), Cells(botRow, useCol))
    Set fRange = FindInRange("Disconnected", sRange)
    Debug_Print fRange.count
    '
    ' Make the Disconnected sheet
    '
    If SheetExists("Disconnected") Then
        retCode = MsgBox("Reset Disconnected sheet?", vbYesNoCancel)
        If retCode = vbCancel Then Exit Sub
        
        If retCode = vbYes Then Call DeleteSheet("Disconnected")

    End If
    
    SHUse = MakeSheet(True, "Disconnected")
        
    Call CopyRangeRowsToSheet(fRange, SHUse)
    
    botRow = ColumnLastRow(1, SHUse)
    'Cells(botRow + 2, 1) = 2 ' row counter
    
    omsCol = FindColumnHeader("oms_incident")
    omsCount = LastRow("Outage")
    
    MsgBox (fRange.count & " DISCONNECTED meters." & vbNewLine & omsCount & " OMS Incident(s).")
    
    Sheets("LastGasp").Activate
    Call ClearClipboard
    Set fRange = Nothing
End Sub

Sub GetUnreachable()
Dim fRange As range
Dim SHUse As String
Dim botRow As Long
    '
    ' Add RowNumbers to the Main Sheet
    '
    Worksheets("LastGasp").Activate
    
    col1 = FindColumnHeader("first_event_time")
    col2 = FindColumnHeader("METER_SERIAL_NUM")
    
    Call SortSheetUp(col1, col2)
    
    'Call AddRowNumbers(1)

    useCol = FindColumnHeader("src_ops_state")
    botRow = ColumnLastRow(useCol)
    
    Set sRange = range(Cells(2, useCol), Cells(botRow, useCol))
    Set fRange = FindInRange("Unreachable", sRange)
    Debug_Print fRange.count
    '
    ' Make the Disconnected sheet
    '
    If SheetExists("Unreachable") Then
        retCode = MsgBox("Reset Unreachable sheet?", vbYesNoCancel)
        If retCode = vbCancel Then Exit Sub
        
        If retCode = vbYes Then Call DeleteSheet("Unreachable")

    End If
    
    SHUse = MakeSheet(True, "Unreachable")
    Call CopyRangeRowsToSheet(fRange, SHUse)
    
    Sheets("LastGasp").Activate
    Call ClearClipboard
    Set fRange = Nothing
End Sub
