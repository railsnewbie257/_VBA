Attribute VB_Name = "UnderVoltage_MOD"
Sub DeleteColumnUseHeader(headerName, Optional SHUse, Optional WBUse)
Dim aCol As Integer

    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    
    aCol = FindColumnHeader(headerName, SHUse, WBUse)
    Workbooks(WBUse).Worksheets(SHUse).Columns(aCol).Delete
End Sub

Sub MeterDeviceTypeLookup(Optional meterCol, Optional deviceCol)
Dim botRow As Long

    'Call ScreenOff
    'Call StartTimer
    
    If IsMissing(meterCol) Then meterCol = FindColumnHeader("src_name")
    If IsMissing(deviceCol) Then deviceCol = meterCol + 1
    botRow = ColumnLastRow(meterCol)
    Call ProgressMeterShow(0, botRow)
    For i = 2 To botRow
        Cells(i, deviceCol) = TeradataLookup(Cells(i, meterCol))
        If i Mod 10 = 0 Then Call ProgressMeterShow(i, botRow)
    Next i
    
    Call ProgressMeterClose
    'Call StopTimer
    'Call ElapsedTime
    'Call ScreenOn
End Sub
