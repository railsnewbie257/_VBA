Attribute VB_Name = "mProximity_MOD"
Sub Proximity()

    circuitCol = FindColumnHeader("circuit_number")
    transformCol = FindColumnHeader("transformer_number")
    cityCol = FindColumnHeader("pos_city_name")
    zipCol = FindColumnHeader("proximity_zip_code")
    
    useRow = ActiveCell.Row
    
    currentColor = ChooseColor(currentColor)
    rightColumn = LastColumn()
    Set aRange = Range(Cells(useRow, 1), Cells(useRow, rightColumn))
    Call ColorRange(aRange, currentColor)
    
    Call ProximityHiLite(circuitCol, useRow)
    Call ProximityHiLite(transformCol, useRow)
    Call ProximityHiLite(cityCol, useRow)
    Call ProximityHiLite(zipCol, useRow)
    
End Sub

Sub lkj()
    psumCol = FindColumnHeader("p_sum")
    botRow = LastRow
    Set fRange = Range(Cells(2, psumCol), Cells(2, psumCol))
    Set sumRange = Range(fRange.Offset(0, 1), fRange.Offset(0, 5))
    Set fRange = Range(Cells(2, psumCol), Cells(botRow, psumCol))
    fRange.Formula = "=SUM(" & sumRange.Address(False, False) & ")"
    Call RangeToValues(fRange)
End Sub
Sub ProximityCount2()

    meterCol = FindColumnHeader("meter_serial_num")
    circuitCol = FindColumnHeader("circuit_number")
    transformCol = FindColumnHeader("transformer_number")
    cityCol = FindColumnHeader("pos_city_name")
    zipCol = FindColumnHeader("proximity_zip_code")
    timeCol = FindColumnHeader("first_event_time")
    useRow = ActiveCell.Row
    
    rightColumn = LastColumn()
    Set aRange = Range(Cells(useRow, 1), Cells(useRow, rightColumn))
    Call ColorRange(aRange, currentColor)
    
    'currentColor = ChooseColor(currentColor)

    
    ptimeCol = FindColumnHeader("p_time")
    pcityCol = FindColumnHeader("p_city")
    pzipCol = FindColumnHeader("p_zip")
    ptimeCol = FindColumnHeader("p_time")
    pcircuitCol = FindColumnHeader("p_circuit")
    ptransformerCol = FindColumnHeader("p_transformer")
    
    useRow = ActiveCell.Row
    
    Cells(useRow, pcircuitCol) = ProximityHiLiteCount(circuitCol, useRow)
    Cells(useRow, ptransformerCol) = ProximityHiLiteCount(transformCol, useRow)
    Cells(useRow, pcityCol) = ProximityHiLiteCount(cityCol, useRow)
    Cells(useRow, pzipCol) = ProximityHiLiteCount(zipCol, useRow)
    Cells(useRow, ptimeCol) = ProximityHiLiteCount(timeCol, useRow)

    
End Sub

Sub ProximityHiLite(useCol, useRow)
    topRow = Application.WorksheetFunction.Max(useRow - ROWSPAN, 2)
    botRow = useRow + ROWSPAN
    useValue = Cells(useRow, useCol)
    
    For i = topRow To botRow
        If Cells(i, useCol) = useValue Then Call ColorRange(Cells(i, useCol), currentColor)
    Next i

End Sub

Function ProximityCount(useCol, useRow)
    topRow = Application.WorksheetFunction.Max(useRow - ROWSPAN, 2)
    botRow = useRow + ROWSPAN
    useValue = Cells(useRow, useCol)
    
    count = 0
    For i = topRow To botRow
        If Cells(i, useCol) = useValue Then
            ' Call ColorRange(Cells(i, useCol), currentColor)
            count = count + 1
        End If
    Next i
    ProximityCount = count
End Function
'
' Sets up the Proximity sheet and columns
'
Sub ProximityColumns()
Dim i As Long, botRow As Long
Dim t As String

    SHFrom = "LastGasp"  'BaseSheet()
    
    zipCol = FindColumnHeader("proximity_zip_code")
    If zipCol < 0 Then Call ProximityZipCodeColumn
    
    SHTo = "Proximity"
    Call CheckSheetExists(SHTo)

    Call AddColumnToSheet(1, SHFrom, SHTo) ' row numbers

    SHUse = "Proximity Columns"

    botRow = ColumnLastRow(1, SHUse, WBMacros)
    For i = 2 To botRow
    t = Workbooks(WBMacros).Sheets(SHUse).Cells(i, 1).Value
    Debug.Print t
        Call AddColumnToSheet(Workbooks(WBMacros).Sheets(SHUse).Cells(i, 1).Value, SHFrom, SHTo, Workbooks(WBMacros).Sheets(SHUse).Cells(i, 2))
    Next i
    
    Call FreezeHeader(SHTo)
    Worksheets("Proximity").Activate
    Exit Sub
    
    Call AddColumnToSheet("meter_serial_num", SHFrom, SHTo, 12)
    Call AddColumnToSheet("src_ops_state", SHFrom, SHTo)
    Call AddColumnToSheet("pos_city_name", SHFrom, SHTo, 15)
    Call AddColumnToSheet("proximity_zip_code", SHFrom, SHTo, 8)
    Call AddColumnToSheet("transformer_number", SHFrom, SHTo, 12)
    Call AddColumnToSheet("transformer_point_install_type", SHFrom, SHTo, 15)
    Call AddColumnToSheet("transformer_point_num_of_units", SHFrom, SHTo)
    Call AddColumnToSheet("circuit_number", SHFrom, SHTo)
    

    
    Worksheets(SHTo).Activate

End Sub
