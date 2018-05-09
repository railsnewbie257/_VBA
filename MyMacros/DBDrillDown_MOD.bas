Attribute VB_Name = "DBDrillDown_MOD"
Sub DrillDown()
    useValue = ActiveCell.Text
    WBOrig = ActiveWorkbook.Name
    SHOrig = ActiveSheet.Name
    
    topic = Cells(1, ActiveCell.Column).Value
    
    dateCol = FindColumnHeader("RunDate")
    useDate = Cells(2, dateCol)
    Select Case UCase(topic)
    
        Case "METER_SERIAL_NUM", "EQUIP_MFG_SERIAL_NUMBER"
            meterNum = ActiveCell.Value
            SHQuery = "MeterQuery"
            Call QueryNewCondition(SHQuery, MACROWORKBOOK, "m.EQUIP_MFG_SERIAL_NUMBER =", "'" & meterNum & "'")
            Call QueryNewCondition(SHQuery, "e.Event_Start_Dt =", format(useDate, "'" & "YYYY-MM-DD" & "'"))
            'Workbooks(MACROWORKBOOK).Sheets("Meter Query").Cells(33, 2) = "m.EQUIP_MFG_SERIAL_NUMBER =" & "'" & meterNum & "'"
            'Workbooks(MACROWORKBOOK).Sheets("Meter Query").Cells(31, 2) = "e.Event_Start_Dt = " & Format(useDate, "'" & "YYYY-MM-DD" & "'")
        
        Case "CIRCUIT_NUMBER"
            SHQuery = "MeterQuery"
            Call QueryNewCondition(SHQuery, MACROWORKBOOK, "m.CIRCUIT_NUMBER =", "'" & ActiveCell.Value & "'")
            'Workbooks(MACROWORKBOOK).Sheets("Meter Query").Cells(31, 2) = "m.CIRCUIT_NUMBER =" & "'" & ActiveCell.Value & "'"
    
        Case "FIRST_EVENT_TIME"
            SHQuery = "MeterQuery"
            Call QueryNewCondition(SHQuery, MACROWORKBOOK, "e.EVENT_START_TM =", "'" & ActiveCell.Value & "'")
            'Workbooks(MACROWORKBOOK).Sheets("Meter Query").Cells(31, 2) = "e.EVENT_START_TM =" & "'" & ActiveCell.Text & "'"
                
        Case Else
            MsgBox "Can not DrillDown on " & topic, vbOKOnly
            Exit Sub
    
    End Select

    Workbooks(WBOrig).Activate
    ' Workbooks(WBOrig).Worksheets.Add
    SHUse = ActiveSheet.Name
    
    
    s = QueryBuilder("MeterQuery")
    Debug.Print s
    Sheets.Add

    Call Query(s)
    
    If DBGlbRecordsToRead > 0 Then
        WBNew = ActiveWorkbook.Name
        SHNew = ActiveSheet.Name
        sortCol = FindColumnHeader("First_Event_Time", SHNew, WBNew)
        If sortCol = -1 Then sortCol = FindColumnHeader("Event_Start_Tm")
        Call SortSheetUp(sortCol)
        Sheets(SHNew).Name = useValue
    Else
        Call DeleteSheet(SHNew)
    End If
End Sub

Sub GetDate2()
    Workbooks(MACROWORKBOOK).Sheets("Meter Query").Cells(27, 1) = e.Event_Start_Dt = format(Cells(2, 1), "'" & "YYYY-MM-DD" & "'")
End Sub
