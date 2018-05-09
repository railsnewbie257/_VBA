Attribute VB_Name = "DailyProcessing_MOD"
Sub LastGasp()
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset
Dim t As Date

    On Error GoTo gotError
    
10    Call UsageTracker("Last Gasp", "Start")

20    LOAD ChooseDateForm
30    ChooseDateForm.MonthView1.Value = format(Date - 1, "m/d/yyyy")
40    ChooseDateForm.Show
50    If formCancel Then
60        Unload ChooseDateForm
70        Exit Sub
80    End If
    t = DateValue(ChooseDateForm.MonthView1.Value)
90    GlbUseDate = format(t, "YYYY-MM-DD")
100    Unload ChooseDateForm
    '
    ' Check for Silver Spring file
    '
110    useName = "SSN-" & GlbUseDate & ".xlsx"
120    resultCode = Dir(SSNPATH & useName)
130    If Not (useName = resultCode) Then
140        MsgBox "SSN Meter file not found." & vbNewLine & vbNewLine & "Please process SNS Meters for " & GlbUseDate
        Exit Sub
        'Set fromRange = Application.InputBox("Click on SSN Worksheet.", Type:=8)
        'WBFrom = fromRange.Parent.Parent.Name
        'SHFrom = fromRange.Parent.Name
    End If
    '
    ' Check Last Gasp database date
    '
150    userQuery = "select RunDate from dl_oge_analytics." & TD_LASTGASP & " where RunDate = '" & GlbUseDate & "';"
    Debug_Print userQuery
160    Set DBCn = DBCheckConnection(DBCn)
170    Set DBRs = DBCheckRecordset(DBRs)
    
180    Set DBRs.ActiveConnection = DBCn
    
190    DBRs.Open userQuery
    
200    todayDate = format(Now(), "m/d/yyyy")
    t = DBRs.Fields.count
210    If DBRs.BOF And DBRs.EOF Then  ' if empty then load it
        'dbDate = Format(DBRs.Fields(0).Value, "YYYY-MM-DD") ' this is the date
        'If Not dbDate = GlbUseDate Then
220            MsgBox "Please Update Last Gasp database for " & GlbUseDate & "."
230            Call DBCloseRecordset(DBRs)
            ' Call DBCloseConnection(DBCn)
240            Exit Sub
        'End If
    End If
    
    '====================================================================================================================
DoIt:
    'Workbooks.Add
    '
    ' Load the Last Gasp report
    '
250    Call LastGaspDaily(GlbUseDate)
260    If formCancel Then Exit Sub
    
    '
    ' first sort the Last Gasp report
    '
270    sortCol = FindColumnHeader("First_Event_Time_12007")
280    Call SortSheetUp(sortCol)
    
290    Call ScreenOff
    
300    Call ProximityZipCodeColumn
310    Call filterMultipleWorkOrders
320    Call SSNMeterStatus
330    Call EventTimeHilite
340    Call MakeSingletons
350    Call SingletonsHilite
355    Call RemoveOMS
360    Call GetSingletons
370    Call GetDisconnects
    
380    Call ScreenOn
    'Call ProximityColumns
    
    'Worksheets("A-Single").Activate
    'Call SingletonsHilite
    'Worksheets("D-Single").Activate
    'Call SingletonsHilite
    'Worksheets("Disconnected").Activate
    'Call EventTimeHilite
    
390    Call UsageTracker("Last Gasp", "Finished")
    
400    MsgBox "Last Gasp Processing Finished."
    
    Exit Sub
    
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl
    Stop
    Resume Next
End Sub

Sub ZeroKWH()

    Call UsageTracker("Zero KWH", "Start")
    
    ' GLBUserQuery = "SELECT * FROM dl_oge_analytics.Zero_KWH ORDER BY 2;"
    
    GLBUserQuery = QueryBuilder("ZeroKWHSelect", MACROWORKBOOK)
    GLBQueryName = "ZeroKWH"
    GlbStatusBarTxt = "Running Zero KWH"
    
    Call Query(GLBUserQuery)
    
    bpCol = FindColumnHeader("BP_NUM")
    addressCol = FindColumnHeader("POS_ADDRESS_LINE_1")
    
    Call SortSheetUp(bpCol, addressCol)
    
    Call CollapseApt

    Call UsageTracker("Zero KWH", "Finished")
    
    MsgBox "Zero KWH Processing Finished."
End Sub

Sub ReceivedEnergy()
    Call UsageTracker("ReceivedEnergy", "Start")

    GLBQueryName = "ReceivedEnergy"
    GlbStatusBarTxt = "Running ReceivedEnergy"
    
    GLBUserQuery = QueryBuilder("ReceivedEnergySelect", MACROWORKBOOK)
    Call Query(GLBUserQuery)
    
    Call UsageTracker("ReceivedEnergy", "Finished")
    
    MsgBox "Received Energy Processing Finished."
End Sub

Sub PhaseAngleAlarm()

    Call UsageTracker("PhaseAngleAlarm", "Start")

    'GLBUserQuery = "select * from dl_oge_analytics.PhaseAngleAlarm"
    
    GLBUserQuery = QueryBuilder("PhaseAngleSelect", MACROWORKBOOK)
    GLBQueryName = "PhaseAngleAlarm"
    GlbStatusBarTxt = "Running PhaseAngleAlarm"
    
    Call Query(GLBUserQuery)

    Call UsageTracker("PhaseAngleAlarm", "Finished")
    
    MsgBox "Phase Angle Alert Processing Finished."
End Sub

Sub UnderVoltage()

    Call UsageTracker("Under Voltage", "Start")

    'GLBUserQuery = "select * from dl_oge_analytics.KV2C_Under_Voltage"
    'GLBQueryName = "KV2CUnderVoltage"
    'GlbStatusBarTxt = "Running UnderVoltage"
    
    'Call Query(GLBUserQuery)

    '
    ' Is currently run off of an excel download
    '
    ' Get the file
    '
    ' Check the file is correct
    WBUse = ActiveWorkbook.Name
    SHUse = ActiveSheet.Name
    useCol = FindColumnHeader("event_id", SHUse, WBUse)
    On Error Resume Next
    If Workbooks(WBUse).Worksheets(SHUse).Cells(2, useCol) <> "15060" Or useCol < 0 Then
        MsgBox "Please load an UnderVoltage file", Title:="UnderVoltage"
        Exit Sub
    End If
    
    WBUse = ActiveWorkbook.Name
    SHUse = ActiveSheet.Name
    '
    'Copy the key
    '
    useCol = FindColumnHeader("event_log_id", SHUse, WBUse)
    Workbooks.Add
    WBNew = ActiveWorkbook.Name
    SHNew = ActiveSheet.Name
    
    Workbooks(WBUse).Worksheets(SHUse).Columns(useCol).Copy Destination:=Workbooks(WBNew).Worksheets(SHNew).Cells(1, RowNextColumn(1))
    
    useCol = FindColumnHeader("event_time", SHUse, WBUse)
    i = RowNextColumn(1, SHNew, WBNew)
    Workbooks(WBUse).Worksheets(SHUse).Columns(useCol).Copy Destination:=Workbooks(WBNew).Worksheets(SHNew).Cells(1, i)
    
    useCol = FindColumnHeader("event_id", SHUse, WBUse)
    i = RowNextColumn(1, SHNew, WBNew)
    Workbooks(WBUse).Worksheets(SHUse).Columns(useCol).Copy Destination:=Workbooks(WBNew).Worksheets(SHNew).Cells(1, i)
    
    useCol = FindColumnHeader("src_name", SHUse, WBUse)
    i = RowNextColumn(1, SHNew, WBNew)
    Workbooks(WBUse).Worksheets(SHUse).Columns(useCol).Copy Destination:=Workbooks(WBNew).Worksheets(SHNew).Cells(1, i)

    useCol = FindColumnHeader("src_location_util_id", SHUse, WBUse)
    i = RowNextColumn(1, SHNew, WBNew)
    Workbooks(WBUse).Worksheets(SHUse).Columns(useCol).Copy Destination:=Workbooks(WBNew).Worksheets(SHNew).Cells(1, i)
    
    useCol = FindColumnHeader("src_device", SHUse, WBUse)
    i = RowNextColumn(1, SHNew, WBNew)
    Workbooks(WBUse).Worksheets(SHUse).Columns(useCol).Copy Destination:=Workbooks(WBNew).Worksheets(SHNew).Cells(1, i)
    
    useCol = FindColumnHeader("src_addr_line1", SHUse, WBUse)
    i = RowNextColumn(1, SHNew, WBNew)
    Workbooks(WBUse).Worksheets(SHUse).Columns(useCol).Copy Destination:=Workbooks(WBNew).Worksheets(SHNew).Cells(1, i)
    
    useCol = FindColumnHeader("src_city", SHUse, WBUse)
    i = RowNextColumn(1, SHNew, WBNew)
    Workbooks(WBUse).Worksheets(SHUse).Columns(useCol).Copy Destination:=Workbooks(WBNew).Worksheets(SHNew).Cells(1, i)
    
    'useCol = FindColumnHeader("src_latitude", SHUse, WBUse)
    'i = RowNextColumn(1, SHNew, WBNew)
    'Workbooks(WBUse).Worksheets(SHUse).Columns(useCol).Copy Destination:=Workbooks(WBNew).Worksheets(SHNew).Cells(1, i)
    
    'useCol = FindColumnHeader("src_longitude", SHUse, WBUse)
    'i = RowNextColumn(1, SHNew, WBNew)
    'Workbooks(WBUse).Worksheets(SHUse).Columns(useCol).Copy Destination:=Workbooks(WBNew).Worksheets(SHNew).Cells(1, i)
    
    useCol = FindColumnHeader("src_dist_net_transformer_util_id", SHUse, WBUse)
    i = RowNextColumn(1, SHNew, WBNew)
    Workbooks(WBUse).Worksheets(SHUse).Columns(useCol).Copy Destination:=Workbooks(WBNew).Worksheets(SHNew).Cells(1, i)
    
    useCol = FindColumnHeader("event_text", SHUse, WBUse)
    i = RowNextColumn(1, SHNew, WBNew)
    Workbooks(WBUse).Worksheets(SHUse).Columns(useCol).Copy Destination:=Workbooks(WBNew).Worksheets(SHNew).Cells(1, i)
    
    
    Call ColumnWidthAutoMax(SHNew)
    Workbooks(WBNew).Worksheets(SHNew).Activate

    Workbooks(WBNew).Worksheets(SHNew).Name = "UnderVoltage"

    Call PhaseText

    Call UnderVoltageProcess
    
    Call ColumnWidthAutoMax
    Call HeaderBold
    
    Call UsageTracker("Under Voltage", "Finished")
    
    MsgBox "Under Voltage Processing Finished."
End Sub

Sub UnderVoltageProcess()
Dim aCol As Integer
Dim newCol As Integer
Dim botRow As Long
Dim aRange As range
Dim count As Integer

    ActiveSheet.Name = "UnderVoltage"

    aCol = FindColumnHeader("event_time")
    newCol = ColumnInsertRight(aCol)
    botRow = ColumnLastRow(aCol)
    
    Set aRange = range(Cells(2, newCol), Cells(botRow, newCol))
    fff = "=mid(" & Cells(2, aCol).Address(False, False) & ",12,8)"
    aRange.Formula = fff
    Call RangeToValues(aRange)
    Cells(1, newCol) = "EventTime"
    Columns(newCol).AutoFit
    
    newCol = ColumnInsertRight(aCol)
    Set aRange = range(Cells(2, newCol), Cells(botRow, newCol))
    fff = "=LEFT(" & Cells(2, aCol).Address(False, False) & ",10)"
    aRange.Formula = fff
    Call RangeToValues(aRange)
    Cells(1, newCol) = "RunDate"
    Columns(newCol).AutoFit
    
    Columns(aCol).Delete
    
    timeCol = FindColumnHeader("EventTime")
    countCol = ColumnInsertRight(timeCol)
    botRow = ColumnLastRow(timeCol)
    Cells(1, countCol) = "EventCount"
    
    meterCol = FindColumnHeader("src_name")
    countRow = 2
    count = 1
    i = 2
    
    Call SortSheetUp(meterCol, timeCol)
    
    While Cells(i, meterCol) <> ""
        If Cells(i, meterCol) <> Cells(i + 1, meterCol) Then
            
            Cells(i, countCol) = count
            count = 1
            i = i + 1
        Else
            Rows(i + 1).Delete
            count = count + 1
        End If
    Wend
    
    Call DeleteColumnUseHeader("event_log_id")
    'Call DeleteColumnUseHeader("RunDate")
    Call DeleteColumnUseHeader("EventTime")
    Call DeleteColumnUseHeader("event_text")
    
    installCol = FindColumnHeader("src_location_util_id")
        Cells(1, installCol) = "Installation_Num"
    meterCol = FindColumnHeader("src_name")
        Cells(1, meterCol) = "METER_SERIAL_NUM"
    deviceCol = FindColumnHeader("src_device_type")
        Cells(1, deviceCol) = "DeviceType"
    Call MeterDeviceTypeLookup(meterCol, deviceCol)
    
    countCol = FindColumnHeader("EventCount")
    Call SortSheetDown(countCol)
    
    MsgBox "Under Voltage Processing Finished."
End Sub

Sub MarkieRevenue()

    Call UsageTracker("MarkieRevenue", "Start")

    GLBQueryName = "MarkieRevenue"
    GlbStatusBarTxt = "Markie Revenue"
    
    GLBUserQuery = QueryBuilder("RevenueSelect", MACROWORKBOOK)
    Call Query(GLBUserQuery)

    Call UsageTracker("MarkieRevenue", "Finished")
    
    MsgBox "Markie Revenue Processing Finished."
End Sub

Sub UsageDrop()
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset
Dim useDate As String
Dim rateCol As Integer
    
10    Call UsageTracker("Usage Drop", "Start")
    
20    LOAD ChooseDateForm
30    ChooseDateForm.MonthView1.Value = format(Date, "m/d/yyyy")
40    ChooseDateForm.Show
50    If formCancel Then
60        Unload ChooseDateForm
70        Exit Sub
80    End If
      t = DateValue(ChooseDateForm.MonthView1.Value)
90    GlbUseDate = format(t, "YYYY-MM-DD")
100   Unload ChooseDateForm
    
110      useDate = GlbUseDate
    
120    qq = "SELECT UsageDropDate.startDate"
    
130    Set DBCn = DBCheckConnection(DBCn)
140    Set DBRs = DBCheckRecordset(DBRs)
    
150    With DBRs
160        .CursorLocation = adUseClient ' adUseServer
170        .CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
180        .LockType = adLockReadOnly ' adLockOptimistic
190        Set .ActiveConnection = DBCn
200    End With

210    On Error GoTo SkipDropTable  ' table not there
    
220    DBRs.Open qq, DBCn

DropTable:
    On Error GoTo 0
240    On Error GoTo gotError
    
250    qq = "Drop Table UsageDropDate"
        DBRs.Close

260    Set DBCn = DBCheckConnection(DBCn)
270    Set DBRs = DBCheckRecordset(DBRs)

280    DBRs.Open qq, DBCn
        
SkipDropTable:
290    On Error GoTo gotError
300    qq = "create volatile table UsageDropDate as( select cast('" & useDate & "' as DATE) as startDate) with data no primary index on commit preserve rows"
310    DBRs.Open qq, DBCn
    
320    SHUsagedrop = TD_USAGEDROP & "Select"
    
330    GLBUserQuery = QueryBuilder(SHUsagedrop, MACROWORKBOOK)
340    GLBQueryName = "UsageDrop"
350    GlbStatusBarTxt = "Running UsageDrop"
    
    'GLBUserQuery = QueryBuilder("UsageDropSelect")
360    Call Query(GLBUserQuery)
370    If formCancel Then Exit Sub
    
380    rateCol = FindColumnHeader("Curr_RateCode")
    
390    Call SortSheetUp(rateCol)
    
400    Call ColumnValuesToTabs(rateCol)
    
410    Call UsageTracker("Usage Drop", "Finished")
    
420    MsgBox "Usage Drop Processing Finished."
    
430    Exit Sub
    
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="UsageDrop"
    Stop
    Resume Next

End Sub

Sub CTSnoop()
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset
Dim useDate As String
Dim rateCol As Integer
    
10    Call UsageTracker("CTSnoop", "Start")
    
20    LOAD ChooseDateForm
30    ChooseDateForm.MonthView1.Value = format(Date, "m/d/yyyy")
40    ChooseDateForm.Show
50    If formCancel Then
60        Unload ChooseDateForm
70        Exit Sub
80    End If
      t = DateValue(ChooseDateForm.MonthView1.Value)
90    GlbUseDate = format(t, "YYYY-MM-DD")
100   Unload ChooseDateForm
    
110      useDate = GlbUseDate
    
120    qq = "SELECT CTSnoopDate.startDate"
    
130    Set DBCn = DBCheckConnection(DBCn)
140    Set DBRs = DBCheckRecordset(DBRs)
    
150    With DBRs
160        .CursorLocation = adUseClient ' adUseServer
170        .CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
180        .LockType = adLockReadOnly ' adLockOptimistic
190        Set .ActiveConnection = DBCn
200    End With

210    On Error GoTo SkipDropTable  ' table not there
    
220    DBRs.Open qq, DBCn

DropTable:                          ' get rid of an old date table
230    On Error GoTo 0
240    On Error GoTo gotError
    
250    qq = "Drop Table CTSnoopDate"
        DBRs.Close

260    Set DBCn = DBCheckConnection(DBCn)
270    Set DBRs = DBCheckRecordset(DBRs)

280    DBRs.Open qq, DBCn
        
SkipDropTable:
290    On Error GoTo gotError
300    qq = "create volatile table CTSnoopDate as( select cast('" & useDate & "' as DATE) as startDate) with data no primary index on commit preserve rows"
310    DBRs.Open qq, DBCn
    
320    SHCTSnoop = TD_CTSNOOOP & "Select"
    
330    GLBUserQuery = QueryBuilder(SHCTSnoop, MACROWORKBOOK)
340    GLBQueryName = "CTSnoop"
350    GlbStatusBarTxt = "Running CTSnoop"
    
    'GLBUserQuery = QueryBuilder("UsageDropSelect")
360    Call Query(GLBUserQuery)
370    If formCancel Then Exit Sub
    
380    rateCol = FindColumnHeader("Curr_RateCode")
    
390    Call SortSheetUp(rateCol)
    
400    Call ColumnValuesToTabs(rateCol)
    
410    Call UsageTracker("CTSnoop", "Finished")
    
420    MsgBox "CT Snoop Processing Finished."
    
430    Exit Sub
    
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="CTSnoop"
    Stop
    Resume Next

End Sub


Sub EventsTemplate()

    If ActiveWorkbook.Name = ThisWorkbook.Name And ActiveSheet.Name = "EventTemplate" Then
        EventsQueryForm.Show
    Else
        ThisWorkbook.Worksheets("EventTemplate").Activate
    End If

End Sub

