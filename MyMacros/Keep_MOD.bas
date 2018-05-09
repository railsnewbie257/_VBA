Attribute VB_Name = "Keep_MOD"
'
' This routines will save meter information into a MeterKeeps table for follow up
'
Sub LG_KEEP()
Dim keepRow As Long

If ActiveSheet.Name = "Keep" Or ActiveSheet.Name = "MeterKeeps" Then
    Call KeepDelete
    Exit Sub
End If

On Error GoTo gotError

    'Call ScreenOff
10    SHOrig = ActiveSheet.Name
20    origRow = ActiveCell.Row
    
    '
    ' Create the Keep sheet and form
    '
30    SHkeep = "Keep"
40    If Not SheetExists(SHkeep) Then
50        Sheets.Add
60        With ActiveSheet
70          .Name = SHkeep
80          .Tab.color = RED
90          .Cells(1, 1) = "Rundate"
100         .Cells(1, 2) = "EventTime"
110         .Columns(2).NumberFormat = "h:mm;@"
120         .Cells(1, 3) = "Meter_ID"
130         .Cells(1, 4) = "Installation_num"
140         .Cells(1, 5) = "Reason"
150         .Rows(1).Font.Bold = True
160      End With
170    End If
    
    'Worksheets(SHOrig).Activate
    '
    ' get the dates and times to fill in Keep form
    '
180    rundateCol = FindColumnHeader("rundate", SHOrig)
190    timeCol = FindColumnHeader("First_Event_Time_12007", SHOrig)
200    meterCol = FindColumnHeader("meter_serial_num", SHOrig)
210    installCol = FindColumnHeader("installation_num", SHOrig)
    
220    If ActiveSheet.Name = "Proximity" Then
230        botRow = ColumnLastRow(1, "Disconnected")
240        useRow = Worksheets("Disconnected").Cells(botRow, 1)
250        useRow = Worksheets("Disconnected").Cells(useRow, 1)
260        Rundate = Cells(useRow, rundateCol)
270        useTime = Cells(useRow, timeCol)

280        useMeter = Cells(useRow, meterCol)
290    Else
300        useRow = ActiveCell.Row
310        Rundate = Worksheets(SHOrig).Cells(origRow, rundateCol)
320       If timeCol > 0 Then
330            useTime = Worksheets(SHOrig).Cells(useRow, timeCol)
340        Else
350            useTime = 0
360        End If
370        useMeter = Worksheets(SHOrig).Cells(origRow, meterCol)
380        useInstall = Worksheets(SHOrig).Cells(origRow, installCol)
390    End If
    '
    ' Save to Keep sheet
    '
400    keepRow = LastRow(SHkeep) + 1
410    With Worksheets(SHkeep)
420        .Cells(keepRow, 1) = Rundate
430        .Cells(keepRow, 2) = useTime
440        .Cells(keepRow, 3) = useMeter
450        .Cells(keepRow, 4) = useInstall
460        .Cells(keepRow, 5) = SHOrig
470    End With
    '
    ' Save to Database
    '
480    t = IdentifyWorkbookType()
490    Call KeepSave(Rundate, useInstall, useMeter, IdentifyWorkbookType(), ActiveSheet.Name, "")
    

    'useRow = Worksheets(SHOrder).Cells(botRow, 1)
    
500    Call WriteTicket(SHOrig, origRow)
    
    'note = "Check for Fraud"

510    Call ClearClipboard
520    Worksheets(SHOrig).Activate
530    Call ScreenOn
    Exit Sub
    
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="LG_KEEP"
    Stop
    Resume Next
End Sub

Sub WriteTicket(SHOrig, origRow)
    SHkeep = "Keep"
    SHticket = "Ticket"
        
    If Not SheetExists(SHticket) Then
        Sheets.Add
        With ActiveSheet
            .Name = SHticket
            .Tab.color = RED
            .Cells(1, 1) = "Meter ID"
            .Cells(1, 2) = "Note"
            .Rows(1).Font.Bold = True
        End With
    End If
    
    botRow = LastRow(SHticket)
    
    note = ""
    t = IdentifyWorkbookType()
    Select Case t
    
    Case "LastGasp"
        rundateCol = FindColumnHeader("RunDate", SHkeep)
        timeCol = FindColumnHeader("eventtime", SHkeep)
        meterCol = FindColumnHeader("meter_serial_num", SHOrig)
        keepRow = ColumnLastRow(rundateCol, SHkeep)
        useDate = format(Worksheets(SHkeep).Cells(keepRow, rundateCol), "[$-409]mmmm d, yyyy;@")
        useTime = format(Worksheets(SHkeep).Cells(keepRow, timeCol), "[$-F400]h:mm:ss AM/PM")
        'Worksheets(SHticket).Cells(1, botRow) = Worksheets(SHkeep).Cells(keepRow, meterCol)
        note = "Check for Fraud // METER REPORTS LAST GASP/" & useDate & " - " & useTime & "//CUST HERE//SECURE EQUIP AND SEND CHARGES TO CLAIMS."
    
    Case "UsageDrop"
        pctCol = FindColumnHeader("PCT_CHG", SHOrig)
        meterCol = FindColumnHeader("meter_serial_num", SHOrig)
        note = "SUSPECT HIGH OR LOW USAGE METER HAS (DROPPED " & Worksheets(SHOrig).Cells(origRow, pctCol) & "%) NOTE ALL WORK PERFORMED."
            
    Case "PhaseAngleAlarm"
        meterCol = FindColumnHeader("meter_serial_num", SHOrig)
        note = "PHASE ANGLE ALARM"
        
    Case "UnderVoltage"
        meterCol = FindColumnHeader("meter_serial_num", SHOrig)
        note = "UNDERVOLTAGE"
        
    Case "ReceivedEnergy"
        meterCol = FindColumnHeader("meter_serial_num", SHOrig)
        note = "RECEIVED ENERGY"
        
    Case "ZeroKWH"
        meterCol = FindColumnHeader("meter_serial_num", SHOrig)
        note = "ZERO KWH"
        
    Case Else
        MsgBox "*** Unknown Workbook Type ***"
        
    End Select
    
    Worksheets(SHticket).Cells(botRow + 1, 2) = note
    Worksheets(SHticket).Cells(botRow + 1, 1) = Worksheets(SHOrig).Cells(origRow, meterCol)
    
End Sub

Sub KeepSave(useDate, useInstall, useMeter, Optional useReport, Optional useTab, Optional useNote)
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset
    
    On Error GoTo gotError
    
10    If IsMissing(useReport) Then useReport = ""
20    If IsMissing(useTab) Then useTab = ""
30    If IsMissing(useNote) Then useNote = ""
    
35    userName = LCase(Environ$("Username"))
        
40    t = "'" & format(useDate, "yyyy-mm-dd") & "','" & useInstall & "','" & useMeter & "','" & useReport & "','" & useTab & "','" & userName & "','" & useNote & "'"
50    qq = "INSERT INTO " & "dl_oge_analytics.meter_keeps" & " VALUES (" & t & ")"
    Debug_Print qq

60    Set DBCn = DBCheckConnection(DBCn)
70    Set DBRs = DBCheckRecordset(DBRs)

80    With DBRs
90        .CursorLocation = adUseClient  'adUseServer
100       .CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
110       .LockType = adLockOptimistic ' adLockReadOnly
120       Set .ActiveConnection = DBCn
130    End With

140    DBRs.Open qq, DBCn
150    'DBRs.Close
        MsgBox "Meter  " & useMeter & "  Kept", Title:="KeepSave"
160    Exit Sub
    
gotError:
    k = InStr(Err.Description, "Duplicate row")
    If k > 0 Then
        MsgBox "Duplicate meter entry", Title:="KeepSave"
        Exit Sub
    Else
        MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="KeepSave"
    End If

    DBGlbAdodbError = True
    Stop
    Resume Next
    
End Sub

Sub KeepDelete()
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

    SHOrig = ActiveSheet.Name

    On Error GoTo gotError
    
10    useRow = ActiveCell.Row
    
20    dateCol = FindColumnHeader("RunDate")
30    meterCol = FindColumnHeader("Meter_ID")
      If meterCol = -1 Then meterCol = FindColumnHeader("Meter_Serial_Num")
    
40    useDate = format(Cells(useRow, dateCol), "'YYYY-MM-DD'")
50    useMeter = Cells(useRow, meterCol)
    
    qq = "DELETE FROM dl_oge_analytics.meter_keeps where RunDate = " & useDate & " AND Meter_Serial_Num = '" & useMeter & "' AND UserName = '" & LCase(Environ$("Username")) & "'"
    Debug.Print qq
60    Set DBCn = DBCheckConnection(DBCn)
70    Set DBRs = DBCheckRecordset(DBRs)

80    With DBRs
90        .CursorLocation = adUseClient  'adUseServer
100       .CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
110       .LockType = adLockOptimistic ' adLockReadOnly
120       Set .ActiveConnection = DBCn
130    End With
        retCode = MsgBox("Are you sure you want to delete  " & useMeter & "  ?", vbYesNo, Title:="KeepDelete")
        If retCode = vbNo Then Exit Sub
        
140    DBRs.Open qq, DBCn
150    'DBRs.Close
        MsgBox "Keep for  " & useMeter & "  Deleted", Title:="KeepDelete"
        Worksheets(SHOrig).Rows(useRow).Delete
160    Exit Sub

gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="KeepDelete"

    DBGlbAdodbError = True
    Stop
    Resume Next
End Sub

