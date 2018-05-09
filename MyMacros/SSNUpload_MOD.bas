Attribute VB_Name = "SSNUpload_MOD"
Sub InsertTest()
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)

    insertString = "INSERT INTO dl_oge_analytics.SSN_METER_STATUS VALUES"
    On Error GoTo gotError
    s = insertString
    Call StartTimer
    For i = 1 To 3
        If i > 1 Then s = s & ","
        s = s & " (" & format(i, 0) & ", 'test', 'test')"
    Next i
    s = s & ";"
    
    Debug.Print s
    Set DBRs.ActiveConnection = DBCn
    DBRs.Open s
    
    Call StopTimer
    MsgBox ElapsedTime
    Exit Sub
gotError:
    'Debug.Print Err.Description
    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    Resume Next
End Sub


Sub SSNUploadMeterStatus()
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset
    
    WBOrig = ActiveWorkbook.Name
    SHOrig = ActiveSheet.Name
    '
    ' Get the date from the file name
    '
    dateCol = FindColumnHeader("event_time")
    statusCol = FindColumnHeader("src_ops_state")
    meterCol = FindColumnHeader("src_name")
    
    useTable = "dl_oge_analytics.SSN_METER_STATUS"
    
    botRow = LastRow(SHOrig, WBOrig)
    
    For i = DATASTARTROW To botRow
        Rows(i).Copy
        useStatus = Cells(i, statusCol)
        useMeter = Cells(i, meterCol)
        useDate = left(Cells(i, dateCol), 10)
        'useDate = Format(useDate, "m/d/yyyy")
        
        useData = "(CAST('" & useDate & "' AS DATE), '" & useMeter & "', '" & useStatus & "')"

        userQuery = "INSERT INTO " & useTable & " VALUES " & useData
        'Debug.Print userQuery
        
        Set DBCn = DBCheckConnection(DBCn)
        Set DBRs = DBCheckRecordset(DBRs)
        Set DBRs.ActiveConnection = DBCn

        On Error GoTo GotErr
        DBRs.Open userQuery
        'DBrs.

        If i Mod 100 = 0 Then
            Call StopTimer
            Debug.Print i & " " & ElapsedTime
            Call StatusbarDisplay(format(i, "#,##0") & " / " & format(botRow, "#,##0"))
            Call StartTimer
            k = 1
        End If
    Next i
    
     Call StopTimer
            Debug.Print ElapsedTime
    Call StatusbarDisplay("Finished")
    Exit Sub
GotErr:
    'Debug.Print "Error: ~" & Err.Description & "~"

    'Set DBCn = DBCheckConnection(DBCn)
    'Set DBRs = DBCheckRecordset(DBRs)
    Resume Next

End Sub
