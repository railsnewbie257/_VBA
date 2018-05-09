Attribute VB_Name = "DBQuery_MOD"

Sub Query()
Dim objRecordset

    On Error Resume Next
        If Not DBConnect.State = adStateOpen Then
            On Error GoTo 0
            Call ConnectDB
        End If
    
    Set objRecordset = CreateObject("ADODB.Recordset")
    
    QueryForm.Show
    If formCancel Then Exit Sub
    Debug.Print "1 - " & Now()
    On Error Resume Next
        objRecordset.Close
    On Error GoTo 0
    
    With objRecordset
        .CursorLocation = adUseClient ' adUseServer
        .CursorType = adOpenStatic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
        .LockType = adLockOptimistic ' adLockReadOnly
        Set .ActiveConnection = DBConnect
    End With

    objRecordset.Open userQuery
    'MsgBox Format(objRecordset.Fields(0).Value, "#,##0") & " records"
    MsgBox objRecordset.recordCount
    objRecordset.Close
    
    Set Recordset = New ADODB.Recordset
    Debug.Print "3 - " & Now()
    Recordset.Open userQuery, DBConnect
    'Debug.Print "4 - " & Now()
    'Debug.Print "Record Count: " & Recordset.RecordCount
    retCode = MsgBox("Field count: " & Recordset.Fields.Count & vbNewLine & "Download Header Names?", vbYesNoCancel)
    If retCode = vbCancel Then Exit Sub
    '
    ' download header
    '
    fieldCount = Recordset.Fields.Count
    useRow = ColumnLastRow(1)
    For i = 0 To fieldCount - 1
        Cells(useRow, i + 1) = Recordset.Fields(i).Name
    Next i
    
    retCode = MsgBox("Download Data?", vbYesNoCancel)
    If retCode = vbNo Or retCode = vbCancel Then
        Recordset.Close
        Exit Sub
    End If
    
    useRow = ColumnLastRow(1) + 1
    'For i = 1 To 100
    While Not Recordset.EOF
        For iCol = 1 To fieldCount
            'Application.CutCopyMode = False
            If Not (iCol = 67) Then Cells(useRow, iCol) = Recordset.Fields(iCol - 1).Value
        Next iCol
        Recordset.MoveNext
        useRow = useRow + 1
        Debug.Print useRow
    Wend
    'Next i
    Recordset.Close
    
    Columns(1).AutoFit
    
    MsgBox "Teradata Download Finished."
End Sub

    

Function Query2()
    Dim objRS

    Set objRS = CreateObject("ADODB.Recordset")
    With objRS
        ' Cursor properties
        .CursorLocation = adUseServer
        .CursorType = adOpenForwardOnly
        ' set lock type
        .LockType = adLockReadOnly
        ' set connection for Recordset
        Set .ActiveConnection = DBConnect
        ' Get record count
        QueryForm.Show
        .Open userQuery
        MsgBox Format(.Fields(0).Value, "#,##0") & " records"

        ' Now close Recordset and reopen with data to be processed
        If .State = adStateOpen Then .Close
        '.Open "select account, type from account"

        ' Do some useless processing...
        'While Not (.BOF Or .EOF)
        '      Listbox1.Items.Add .Fields("account").Value & ""
        '      .MoveNext
        'Wend
        On Error Resume Next
            .Close
        On Error GoTo 0
    End With
    Set objRS = Nothing
End Function

Function DoBigQuery()

selectQuery = _
"SELECT e.Event_Start_Dt AS RunDate, " & _
"m.EQUIP_MFG_SERIAL_NUMBER," & _
"e.Event_External_Id," & _
"Min(Cast(e.Event_Start_Tm AS CHAR(15))) AS First_Event_Time," & _
"Max(Cast(e.Event_Start_Tm AS CHAR(15))) AS Last_Event_Time," & _
"sme.SERVICE_POINT_ID," & _
"sme.METER_ID," & _
"m.INSTALLATION_NUMBER," & _
"m.PREMISE_NUMBER," & _
"m.METER_INSTALLATION_DATE," & _
"m.METER_REMOVAL_DATE," & _
"m.METER_LATITUDE," & _
"m.METER_LONGITUDE," & _
"m.METER_ACTIVE_STATUS_CODE," & _
"m.METER_ACTIVE_STATUS_DESC," & _
"m.METER_MAC_ADDRESS," & _
"m.PHASE_COUNT," & _
"m.VOLTAGE_MEASURE," & _
"m.Circuit_Id," & _
"m.CIRCUIT_NUMBER "

fromQuery = _
"FROM    putlvw.EVENT e "

joinQuery = _
"JOIN    putlvw.SERVICE_METER_EVENT sme ON e.Event_Id = sme.Event_Id " & _
"JOIN    putlvw.EUL_POS_METERS_D m ON sme.METER_ID = m.METER_ID "

whereQuery = _
"WHERE e.Event_External_Event_Cd = '12007' "

condition1Query = "e.Event_Start_Dt = '2017-08-09' "

condition2Query = "m.EQUIP_MFG_SERIAL_NUMBER = '45217748G' "

groupQuery = _
"GROUP BY 1,2,3,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20 --ORDER BY 1"

DoBigQuery = selectQuery & fromQuery & joinQuery & whereQuery & " AND " & cond1Query & " AND " & cond2Query & groupQuery

Debug.Print DoBigQuery
End Function
