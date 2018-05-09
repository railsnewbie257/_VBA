Attribute VB_Name = "UsageTracker_MOD"
Sub UsageTracker(Comment1, Optional Comment2)
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    Set DBRs.ActiveConnection = DBCn
    
    useTable = "dl_oge_analytics.UsageTracker"
    
    userName = LCase(Environ$("Username"))
    
    If Not IsMissing(Comment2) Then
        s = "INSERT INTO " & useTable & " VALUES (" & "'" & userName & "','" & Comment1 & "','" & Now() & "','" & Comment2 & "')"
    Else
        s = "INSERT INTO " & useTable & " VALUES (" & "'" & userName & "','" & Comment1 & "','" & Now() & "','')"
    End If
    Debug_Print s
    On Error GoTo GotErr
        DBRs.Open s

    Exit Sub
    
GotErr:
    Debug_Print i & "    Error: ~" & Err.Description & "~"
    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    Set DBRs.ActiveConnection = DBCn
    Resume Next
End Sub

