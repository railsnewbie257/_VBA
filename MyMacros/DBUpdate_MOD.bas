Attribute VB_Name = "DBUpdate_MOD"


Sub DBCheckLastGasp()
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

    userQuery = "select RunDate from dl_oge_analytics.Last_Gasp where RunDate = DATE-1"
    Debug.Print userQuery
    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    
    Set DBRs.ActiveConnection = DBCn
    
    On Error GoTo GotErr
    DBRs.Open userQuery
    
    If DBRs.BOF And DBRs.EOF Then
        MsgBox "Database needs updating."
    Else
        MsgBox "Dataabse is updated."
    End If
    
GotErr:
    Debug.Print format(Err.Number, "0") & " " & Err.Description
    Call DBCloseRecordset(DBRs)
    Exit Sub
End Sub

