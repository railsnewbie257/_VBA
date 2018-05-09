Attribute VB_Name = "Recordset_MOD"
Sub MyRecordset()
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

On Error GoTo gotError
10    qq = "SELECT TOP 3 * from putlvw.EUL_POS_METERS_D"

20    Set DBCn = DBCheckConnection(DBCn)
30    Set DBRs = DBCheckRecordset(DBRs)

40    With DBRs
50        .CursorLocation = adUseServer  ' adUseClient '
60        .CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
70        .LockType = adLockReadOnly ' adLockOptimistic
80        Set .ActiveConnection = DBCn
90    End With

100   DBRs.Open qq, DBCn

110   Sheet1.Range("A1").CopyFromRecordset DBRs
    
    MsgBox "Finished"
    
    DBRs.Close
    Set DBRs = Nothing
    DBCn.Close
    Set DBCn = Nothing
    
   Exit Sub
   
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:=" "
    Stop
    Resume Next
End Sub

Application
