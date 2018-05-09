Attribute VB_Name = "Teradata_MOD"

Function TDTableExists(tableName) As Boolean
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

On Error GoTo gotError
10    qq = "SELECT TOP 1 * from " & tableName
        Debug_Print qq
20    Set DBCn = DBCheckConnection(DBCn)
30    Set DBRs = DBCheckRecordset(DBRs)

40    With DBRs
50        .CursorLocation = adUseClient ' adUseServer
60        .CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
70        .LockType = adLockReadOnly ' adLockOptimistic
80        Set .ActiveConnection = DBCn
90    End With

100   DBRs.Open qq, DBCn
    
110   fieldCount = DBRs.Fields.count
      TDTableExists = True
120   Exit Function

gotError:
    k = InStr(Err.Description, "does not exist")
    If k > 0 Then
        TDtablexists = False
        Exit Function
    End If
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:=" "
    Stop
    Resume Next
End Function

