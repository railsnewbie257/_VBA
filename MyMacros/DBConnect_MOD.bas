Attribute VB_Name = "DBConnect_MOD"
Function ConnectDB()
Set DBCn = New ADODB.Connection

    Set DBCn = DBCheckConnection(DBCn)
    
    Call StatusbarDisplay("DBConnect: Connected")
End Function

Sub ConnectDBState()
    On Error Resume Next
    If (Not DBConnect.State = adStateOpen) Then
       s = "Connection Closed."
    Else
        s = "Connection OPEN" & vbNewLine
    End If
    s = s & vbNewLine & "Username: ~" & userName & "~"
    s = s & vbNewLine & "Password: ~" & Password & "~"
    
    MsgBox s
End Sub


