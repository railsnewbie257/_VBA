Attribute VB_Name = "DatabaseAccess_MOD"
Sub InitDatabaseAndTables()
    ReDim Preserve GLBDatabaseNameList(1)
    GLBDatabaseNameList(UBound(GLBDatabaseNameList)) = "dl_oge_analytics"
    ReDim Preserve GLBDatabaseNameList(UBound(GLBDatabaseNameList) + 1)
    GLBDatabaseNameList(UBound(GLBDatabaseNameList)) = "da_customer_vw"
    
    ReDim Preserve GLBTableNameList(1)
    GLBTableNameList(UBound(GLBTableNameList)) = "billing_statement_charge"
End Sub

Function DBMakeConnection(DBConn)
    If (Len(userName) = 0 Or Len(Password) = 0) And (Not DBConn.State = adStateOpen) Then
        LoginForm.Show
    End If
    If formCancel Then
        If Not DBConn Is Nothing Then Set DBConn = Nothing
        Exit Function
    End If
    '
    ' Connection string
    '
    s = "DSN=OGE;Databasename=dbc;Uid=" & userName & ";PWD=" & Password & ";Authentication Mechanism=LDAP;"

    Debug_Print s
    Call StatusbarShow("DBMakeConnection: Open")
    DBConn.Open s
    If DBConnect.State = adStateOpen Then 'If connection is success, continue
        Call StatusbarDisplay("DBMakeConnection: Connected to Database")
        Application.ODBCTimeout = 900
    End If
End Function

Sub DBConnectionProperties()
Dim DBCn As ADODB.Connection
    Set DBCn = DBCheckConnection(DBCn)
    
    For i = 0 To DBCn.Properties.count - 1
        Cells(i + 1, 1) = DBCn.Properties(i).Name
        Cells(i + 1, 2) = DBCn.Properties(i).Attributes
        Cells(i + 1, 3) = DBCn.Properties(i).Value
    Next i
    
    Cells(i + 1, 1) = "Command Timeout"
    Cells(i + 1, 2) = ""
    Cells(i + 1, 3) = DBCn.CommandTimeout
    
    Application.odbc
    
End Sub

Sub TestDBCheckConnection()
Dim DBCn As ADODB.Connection

    Set DBGlbConnection = Nothing

    Set DBCn = DBCheckConnection(DBCn)
    
    Debug_Print DBCn.Properties.count
    
    
    For i = 0 To DBCn.Properties.count
        debug_print i & ") " & DBCn.Properties(i).Name & " " & DBCn.Properties(i).Attributes & " "; DBCn.Properties(i).Value
        
    Next i
    Set DBCn = DBCheckConnection(DBCn)
    
    Set DBCn = DBCheckConnection(DBCn)
    
    Set DBGlbConnection = DBCn
End Sub

'
' This routine is a workhorse
' It checks to see if the provided object is connected
' if not it checks if the global object is connected
' if so, it uses the global connection
' otherwise, it opens a new connection and saves to the global connection
'
Function DBCheckConnection(Optional DBConn)
Dim haderror As Boolean

    Call StatusbarDisplay("DBCheckConnection: Check is Nothing.")
    haderror = False
    If Not DBConn Is Nothing Then
        Set DBCheckConnection = DBConn
        Exit Function
    End If
    If DBGlbConnection Is Nothing Then
        Call StatusbarDisplay("DBCheckConnection: Allocate New.")
        Set DBConn = New ADODB.Connection
    Else
        Set DBConn = DBGlbConnection
    End If
    
    Call StatusbarDisplay("DBCheckConnection: Check Open or Closed")
    If DBConn.State = adStateClosed Then
        userName = LCase(Environ$("Username"))
        Password = Workbooks(MACROWORKBOOK).Sheets("Pallette").Cells(1, 1)
        If (Len(userName) = 0 Or Len(Password) = 0) Or Password = "" Then
            LoginForm.Show
            If formCancel Then
                Set DBCheckConnection = Nothing
                Exit Function
            End If
        End If
        
        Call StatusbarDisplay("DBCheckConnection: Opening...")
        
        loginString = "DSN=OGE;Databasename=dbc;Uid=" & userName & ";PWD=" & Password & ";Authentication Mechanism=LDAP;"
        'loginString = "DSN=OGE2;"
        
        On Error GoTo LoginError
        DBConn.ConnectionTimeout = 0 'To wait till the query finishes without generating error
        
        DBConn.Open loginString
        Call StatusbarDisplay("DBCheckConnection: Config")
        Application.ODBCTimeout = 900
        DBConn.CommandTimeout = 1200
        '
        ' Save Password
        '
        If Not haderror Then
            With Workbooks(MACROWORKBOOK).Sheets("Pallette")
                .Cells(1, 1) = Password
                .Cells(1, 1).Font.ThemeColor = xlThemeColorDark1
                .Cells(1, 1).Font.TintAndShade = 0
            End With
        Else
            DBCheckConnection (DBConn)
            Set DBCheckConnection = Nothing
            DBConn.Close
            Set DBConn = Nothing
            Exit Function
        End If
    End If
    
    Call StatusbarDisplay("DBCheckConnection: Opened")
    Set DBGlbConnection = DBConn
    Set DBCheckConnection = DBConn
    Exit Function
    
LoginError:
    MsgBox "DBCheckConnection: " & vbNewLine & Err.Description & vbNewLine & vbNewLine & loginString, Title:="Login Error"
    ThisWorkbook.Sheets("Pallette").Cells(1, 1) = "" ' only way to correct an incorrect Password
    haderror = True
    On Error GoTo 0
    Resume Next
    
End Function

Function DBCloseConnection(Optional DBConn)
    If IsMissing(DBConn) Then Set DBConn = DBGlbConnection
    If Not DBConn Is Nothing Then
        If DBConn.State <> 0 Then DBConn.Close
        Set DBConn = Nothing
        If DBGlbConnection.State <> 0 Then DBGlbConnection.Close
        Set DBGlbConnection = Nothing
        On Error Resume Next
        ' DBConn.Close
        ' Set DBConn = Nothing
    End If
    MsgBox "Database Connection Reset", Title:="DBCloseConnection"
End Function

Function DBCheckRecordset(DBRecordset)
    Call StatusbarDisplay("DBCheckRecordset: Check for Nothing.")
    If DBRecordset Is Nothing Then
        Call StatusbarDisplay("DBCheckRecordset: Allocate New.")
        Set DBCheckRecordset = New ADODB.Recordset
    Else
        Set DBCheckRecordset = DBRecordset
    End If
    Call StatusbarDisplay("DBCheckRecordset: Return.")
End Function

Function DBCloseRecordset(DBRecordset)
    If Not DBRecordset Is Nothing Then
        DBRecordset.Close
        Set DBRecordset = Nothing
    End If
End Function

Sub ShowRecordset(rst)
    Dim DBRs As ADODB.Recordset
    
End Sub


Sub MyDB()
    
    Set db_cn = New ADODB.Connection
    Set record_set = New ADODB.Recordset
    
    db_cn.ConnectionTimeout = 0 'To wait till the query finishes without generating error
    db_cn.CommandTimeout = 1200
    login_string = "DSN=OGE;Databasename=dbc;Uid=loginid;PWD=password;Authentication Mechanism=LDAP;"
    login_string = "DSN=OGE;"
    
    db_cn.Open login_string

    DBName = "dl_oge_analytics"
    DBName = "dbc"
    tblname = "mp_01_cons_large_project_report"
    tblname = "dbcinfo"
    
    useQuery = "select * from " & DBName & "." & tblname & ";"
                ' "WHERE Worksheet_Year = " & "'" & ComboBox1.Text & "'" & " AND Worksheet_Month = " & "'" & ComboBox2.Text & "'" & " order by 1; "
    
    Debug_Print useQuery
    ' debug_print adStateOpen
    ' debug_print db_cn.State
    record_set.Open useQuery & ";", db_cn 'Issue SQL statement'
 
    Debug_Print record_set.Fields.count
    record_set.Close
End Sub


