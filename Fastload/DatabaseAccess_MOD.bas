Attribute VB_Name = "DatabaseAccess_MOD"

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
    'Set DBGlbConnection = Nothing
    Resume Next
    
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


