VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QueryForm 
   Caption         =   "SQL Query Form"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14340
   OleObjectBlob   =   "QueryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "QueryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ThisQuerySheet As Boolean
Dim cBar As clsBar

Private Function fullTableName()
    If tbUseSpreadsheet Then
        useCol = ActiveCell.Column
        fullTableName = left(Cells(1, useCol), Len(Cells(1, useCol)) - 2)
    Else
        fullTableName = cbDatabaseName & "." & cbTableName
    End If
End Function


Private Sub btnCreateTable_Click()
    Call DBCreateTableScript
    formCancel = True
    Unload Me
End Sub

Private Sub btnDefaultDatabase_Click()
    cbDatabaseName = "dl_oge_analytics"
End Sub

Private Sub btnLoadTable_Click()
    Call DBInsert
    Unload Me
End Sub

Private Sub btnPreviousQuery_Click()
    txtQuery = GLBUserQuery
End Sub

Private Sub btnQueries_Click()
    QueryForm2.Show
End Sub

Private Sub btnResetConnection_Click()
    Call DBCloseConnection
End Sub

Private Sub btnSubmit_Click()

    GLBUserQuery = txtQuery.Text
    formCancel = False
    
    If optTable Then
    ElseIf optView Then
        userQuery = "select * from da_customer_vw." & txtName & ";"
    ElseIf optLastGaspDaily Then
        userQuery = QueryBuilder("LastGasp2Select")
        GLBNewWorkbook = True
    ElseIf optZeroKwh Then
        GLBNewWorkbook = True
    ElseIf optUnderVoltage Then
        GLBNewWorkbook = True
    ElseIf optUsageDrop Then
        GLBNewWorkbook = True
    Else
        GLBUserQuery = txtQuery.Text
    End If
    
    GLBDatabaseName = cbDatabaseName  'txtDatabaseName
    Call DatabaseNameListAdd(cbDatabaseName)
    
    GLBTableName = cbTableName  'txtTableName
    Call TableNameListAdd(cbTableName)
    
    If txtQuery.Text = "" Then cancelform = True
    Unload QueryForm
    
    Call StatusbarDisplay(GlbStatusBarTxt)

End Sub

Private Sub optInfo_Click()
    txtQuery = "select * from dbc.dbcinfo;"
    GLBQueryName = "KV2CUnderVoltage"
    GlbStatusBarTxt = "Running KV2C UnderVoltage"
End Sub

Private Sub cbDatabaseName_AfterUpdate()
    Debug.Print "After Update"
End Sub

Private Sub optColumnNames_Click()
    txtQuery = "SELECT columnname FROM dbc.columns WHERE databasename='" & cbDatabaseName & "' and tablename='" & cbTableName & "' ORDER BY columnid;"
    GLBQueryName = "ColumnNames"
    GLBTableName = txtTableName
End Sub

Private Sub optDatabaseName_Click()
    cbDatabaseName = "dbc"
    cbTableName = "tables"
    txtQuery = "SELECT DISTINCT(databasename) FROM " & fullTableName & " WHERE tablekind = 'T' ORDER BY 1"
    cbDatabaseName = "dbc"
    GLBQueryName = "DatabaseNames"
End Sub

Private Sub optDatabaseTables_Click()

    txtQuery = "SELECT databasename, tablename, creatorname, lastaltertimestamp, commentstring FROM dbc.tables WHERE tablekind = 'T' "
    cbDatabaseName = "dbc"
    cbTableName = "tables"
    
    If cbDatabaseName <> "" Then txtQuery = txtQuery & " AND DatabaseName = " & "'" & cbDatabaseName & "'"
    txtQuery = txtQuery & " ORDER BY 1,2"
    GLBQueryName = "DatabaseTables"
End Sub

Private Sub optDatabaseViews_Click()
    txtQuery = "SELECT databasename, tablename, creatorname, lastaltertimestamp, commentstring FROM dbc.tables WHERE tablekind = 'V' ORDER BY 1,2"
    GLBQueryName = "DatabaseViews"
End Sub

Private Sub optDefaultDatabase_Click()
    cbDatabaseName = "dl_oge_analytics"
End Sub

Private Sub optRowCount_Click()
    txtQuery = "select count(*) from " & cbDatabaseName & "." & cbTableName
End Sub

Private Sub optTablePrefix_Click()
    cbTableName = LCase(Environ$("Username")) & "_"
End Sub

Private Sub optSelectAllFromTable_Click()
    txtQuery = "select * from " & cbDatabaseName & "." & cbTableName
End Sub

Private Sub optSelectTop5FromTable_Click()
    txtQuery = "select top 5 * from " & cbDatabaseName & "." & cbTableName
End Sub

Private Sub optUnderVoltage_Click()
    txtQuery = "select * from dl_oge_analytics.KV2C_Under_Voltage"
    GLBQueryName = "KV2CUnderVoltage"
    GlbStatusBarTxt = "Running KV2C UnderVoltage"
End Sub

Private Sub optUsageDrop_Click()
    txtQuery = "SELECT * FROM dl_oge_analytics.Usage_Drop ORDER BY 2 ;"
    GLBQueryName = "UsageDrop"
    GlbStatusBarTxt = "Running Usage Drop"
End Sub

Private Sub optZeroKWH_Click()
    txtQuery = "SELECT * FROM dl_oge_analytics.Zero_KWH ORDER BY 2;"
    GLBQueryName = "ZeroKwH"
    GlbStatusBarTxt = "Running Zero KwH"
End Sub

Private Sub txtDatabaseName_AfterUpdate()
    t = InStr(1, txtDatabaseName, ".")
    If t > 0 Then
        dbTableName = Right(cbDatabaseName, Len(cbDatabaseName) - t)
        cbDatabaseName = left(cbDatabaseName, t - 1)
    End If
End Sub

Private Sub txtTableName_AfterUpdate()
    t = InStr(1, dbTableName, ".")
    If t > 0 Then
        cbDatabaseName = left(cbTableName, t - 1)
        cbTableName = Right(cbTableName, Len(cbTableName) - t)
    End If
End Sub

Private Sub cbTableName_AfterUpdate()
    userName = LCase(Environ$("Username"))
    If Len(cbTableName) > 0 Then
        t = left(cbTableName, Len(userName)) ' check if username is already prefixed
        If userName = t Then
            tbTablePrefix.Value = True
        Else
            tbTablePrefix.Value = False
        End If
    Else
        tbTablePrefix.Value = False
        'cbTableName = LCase(Environ$("Username")) & "_"
    End If
End Sub

Private Sub tbTablePrefix_Click()
    userName = LCase(Environ$("Username"))
    If cbTableName = "" And tbTablePrefix.Value = False Then
        tbTablePrefix.Value = False
        cbTableName = ""
        Exit Sub
    End If
    If Len(cbTableName) > 0 Then
        t = left(cbTableName, Len(userName))
        If tbTablePrefix.Value = True Then
            If Not userName = t Then cbTableName = userName & "_" & cbTableName
        Else
            If userName = t Then cbTableName = Right(cbTableName, Len(cbTableName) - (Len(userName) + 1))
        End If
    Else
        cbTableName = LCase(Environ$("Username")) & "_"
    End If
End Sub

Private Sub txtQuery_Enter()
    optLastGaspDaily = False
    optLastGaspDates = False
    optBigQuery = False
    optThisQuerySheet = False
End Sub

Private Sub optLastGaspDates_Click()
    txtQuery = "SELECT UNIQUE(RunDate) FROM dl_oge_analytics." & TD_LASTGASP & " ORDER BY RunDate DESC;"
    GLBQueryName = "LastGaspDates"
    GlbStatusBarTxt = "Running Last Gasp Dates"
End Sub

Private Sub optLastGaspTable_Click()
    cbDatabaseName = "dl_oge_analytics"
    cbTableName = TD_LASTGASP
End Sub

Private Sub optZeroKWHTable_Click()
    cbDatabaseName = "dl_oge_analytics"
    cbTableName = "Zero_KWH"
End Sub

Private Sub optUndervoltageTable_Click()
    cbDatabaseName = "dl_oge_analytics"
    cbTableName = "KV2C_Under_Voltage"
End Sub

Private Sub optUsageDropTable_Click()
    cbDatabaseName = "dl_oge_analytics"
    cbTableName = "Usage_Drop"
End Sub

Private Sub optDropTable_Click()
    txtQuery = "DROP TABLE " & cbDatabaseName & "." & cbTableName
End Sub

Private Sub optDatabaseVersion_Click()
    txtQuery = "SELECT * FROM dbc.dbcinfo;"
    GLBQueryName = "DatabaseVersion"
    GlbStatusBarTxt = "Running Database Version"
End Sub

Private Sub optLastGaspReport_Click()
    SHQuery = "LastGasp2Select"
    If Len(txtStartDate) = 0 Then
        txtMessage = "Please enter Start Date"
        lblStartDate.ForeColor = &HFF&
        optLastGaspReport = False
    Else
        Call QueryNewCondition(SHQuery, MACROWORKBOOK, "RunDate =", "'" & txtStartDate & "'")
        txtQuery = QueryBuilder(SHQuery)
        GLBQueryName = "LastGasp"
        GlbStatusBarTxt = "Running Last Gasp Daily"
    End If
End Sub

Sub optEvents_Click()
    SHQuery = "Events"
    Call QueryNewCondition(SHQuery, MACROWORKBOOK, "m.EQUIP_MFG_SERIAL_NUMBER =", "'" & txtMeterNumber & "'")
    If txtStartDate <> "" Then
        Call QueryNewCondition(SHQuery, "RunDate =", "'" & txtStartDate & "'")
    End If
    txtQuery = QueryBuilder(SHQuery)
    
    cbDatabaseName = "putlvw"
    cbTableName = "EUL_POS_METERS_D"
    
    GLBQueryName = "Drilldown-Event"
    GlbStatusBarTxt = "Running Last Gasp Daily"
End Sub

Sub optBilling_Click()
    SHQuery = "Billing"
    Call QueryNewCondition(SHQuery, MACROWORKBOOK, "CUSTOMER_NUMBER =", "'" & txtCustomer & "'")
    Call QueryNewCondition(SHQuery, "RunDate =", "'" & txtDate & "'")
    txtQuery = QueryBuilder(SHQuery)
    GlbStatusBarTxt = "Running Billing"
End Sub

Sub optPremise_Click()
    SHQuery = "Premise"
    Call QueryNewCondition(SHQuery, MACROWORKBOOK, "RunDate =", "'" & txtDate & "'")
    Call QueryNewCondition(SHQuery, "CUSTOMER_NUMBER =", "'" & txtCustomer & "'")
    txtQuery = QueryBuilder(SHQuery)
    GLBDownloadByRow = False
    GlbStatusBarTxt = "Running Premise"
End Sub

Private Sub optReset_Click()
    txtQuery = "DELETE FROM dl_oge_analytics." & tdlastgasp & " WHERE RunDate = '" & txtDate & "'"
    GlbStatusBarTxt = "Resetting Database Table"
End Sub

Private Sub btnUseQuerySheet_Click()
    SHQuery = ActiveSheet.Name
    txtQuery = QueryBuilder()
    GlbStatusBarTxt = "Query Sheet"
    GLBQueryName = "UseQuerySheet"
    ThisQuerySheet = True
End Sub

Private Sub cmdClear_Click()
    txtMeterNumber = ""
    txtDate = ""
    txtQuery = ""
    txtCustomer = ""
End Sub
Private Sub btnCancel_Click()
    Debug.Print "QueryForm Cancel"
    formCancel = True
    Unload QueryForm
    Call StatusbarDisplay("Cancelled Query")
End Sub

Private Sub txtStartDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim t As Long

    Debug.Print "txtStartDate click"

    LOAD ChooseDateForm
    ChooseDateForm.MonthView1.Value = format(Date - 1, "m/d/yyyy")
    ChooseDateForm.Show
    If ChooseDateForm.optCancel Then
        Unload ChooseDateForm
        Exit Sub
    End If
    t = DateValue(ChooseDateForm.MonthView1.Value)
    txtStartDate = format(t, "YYYY-MM-DD")
    Unload ChooseDateForm
    lblStartDate.ForeColor = &H80000012
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        ' Your code
        ' Tip: If you want to prevent closing UserForm by Close (×) button in the right-top corner of the UserForm, just uncomment the following line:
        Cancel = True
    End If
End Sub

Private Sub UserForm_Initialize()

    useRow = Application.WorksheetFunction.Max(ActiveCell.Row, 2) ' do not use header
    
    meterCol = FindColumnHeader("meter_serial_num")
    If meterCol > 0 Then txtMeterNumber = Cells(useRow, meterCol)
    meterCol = FindColumnHeader("src_name")
    If meterCol > 0 Then txtMeterNumber = Cells(useRow, meterCol)
    '
    ' Date to use
    '
    dateCol = FindColumnHeader("RunDate")
    If dateCol > 0 Then txtStartDate = format(Cells(useRow, dateCol), "YYYY-MM-DD")
    dateCol = FindColumnHeader("Event_Start_Dt")
    If dateCol > 0 Then txtStartDate = format(Cells(useRow, dateCol), "YYYY-MM-DD")


    customerCol = FindColumnHeader("customer_number")
    If customerCol > 0 Then txtCustomer = Cells(useRow, customerCol)
    '
    ' Table Name ----------------------------------------------------------------------------------
    '
    If SheetExists("CreateTable") Then
        useTable = Replace(Worksheets("CreateTable").Cells(2, 1), ",", "")
        t = InStr(useTable, ".")
        cbDatabaseName = left(useTable, t - 1)
        Call DatabaseNameListAdd(cbDatabaseName)
        cbTableName = Right(useTable, Len(useTable) - t)
        Call TableNameListAdd(cbTableName)
    Else
        dbCol = FindColumnHeader("DatabaseName")
        If dbCol > 0 Then cbDatabaseName = Trim(Cells(useRow, dbCol))
        
        tableCol = FindColumnHeader("TableName")
        If tableCol > 0 Then cbTableName = Trim(Cells(useRow, tableCol))
        
        ' createtable = FindColumnHeader("Create Set Table")
        ' If createtable > 0 Then txtTableName = Replace(Cells(2, 1), ",", "")
    End If
    '
    '
    tbTablePrefix = False
    formCancel = False
    ThisQuerySheet = False
    GLBQueryName = ""
    '
    ' Database and Table Names
    '
    If IsArrayAllocated(GLBDatabaseNameList) Then
        For i = UBound(GLBDatabaseNameList) To 1 Step -1
            cbDatabaseName.AddItem GLBDatabaseNameList(i)
        Next i
        cbDatabaseName = GLBDatabaseNameList(UBound(GLBDatabaseNameList))
    End If
    If IsArrayAllocated(GLBTableNameList) Then
        For i = UBound(GLBTableNameList) To 1 Step -1
            cbTableName.AddItem GLBTableNameList(i)
        Next i
        cbTableName = GLBTableNameList(UBound(GLBTableNameList))
        Call cbTableName_AfterUpdate
    End If
    '
    ' Right click Copy / Paste
    '
    Set cBar = New clsBar
    cBar.Initialize Me
    '
    ' Center form on ActiveWindow
    '
    Me.top = Application.top + Application.Height / 2 - Me.Height / 2
    Me.left = Application.left + Application.width / 2 - Me.width / 2
End Sub

