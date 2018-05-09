VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DBQueryAnalystForm 
   Caption         =   "Teradata Query Analyst"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14985
   OleObjectBlob   =   "DBQueryAnalystForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DBQueryAnalystForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cBar As clsBar
Dim queryBaseJustSaved As Boolean

Sub test()
    txtQuery = "123" & vbNewLine & "456" & vbNewLine & "789"
    
End Sub

Private Function InsertIntoQuery(Optional s) As String
Dim start As Integer, length As Integer
Dim selectionLen As Integer
Dim queryLen As Integer

    start = Me.txtQuery.SelStart
    queryLen = Len(txtQuery)
    subLen = Me.txtQuery.SelLength
    newLines = Len(left(txtQuery, start)) - Len(Replace(left(txtQuery, start), Chr(13), ""))
    start = start + newLines  ' the real start with linebreaks
    
    If start >= queryLen Then
        InsertIntoQuery = txtQuery & s
    Else
        s1 = left(txtQuery, start)
        s2 = Right(txtQuery, Len(txtQuery) - (start + subLen + 1))
    
        InsertIntoQuery = s1 & s & s2
    End If
    
    Exit Function
    
    Debug.Print "s1>" & s1 & "<"
    Debug.Print "s2>" & s2 & "<"
        
    
    queryLen = Len(txtQuery)
    
    t = Mid(txtQuery, start, -1)
    Debug.Print "before start=" & Asc(t) & "/" & t  ' 13,10
    t = Mid(txtQuery, start, 1)
    Debug.Print "after start=" & Asc(t) & "/" & t
    t = Mid(txtQuery, start + len2 - 1, 1)
    Debug.Print "before end=" & Asc(t) & "/" & t  ' 13,10
    t = Mid(txtQuery, start + len2, 1)
    Debug.Print "after end=" & Asc(t) & "/" & t

    If t = Chr(13) Then start = start + 1  ' if inbetween CR/LF
    If t = Chr(10) Then start = start + 1
    'If t = vbLf Then start = start + 1  ' if inbetween CR/LF
    If start >= queryLen Then
        InsertIntoQuery = txtQuery & s
    Else
        'newLines = Len(left(txtQuery, start)) - Len(Replace(left(txtQuery, start), Chr(13), ""))
        'start = start + newLines
    
        't = Mid(txtQuery, start, 1)
        'If t = Chr(13) Then start = start + 1  ' if inbetween CR/LF
        
    
        't2 = txtQuery
        'txtQuery = t2
        Debug.Print Asc(t)
        s1 = left(txtQuery, start)
        Debug.Print "s1>" & s1 & "<"
        selectionLen = Me.txtQuery.SelLength
        If selectionLen > 0 Then
            s2 = Right(txtQuery, Len(txtQuery) - (start + selectionLen))
        Else
            s2 = Right(txtQuery, Len(txtQuery) - start)
        End If
        Debug.Print "s1>" & s1 & "<"
        Debug.Print "s2>" & s2 & "<"
        
        t = Len(s2)
        If Len(s2) = 0 Then s2 = vbNewLine
        InsertIntoQuery = s1 & s & s2
        Debug.Print "~" & InsertIntoQuery & "~"
        'If t = Mid(s2, Len(s2) - 2, 1) Then i = 1
    End If
    
End Function

Private Sub Tabstrip1_Handler()
    If TabStrip1.Value = 0 Then
        GLBQueryBaseWB = ActiveWorkbook.Name
        GLBQueryBaseSH = ActiveSheet.Name
        queryBaseJustSaved = True
        TabStrip1.Value = 1
    ElseIf TabStrip1.Value = 2 Then
        If GLBDownloadWB <> "" Then
            'DBQueryAnalystForm.Hide
            Workbooks(GLBDownloadWB).Worksheets(GLBDownloadSH).Activate
            Unload Me
        Else
            MsgBox "There are no downloads yet."
        End If
    ElseIf Not queryBaseJustSaved Then
        If GLBDownloadWB <> "" Then
            'DBQueryAnalystForm.Hide
            Workbooks(GLBQueryBaseWB).Worksheets(GLBQueryBaseSH).Activate
            'DBQueryAnalystForm.Show
            Unload Me
        End If
    Else
        queryBaseJustSaved = False
    End If
End Sub

Private Sub btnClear_Click()
    txtQuery = ""
    cbDatabaseName = ""
    cbTableName = ""
End Sub

Private Sub btnCount_Click()
Dim aRange As Range
    On Error Resume Next
    Set aRange = ActiveCell
    'Set aRange = Application.InputBox("Choose Field", Title:="Choose Distinct", Default:=ActiveCell.Address(False, False), Type:=8)
    If aRange Is Nothing Then Exit Sub
    
    t = "COUNT(" & aRange.Text & ")," & vbNewLine
    txtQuery = InsertIntoQuery(t)
End Sub

Private Sub btnDistinct_Click()
Dim aRange As Range
    On Error Resume Next
    Set aRange = ActiveCell
    'Set aRange = Application.InputBox("Choose Field", Title:="Choose Distinct", Default:=ActiveCell.Address(False, False), Type:=8)
    If aRange Is Nothing Then Exit Sub
    
    t = "DISTINCT(" & aRange.Text & ")," & vbNewLine
    txtQuery = InsertIntoQuery(t)
End Sub

Private Sub btnJoin_Click()
Dim aRange As Range, bRange As Range

    On Error Resume Next
    Set aRange = Nothing
    Set aRange = Application.InputBox("First Table To Join", Title:="Join", Type:=8)
    If aRange Is Nothing Then Exit Sub
    
    Set aKey = Nothing
    Set aKey = Application.InputBox("First Match Column", Title:="Join", Type:=8)
    If aKey Is Nothing Then Exit Sub
    
    Set bRange = Nothing
    Set bRange = Application.InputBox("Second Table To Join", Title:="Join", Type:=8)
    If bRange Is Nothing Then Exit Sub
    
    Set bKey = Nothing
    Set bKey = Application.InputBox("Second Match Column", Title:="Join", Type:=8)
    If bKey Is Nothing Then Exit Sub
    
    t = "JOIN " & aRange.Text & " ON " & aKey.Text & "=" & bKey
    
    txtQuery = InsertIntoQuery(t)
End Sub

Private Sub btnLocator_Click()
    LocatorForm.Show
    If formCancel Then
        formCancel = False
        Exit Sub
    End If
End Sub

Private Sub btnQueries_Click()
    SubQueryFormA.Show
End Sub

Private Sub btnResetConnection_Click()
    Call DBCloseConnection
End Sub

Private Sub btnTableMap_Click()
    t = "SELECT * FROM dl_oge_analytics.TableMap" & vbNewLine & "WHERE DatabaseName = '" & cbDatabaseName & "'" & vbNewLine
    t2 = "AND TableName = '" & cbTableName & "'"
    
    txtQuery = t & t2
    GLBQueryName = "TableMap"
    GLBDatabaseName = cbDatabaseName
    GLBTableName = cbTableName
End Sub

Private Sub btnTables_Click()
    t = "SELECT databasename, tablename, creatorname, lastaltertimestamp, commentstring" & vbNewLine & "FROM dbc.tables" & vbNewLine
    t = t & "WHERE tablekind = 'T'" & vbNewLine
    t = t & "AND databasename = 'dl_oge_analytics'" & vbNewLine
    t = t & "AND tablename like 'Last_Gasp%'" & vbNewLine
    t = t & "OR tablename like 'Usage_Drop%'" & vbNewLine
    t = t & "OR tablename like 'ZERO%'" & vbNewLine
    txtQuery = t
End Sub

Private Sub btnTest_Click()
Dim DataObj As MSForms.DataObject
    
    txtQuery = InsertIntoQuery("xyz")
    Exit Sub
    
    MsgBox Me.txtQuery.SelStart & ">" & Me.txtQuery.SelText & "<" & Me.txtQuery.SelLength
    Exit Sub
    
    Set DataObj = New MSForms.DataObject
    DataObj.GetFromClipboard

    txtQuery.SetFocus
    txtQuery = DataObj.GetText(1)

    'txtQuery.SetFocus
    'Sheets("Sheet2").Rows(1).PasteSpecial Transpose:=True
End Sub

Private Sub btnWhere_Click()
Dim aRange As Range
Dim range1 As Range, range2 As Range
Dim t As String

    Set aRange = Nothing
    'Selection.Copy
    On Error Resume Next
    DBQueryAnalystForm.Hide
    Set aRange = Application.InputBox("Select Field and Value:", Title:="DBQueryAnalyst", Default:=Selection.Address, Type:=8)
    DBQueryAnalystForm.Show False
    If aRange Is Nothing Then Exit Sub
    
    t = "WHERE "
    k = InStr(1, UCase(txtQuery), "WHERE")
    If (k > 0 And k < Me.txtQuery.SelStart) Then t = "AND "
    
    Call RangeChooseTwo(aRange, range1, range2)
    
    If IsDate(range2.Text) Then
        t = t & labelFillSpaces(range1.Text) & " = " & format(range2.Text, "'yyyy-mm-dd'") '& vbNewLine
    ElseIf IsNumeric(range2.Text) Then
        t = t & labelFillSpaces(range1.Text) & " = " & range2.Text '& vbNewLine
    Else
        t = t & labelFillSpaces(range1.Text) & " LIKE '%" & range2.Text & "%'" '& vbNewLine
    End If

    txtQuery = InsertIntoQuery(t)
End Sub

Private Sub TabStrip1_Change()
    Call Tabstrip1_Handler
End Sub

Private Sub TabStrip1_Click(ByVal Index As Long)
    Call Tabstrip1_Handler
End Sub

Private Sub btnBaseSheet_Click()
    If GLBQueryAmarkerWB = "" Then
        GLBQueryAmarkerWB = ActiveWorkbook.Name
        GLBQUeryAmarkerSH = ActiveSheet.Name
        btnBaseSheet.Font.Bold = True
    Else
        On Error GoTo NoSheet
        Workbooks(GLBQueryAmarkerWB).Worksheets(GLBQUeryAmarkerSH).Activate
        Unload Me
    End If
        
    btnBaseSheet.ControlTipText = GLBQUeryAmarkerSH
    Exit Sub
    
NoSheet:
    MsgBox GLBQueryAmarkerWB & " " & GLBQUeryAmarkerSH & " does not exist."
    GLBQueryAmarkerWB = ""
    GLBQUeryAmarkerSH = ""
End Sub

Private Sub btnCancel_Click()
    formCancel = True
    Unload Me
End Sub

Private Sub btnColumnNames_Click()
    txtQuery = "SELECT columnname FROM dbc.columnsv WHERE databasename='" & cbDatabaseName & "' and tablename='" & cbTableName & "' ORDER BY columnid;"
    GLBQueryName = "ColumnNames"
    GLBDatabaseName = txtDatabaseName
    GLBTableName = txtTableName
End Sub

Private Sub btnDatabaseTables_Click()
    t = "SELECT databasename, tablename, creatorname, lastaltertimestamp, commentstring FROM dbc.tables WHERE tablekind = 'T'" & vbNewLine
    If cbDatabaseName <> "" Then t = t & "AND databasename =" & "'" & cbDatabaseName & "'" & vbNewLine
    If cbTableName <> "" Then t = t & "AND tablename =" & "'" & cbTableName & "'" & vbNewLine
    t = t & "ORDER BY 1,2"
    txtQuery = t
    GLBQueryName = "DatabaseViews"
End Sub

Private Sub btnDatabaseViews_Click()
    t = "SELECT databasename, tablename, creatorname, lastaltertimestamp, commentstring FROM dbc.tables WHERE tablekind = 'V'" & vbNewLine
    If cbDatabaseName <> "" Then t = t & "AND databasename =" & "'" & cbDatabaseName & "'" & vbNewLine
    If cbTableName <> "" Then t = t & "AND tablename =" & "'" & cbTableName & "'" & vbNewLine
    t = t & "ORDER BY 1,2"
    txtQuery = t
    GLBQueryName = "DatabaseViews"
End Sub

Private Sub btnDefaultDatabase_Click()
    cbDatabaseName = "dl_oge_analytics"
End Sub

Private Sub btnFrom_Click()
    If cbDatabaseName = "" Then
        t = "FROM  place.holder" & vbNewLine
    Else
        t = "FROM " & cbDatabaseName & "." & cbTableName & vbNewLine
    End If
     
    txtQuery = InsertIntoQuery(t)
End Sub

Private Sub btnLoadTableName_Click()
Dim k As Long, k2 As Long
Dim i As Long, j As Long
Dim tDBName As String: tDBName = ""
Dim tTBName As String: tTBName = ""

    If Selection.count = 1 Then
        If InStr(1, Selection.Text, ".") Then Call SplitFullTableName(Selection.Text, tDBName, tTBName)
        cbDatabaseName = tDBName
        cbTableName = tTBName
    
    ElseIf Selection.Areas.count = 1 And Selection.Columns.count = 2 Then
        cbDatabaseName = Selection.Columns(1)
        cbTableName = Selection.Columns(2)
    
    ElseIf Selection.Areas.count = 2 Then
        cbDatabaseName = Selection.Areas(1).Cells(1, 1)
        cbTableName = Selection.Areas(2).Cells(1, 1)
    End If
    
End Sub

Private Sub btnLoadQuery_Click()
Dim aRange As Range
    Set aRange = Nothing
    On Error Resume Next
    Set aRange = Application.InputBox("Select Query To Load", Title:="DBQueryAnalyst", Default:=Selection.Address, Type:=8)
    If aRange Is Nothing Then Exit Sub
    
    txtQuery = Replace(aRange.Text, "||", vbNewLine)
    
End Sub

Private Sub btnPreviousQuery_Click()
    txtQuery = GLBUserQuery
End Sub

Private Sub btnSample_1_Click()
    t = "SELECT TOP 3 * FROM " & cbDatabaseName & "." & cbTableName & vbNewLine
    txtQuery = txtQuery & t
    GLBDownloadShowTableName = True
    GLBDownloadByColumn = True
End Sub

Private Sub btnSample_10_Click()
    t = "SELECT TOP 10 * FROM " & cbDatabaseName & "." & cbTableName & vbNewLine
    txtQuery = txtQuery & t
    GLBDownloadShowTableName = True
    GLBDownloadByColumn = True
End Sub

Private Sub btnSaveQuery_Click()
Dim aRange As Range
    Set aRange = Nothing
    On Error Resume Next
    Set aRange = Application.InputBox("Select Location To Save This Query", Title:="DBQueryAnalyst", Type:=8)
    If aRange Is Nothing Then Exit Sub
    
    'aRange = txtQuery
    aRange = Replace(txtQuery, vbNewLine, "||")
End Sub

Private Sub btnSelect_Click()
Dim retCode As Integer

    k = InStr(UCase(txtQuery), "SELECT *")
    If k > 0 Then
        txtQuery = Replace(txtQuery, "SELECT *", "SELECT")
    ElseIf InStr(UCase(txtQuery), "SELECT") > 0 Then
        txtQuery = Replace(txtQuery, "SELECT", "SELECT *")
    Else
        txtQuery = "SELECT" & vbNewLine & txtQuery
    End If
    
    txtQuery = InsertIntoQuery(t)
End Sub

Private Sub btnSelectFields_Click()
Dim i As Integer, j As Integer
Dim R As Integer, c As Integer
Dim t As String

    On Error Resume Next
    Set aRange = Nothing
    DBQueryAnalystForm.Hide
    
    Set aRange = Application.InputBox("Select Field(s)", Title:="Select Fields", Default:=Selection.Address, Type:=8)
    DBQueryAnalystForm.Show False
    If aRange Is Nothing Then Exit Sub
    
    t = ""
    count = aRange.count - 1
    For i = 1 To aRange.Areas.count
        For R = 1 To aRange.Areas(i).Rows.count
            For c = 1 To aRange.Areas(i).Columns.count
                t = t & labelFillSpaces(aRange.Areas(i).Cells(R, c))
                If count > 0 Then
                    t = t & "," & vbNewLine
                Else
                    t = t & vbNewLine
                End If
                count = count - 1
            Next c
        Next R
    Next i
    txtQuery = InsertIntoQuery(t)
End Sub

Private Sub btnSubmit_Click()
    formCancel = False
    GLBUserQuery = txtQuery
    If GLBQueryName = "ColumnNames" Then
        GLBDatabaseName = cbDatabaseName
        GLBTableName = cbTableName
    Else
        cbDatabaseName = DatabaseNameFromQuery(GLBUserQuery)
        cbTableName = TableNameFromQuery(GLBUserQuery)
    End If
    GLBQueryBaseWB = ActiveWorkbook.Name
    GLBQueryBaseSH = ActiveSheet.Name
    
    Unload Me
    
    Call Query(GLBUserQuery)
End Sub

Private Sub btnViewNames_Click()
    txtQuery = "SELECT databasename, tablename, creatorname, lastaltertimestamp, commentstring FROM dbc.tables WHERE tablekind = 'V' ORDER BY 1,2"
    'GLBQueryName = "DatabaseViews"
End Sub

Private Sub cbDatabaseName_Change()
Dim tDBName As String, tTBName As String
    
    k = InStr(1, cbDatabaseName, ".")
    If k > 0 Then
        Call SplitFullTableName(cbDatabaseName, tDBName, tTBName)
        cbDatabaseName = tDBName
        cbTableName = tTBName
    End If
    
    Exit Sub
    
    Select Case UCase(cbDatabaseName)
    
        Case "DA_CUSTOMER_VW"
            Call TableNameListAdd("billing_statement_charge")
            cbTableName = "billing_statement_charge"
            
        Case "PUTLVW"
            Call TableNameListAdd("billing_statement_charge")
            cbTableName = "billing_statement_charge"
            
        Case "DL_OGE_ANALYTICS"
            Call TableNameListAdd("Last_Gasp_2")
            cbTableName = "Last_Gasp_2"
    End Select
End Sub
Private Sub cbTableName_Change()
Dim tDBName As String, tTBName As String
    
    k = InStr(1, cbTableName, ".")
    If k > 0 Then
        Call SplitFullTableName(cbTableName, tDBName, tTBName)
        cbDatabaseName = tDBName
        cbTableName = tTBName
    End If

End Sub

Private Sub CommandButton2_Click()

End Sub

Private Sub opt_da_customer_vw_Click()
    cbDatabaseName = "da_customer_vw"
End Sub

Private Sub opt_dl_oge_analytics_Click()
    cbDatabaseName = "dl_oge_analytics"
End Sub

Private Sub opt_putlvw_Click()
    cbDatabaseName = "putlvw"
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    '
    ' Load Database and Table names
    '
    cbDatabaseName = GLBDatabaseName
    cbTableName = GLBTableName
    '
    ' Option Buttons
    '
    If GLBQueryBaseWB <> "" Then
        queryBaseJustSaved = True
        TabStrip1.Value = 1
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

Private Sub UserForm_Terminate()
    If DBQueryAnalystForm.Visible Then Unload Me
End Sub
