Attribute VB_Name = "TableMap_MOD"
Sub SplitFullTableName(s, ByRef DBName As String, ByRef TBName As String)
Dim t As String
Dim k As Integer
    
    k = InStr(1, s, ".")
    DBName = Trim(Mid(s, 1, k - 1))
    TBName = Trim(Mid(s, k + 1, Len(s) - k))
    k = InStr(1, TBName, " ") ' remove alias
    If k > 0 Then TBName = Mid(TBName, 1, k - 1)
End Sub

Sub LoadTableMap()
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset
Dim tablenameRange As Range
Dim fieldRange As Range
Dim DatabaseName As String, tableName As String

    On Error Resume Next
    Set tablenameRange = Nothing
    Set tablenameRange = Application.InputBox("Select FullTable Name", Title:="TableMap", Type:=8)
    If tablenameRange Is Nothing Then Exit Sub
    
    Call SplitFullTableName(tablenameRange.Text, DatabaseName, tableName)
    
    Set fieldRange = Nothing
    Set fieldRange = Application.InputBox("Select Fields", Title:="TableMap", Type:=8)
    If fieldRange Is Nothing Then Exit Sub
    
    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    With DBRs
        .CursorLocation = adUseClient  ' adUseServer
        .CursorType = adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
        .LockType = adLockOptimistic ' adLockReadOnly
        Set .ActiveConnection = DBCn
    End With
    
    On Error GoTo gotError
    
    'DBRs.Open t, DBCn
    
    For i = 1 To fieldRange.Columns.count
        For j = 1 To fieldRange.Rows.count
            t = "INSERT INTO dl_oge_analytics.TableMap (DatabaseName, TableName, PrimaryTableName, FieldName) VALUES ("
            t2 = "'" & DatabaseName & "','" & tableName & "','" & DatabaseName & "." & tableName & "','" & fieldRange.Cells(j, i) & "')"
            
            t = t & t2
            Debug.Print t
            On Error GoTo gotError
    
            DBRs.Open t, DBCn
            
            ' do insert here
        Next j
    Next i
    
    Exit Sub
    
gotError:
    MsgBox "DBQuery Error (" & Err.Number & "): " & Err.Description, vbOKOnly, Title:="LoadTableMap"
    DBGlbAdodbError = True
    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    Resume Next
    
End Sub

'INSERT INTO Customers (CustomerName, City, Country)
'VALUES ('Cardinal', 'Stavanger', 'Norway');
Sub UpdateTableMap()
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset
Dim tablenameRange As Range, rowRange As Range
Dim DatabaseName As String, tableName As String
Dim fieldRange As Range, valueRange As Range


    On Error Resume Next
    Set tablenameRange = Nothing
    Set tablenameRange = Application.InputBox("Select FullTableName", Default:=Selection.Address, Title:="TableMapUpdate", Type:=8)
    If tablenameRange Is Nothing Then Exit Sub
    
    'Set rowRange = Nothing
    'Set rowRange = Application.InputBox("Select Row to Update", Default:=Selection.Address, Title:="TableMapUpdate", Type:=8)
    'If rowRange Is Nothing Then Exit Sub


    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    With DBRs
        .CursorLocation = adUseClient  ' adUseServer
        .CursorType = adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
        .LockType = adLockOptimistic ' adLockReadOnly
        Set .ActiveConnection = DBCn
    End With
    '
    ' Update Customers
    ' SET ContactName = 'Alfred Schmidt', City= 'Frankfurt'
    ' WHERE CustomerID = 1;
    '
    On Error GoTo gotError
    
    Do While True
        On Error Resume Next
        Set rowRange = Nothing
        Set rowRange = Application.InputBox("Select Field Value Pair", Default:="", Title:="TableMapUpdate", Type:=8)
        If rowRange Is Nothing Then Exit Do
        
        Call RangeChooseTwo(rowRange, fieldRange, valueRange)
        
        t = "UPDATE dl_oge_analytics.TableMap SET AKA = '" & valueRange.Text & "' "
        t2 = "WHERE PrimaryTableName = '" & tablenameRange.Text & "' AND FieldName = '" & fieldRange.Text & "'"
        t = t & t2
        
        DBRs.Open t, DBCn
    Loop
    
    Exit Sub
    
    
    
        
    'DBRs.Open t, DBCn
    primaryCol = FindColumnHeader("PrimaryTableName")
    fieldCol = FindColumnHeader("FieldName")
    
    useCol = rowRange.Column
    useRow = rowRange.Row
        t = "UPDATE dl_oge_analytics.TableMap SET " & Cells(2, useCol) & " = '" & Cells(useRow, useCol) & "' "
        Debug.Print t
        t2 = "WHERE PrimaryTableName = '" & Cells(useRow, primaryCol) & "' AND FieldName = '" & Cells(useRow, fieldCol) & "'"
        t = t & t2
        Debug.Print t
        On Error GoTo gotError
    
        DBRs.Open t, DBCn

        Exit Sub
gotError:
    MsgBox "DBQuery Error (" & Err.Number & "): " & Err.Description, vbOKOnly, Title:="UpdateTableMap"
    DBGlbAdodbError = True
    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    Resume Next
    
End Sub

Sub CheckTableMap()
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset
Dim aRange As Range
Dim primaryTableName As String
Dim range1 As Range, range2 As Range
Dim DatabaseName As String, tableName As String

    On Error Resume Next
    Set aRange = Nothing
    Set aRange = Application.InputBox("Select PrimaryTableName", Default:=Selection.Address, Title:="CheckTableMap", Type:=8)
    If aRange Is Nothing Then Exit Sub
    
    If aRange.count = 1 Then
        primaryTableName = aRange.Text
    ElseIf Selection.count = 2 Then
        Call RangeChooseTwo(Selection, range1, range2)
        primaryTableName = range1 & "." & range2
    Else
        MsgBox "Can't Figureout Name " & Selection
        Exit Sub
    End If
        
    

    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    With DBRs
        .CursorLocation = adUseClient  ' adUseServer
        .CursorType = adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
        .LockType = adLockOptimistic ' adLockReadOnly
        Set .ActiveConnection = DBCn
    End With
    '
    ' Update Customers
    ' SET ContactName = 'Alfred Schmidt', City= 'Frankfurt'
    ' WHERE CustomerID = 1;
    '
    On Error GoTo CheckTableMapError
    
    t = "SELECT FieldName, AKA FROM dl_oge_analytics.TableMap WHERE PrimaryTableName = '" & primaryTableName & "'"
    
    DBRs.Open t, DBCn
    
    If DBRs.recordCount > 0 Then
        Set aRange = Nothing
        Set aRange = Application.InputBox("Select Download Location", Default:=Selection.Address, Title:="CheckTableMap", Type:=8)
        If aRange Is Nothing Then Exit Sub
        
        useRow = aRange.Row
        useCol = aRange.Column
        
        With Cells(useRow, useCol)
            .Value = primaryTableName
            .Interior.color = ORANGE
            .Font.color = BLUE
            .Font.Bold = True
        End With
    
         For i = 1 To DBRs.recordCount
            For j = 0 To DBRs.Fields.count - 1
                Cells(useRow + i, useCol + j) = DBRs.Fields(j)
                If j = 0 Then Cells(useRow + i, useCol + j).Font.Bold = True
            Next j
        
            DBRs.MoveNext
        Next i
    Else
        MsgBox "Not Found."
    End If

    Exit Sub
    
CheckTableMapError:
    MsgBox "DBQuery Error (" & Err.Number & "): " & Err.Description, vbOKOnly, Title:="CheckTableMap"
    DBGlbAdodbError = True
    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    Resume Next
    
End Sub

