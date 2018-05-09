Attribute VB_Name = "DBInsert_MOD"
'
' Use DBCreateTable to make table creation script
'
Sub DBInsert()
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

    If SheetExists("CreateTable") Then
        useTable = Worksheets("CreateTable").Cells(2, 1)
        useTable = Replace(useTable, ",", "")
    Else
        useTable = InputBox("Table Name?", Title:="DBInsert")
    End If
    
    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    Set DBRs.ActiveConnection = DBCn
    
    botRow = LastRow()
    rightCol = LastColumn()
    botRow = 1000
    For i = 2 To botRow
    t = ""
        For j = 1 To rightCol
            t = t & "'" & Trim(Cells(i, j)) & "'"
            If j <> rightCol Then t = t & ","
        Next j
        
        s = "INSERT INTO " & useTable & " VALUES (" & t & ")"
        
        On Error GoTo GotErr
            DBRs.Open s
            
        If i Mod 100 = 0 Then StatusbarDisplay (i)
    Next i
Debug_Print Now()
    Exit Sub
    
GotErr:
    Debug_Print i & "    Error: ~" & Err.Description & "~"
    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    Set DBRs.ActiveConnection = DBCn
    Resume Next
End Sub

