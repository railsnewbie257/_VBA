Attribute VB_Name = "DBCreateTable_MOD"
Sub DBCreateTableScript()

    WBUse = ActiveWorkbook.Name
    SHUse = ActiveSheet.Name
    SHQuery = "CreateTable"
    
    Call DeleteSheet(SHQuery)
    
    Sheets.Add
    ActiveSheet.Name = SHQuery
    
    Worksheets(SHQuery).Cells(1, 1) = "CREATE SET TABLE"
    GLBTableName = InputBox("Table Name?", Default:=GLBTableName, Title:="DBCreateTable")
    
    GLBDatabaseName = "dl_oge_analytics"
    Worksheets(SHQuery).Cells(2, 1) = GLBDatabaseName & "." & GLBTableName & ","
    Worksheets(SHQuery).Cells(3, 1) = "FALLBACK,"
    Worksheets(SHQuery).Cells(4, 1) = "NO BEFORE JOURNAL,"
    Worksheets(SHQuery).Cells(5, 1) = "NO AFTER JOURNAL,"
    Worksheets(SHQuery).Cells(6, 1) = "CHECKSUM = DEFAULT,"
    Worksheets(SHQuery).Cells(7, 1) = "DEFAULT MERGEBLOCKRATIO"
    Worksheets(SHQuery).Cells(8, 1) = "("
    
    lastColData = LastColumn(SHUse, WBUse)
    '
    ' replace any "."s since it's illegal
    '
    Worksheets(SHQuery).Cells(9, 1) = "_fl_id varchar(20) CHARACTER SET LATIN NOT CASESPECIFIC,"
    For i = 1 To lastColData
        t = TrimReplace(Workbooks(WBUse).Worksheets(SHUse).Cells(1, i))
        t = Replace(t, ".", "")
        t = Replace(t, "%", "pct")
        t = CheckReservedWord(UCase(t))
        
        If t <> "" Then
            aline = t & " varchar(20) CHARACTER SET LATIN NOT CASESPECIFIC"
            If i < lastColData Then aline = aline & ","
            Worksheets(SHQuery).Cells(i + 9, 1) = aline
        End If
    Next i
    Worksheets(SHQuery).Cells(i + 9, 1) = ")"
End Sub
