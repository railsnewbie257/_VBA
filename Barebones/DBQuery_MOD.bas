Attribute VB_Name = "DBQuery_MOD"
Option Explicit

Sub Query(Optional useQuery, Optional whichQuery)
Attribute Query.VB_ProcData.VB_Invoke_Func = "q\n14"
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset
Dim SHDownload As String
Dim WBDownload As String
Dim fieldCount As Integer
Dim dbnamePrefix As String ' the single letter alias table name
Dim i As Long, j As Long
Dim useRow As Long, useCol As Long
Dim topRow As Long
Dim sheetType As String
Dim sheetName As String
Dim t As Integer
Dim v As Variant

1    DBGlbRecordsFound = -1
2    DBGlbRecordsToRead = -1
3    DBGlbAdodbError = False
    '
    ' Show QUERY form ----------------------------------------------------------------------------
    '
4    Call StatusbarDisplay("DBQuery: Start")
5    If IsMissing(useQuery) Then

6        QueryForm.Show

7        If formCancel Then Exit Sub
8        Call StartTimer
9    Else
10        GLBUserQuery = useQuery
11    End If
12    Call StatusbarDisplay("DBQuery: Download Header.")
13    t = GLBDownloadByColumn
14    Set DBCn = DBCheckConnection(DBCn)
    If DBCn Is Nothing Then Exit Sub
    
    Set DBRs = DBCheckRecordset(DBRs)
    Call StatusbarDisplay("DBQuery: Download Header.")
    '
    ' Get record count
    '
    Call StatusbarDisplay("DBQuery: Setup Recordset.")
    With DBRs
        .CursorLocation = adUseClient ' adUseServer, adUseClient
        .CursorType = adUseClient ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
        .LockType = adLockOptimistic ' adLockReadOnly
        Set .ActiveConnection = DBCn
    End With

    On Error GoTo gotError
        
        Debug_Print GLBUserQuery
        Set DBCn = DBCheckConnection(DBCn)
        If DBCn Is Nothing Then Exit Sub
        Set DBRs = DBCheckRecordset(DBRs)
        Call StatusbarDisplay("DBQuery: Submitting Query.")
100     DBRs.Open GLBUserQuery
        If DBGlbAdodbError Then
            DBRs.Close
            Set DBRs = Nothing
            Call CalculationOn
            Exit Sub
        End If
        '
        ' report records found
        '
110        Call StatusbarDisplay("DBQuery: Prepare Download.")
        DBGlbRecordsFound = -1
        If DBRs.State = adStateOpen Then DBGlbRecordsFound = DBRs.recordCount
        If DBGlbRecordsFound > 0 Then
            Call StatusbarDisplay("DBQuery: Records To Download")

                                 '===========================================================================
            ReadRecordsForm.Show '===========================================================================
                                 '===========================================================================
        Else
            MsgBox "No records found."
            formCancel = True
            Exit Sub
        End If
        If formCancel Then Exit Sub
    
200    If GLBDownloadByColumn Then
       End If
    If DBGlbRecordsToRead = 0 Then
        DBRs.Close
        Set DBRs = Nothing
        Call CalculationOn
        Exit Sub
    End If
    '
    ' DETERMINE TARGET SHEET --------------------------------------------------------------------------------
    '
    ' For downloading make sure sheet is empty and not in Macro workbook
    '
    If ActiveWorkbook.Name = MACROWORKBOOK Or GLBNewWorkbook Then
        If GLBQueryName = "ColumnNames" Then Call SetTableNames
        Workbooks.Add
        GLBNewWorkbook = False ' set default
    End If
    '
300    If WorksheetFunction.CountA(Cells) <> 0 And Not GLBSameSheet And Not GLBManualPlacement Then
        Sheets.Add
        GLBSameSheet = False ' set default
    End If
    
    If GLBQueryName = "ColumnNames" Then
        GLBColumnNamesWB = ActiveWorkbook.Name
        GLBColumnNamesSH = "ColumnNames"
    End If
    
    If GLBManualPlacement Then '------------------------------------------------------------------------------------------------------
        WBDownload = GLBDownloadWB
        If SHDownload = GLBDownloadSH Then GLBSameSheet = True
        SHDownload = GLBDownloadSH
        useCol = GLBPlacementColumn
        useRow = GLBPlacementRow
    Else
        WBDownload = ActiveWorkbook.Name
        SHDownload = ActiveSheet.Name
        '
        If GLBDownloadByColumn Or GLBQueryName = "ColumnNames" Then
            useCol = NextColumn(SHDownload, WBDownload)  ' This should be the upper left corner
            useRow = ColumnNextRow(useCol, SHDownload, WBDownload)
        Else
            useRow = LastRow(SHDownload, WBDownload) ' This should be the upper left corner
            If useRow = 0 Then
                useRow = 1
            Else
                useRow = useRow + 2
            End If
            useCol = RowNextColumn(useRow, SHDownload, WBDownload)
        End If
    End If
    
    
400    DBRs.Close
    Set DBRs = DBCheckRecordset(DBRs)
    '
    ' Reopen Recordset for download
    Call StatusbarDisplay("DBQuery: Submit Query.")
    Debug_Print GLBUserQuery
    DBRs.Open GLBUserQuery, DBCn
    '
    '
    ' Download Column Headers ------------------------------------------------------------------------
    '
    '
    Call StatusbarDisplay("DBQuery: Download Header.")
    Call ProgressMeterShow(0, DBGlbRecordsToRead)
    Call CalculationOff
    '
    ' Insert the TableName
    If GLBDownloadShowTableName Then
    
        With Workbooks(WBDownload).Worksheets(SHDownload).Cells(useRow, useCol)
            If GLBQueryName = "ColumnNames" Or GLBQueryName = "TableMap" Then
                .Value = GLBDatabaseName & "." & GLBTableName
            Else
                .Value = TableFromQuery(GLBUserQuery) 'GLBTableName
            End If
            .Font.Bold = True
            .Font.color = BLUE
            .Interior.color = ORANGE
            .Offset(0, 1) = Replace(GLBUserQuery, vbNewLine, "||")
            .Offset(0, 1).Interior.color = LIGHTGREEN
            .Offset(0, 2) = "<<< QUERY"
            .Offset(0, 2).Interior.color = LIGHTGREEN
            .Offset(0, 2).Font.Bold = True
            .Offset(0, 2).Font.color = RED
        End With
        useRow = useRow + 1
    End If
    '
    fieldCount = DBRs.Fields.count
    If DBGlbDownloadHeader Then
        '
        '
        If Not GLBDownloadByColumn Then  ' Direction By Row --------------------------------------------------------------------------
            
            fieldCount = DBRs.Fields.count
            For i = 0 To fieldCount - 1
                'If Not GLBQueryName = "ColumnNames" And Not GLBQueryName = "QueryForm" Then
                    Workbooks(WBDownload).Worksheets(SHDownload).Cells(useRow, useCol + i) = DBRs.Fields(i).Name
                    Workbooks(WBDownload).Worksheets(SHDownload).Cells(useRow, useCol + i).Font.Bold = True
                'End If
            Next i
            '
            ' ColumnNames -----------------------------------------------------------------------------------------
            '
            If GLBQueryName = "ColumnNames" Then
                dbnamePrefix = DBTableNextAlias
                Workbooks(WBDownload).Worksheets(SHDownload).Cells(1, useCol) = GLBDatabaseName & "." & GLBTableName & " " & dbnamePrefix
                Workbooks(WBDownload).Worksheets(SHDownload).Cells(1, useCol).Interior.color = ORANGE
                Workbooks(WBDownload).Worksheets(SHDownload).Cells(1, useCol).Font.color = BLUE
                Workbooks(WBDownload).Worksheets(SHDownload).Cells(1, useCol).Font.Bold = True
                'Call ColumnWidthAutoMax(useCol)
                'Workbooks(WBDownload).Worksheets(SHDownload).Columns(useCol).AutoFit
            End If

        Else ' Direction By Column ----------------------------------------------------------------------------
        '
        '
            If GLBQueryName = "ColumnNames" Then
                Workbooks(WBDownload).Worksheets(SHDownload).Cells(1, useCol) = GLBDatabaseName & "." & GLBTableName & " " & dbnamePrefix

                Workbooks(WBDownload).Worksheets(SHDownload).Cells(1, useCol).Interior.color = ORANGE
                Workbooks(WBDownload).Worksheets(SHDownload).Cells(1, useCol).Font.color = BLUE
                Workbooks(WBDownload).Worksheets(SHDownload).Cells(1, useCol).Font.Bold = True
                useRow = 2
            End If
            
            If GLBQueryName = "ColumnNames" Or IdentifySheet() = "ColumnNames" Then
                dbnamePrefix = dbnamePrefix & "."
                useRow = 2
            End If
            '
            '
            For i = 0 To fieldCount - 1
                Workbooks(WBDownload).Worksheets(SHDownload).Cells(i + useRow, useCol) = DBRs.Fields(i).Name
                Workbooks(WBDownload).Worksheets(SHDownload).Cells(i + useRow, useCol).Font.Bold = True
            Next i
        End If
    End If  ''If DBGlbDownloadHeader Then
    'If Not IsEmpty(useCol) Then Call ColumnWidthAutoMax(useCol:=useCol)
    '
    '
    ' Download Data ------------------------------------------------------------------------------------
    '
    '
    Call StatusbarDisplay("DBQuery: Download Data")
    
    '' MsgBox "Please wait for" & vbNewLine & vbNewLine & "DOWNLOAD FINISHED", Title:="DBQuery Download"
    
    If Not GLBDownloadByColumn Then ' Download By Row  ---------------------------------------------------------------------
        If GLBManualPlacement Then
            If DBGlbDownloadHeader Then useRow = useRow + 1
        ElseIf GLBSameSheet Then
            useRow = ColumnNextRow(useCol, SHDownload, WBDownload)
        Else
            useRow = ColumnNextRow(useCol, SHDownload, WBDownload)
        End If
        For i = 0 To DBGlbRecordsToRead - 1
            For j = 0 To fieldCount - 1
                Workbooks(WBDownload).Worksheets(SHDownload).Cells(useRow, useCol + j) = Trim(DBRs.Fields(j))
            Next j
            DBRs.MoveNext
            useRow = useRow + 1
            If useRow Mod 100 = 0 Then
                Call ProgressMeterShow(useRow, DBGlbRecordsToRead)
                If formCancel Then
                    Call ProgressMeterClose
                    Exit Sub
                End If
            End If
        Next i
    Else ' by Column -----------------------------------------------------------------------------------
        If GLBManualPlacement Then
            useCol = useCol + 1
        Else
            useCol = RowNextColumn(useRow, SHDownload, WBDownload)
        End If
        For i = 1 To DBGlbRecordsToRead
            For j = 0 To fieldCount - 1
                Workbooks(WBDownload).Worksheets(SHDownload).Cells(j + useRow, useCol) = DBRs.Fields(j).Value
                Workbooks(WBDownload).Worksheets(SHDownload).Cells(j + useRow, useCol).HorizontalAlignment = xlLeft
            Next j
            Workbooks(WBDownload).Worksheets(SHDownload).Columns(useCol).HorizontalAlignment = xlLeft
            Workbooks(WBDownload).Worksheets(SHDownload).Columns(useCol).AutoFit
            DBRs.MoveNext
            useCol = useCol + 1
            If useCol Mod 100 = 0 Then
                Call ProgressMeterShow(useCol, DBGlbRecordsToRead)
            End If
        Next i
    End If
    Call DBCloseRecordset(DBRs)
    Set DBRs = Nothing
    '
    ' Downloading FINISHED --------------------------------------------------------------------------------------------------
    '
    'Workbooks(WBDownload).Worksheets(SHDownload).Columns.AutoFit
    
    Call ColumnWidthAutoMax(SHDownload)

    If GLBQueryName <> "" Then
        Workbooks(WBDownload).Worksheets(SHDownload).Name = GLBQueryName
        SHDownload = GLBQueryName
    End If

    sheetType = IdentifySheet()
    '
    ' Callbacks Here ------------------------------------------------------------------------
    '
    'If DBGlbUseCallback <> "" Then
    '    Application.Run DBGlbUseCallback
    'End If
    If (sheetType = "MYTABLES") Then Call MyTablesCallback
    If (sheetType = "BRUCE") Then Call BruceCallback
    If (sheetType = "FASTLOAD") Then Call Fastload_Callback
    If (sheetType = "SHOWTABLE") Then Call ShowTable_Callback
    If (sheetType = "METERKEEPS") Then Call MeterKeepsCallback
    If (GLBQueryName = "UseQuerySheet") Then
        sheetName = InputBox("Sheet Tab Name:")
        If sheetName = "" Then sheetName = GLBQueryName
        ActiveSheet.Name = sheetName
        SHDownload = sheetName
    End If
    If (GLBQueryName = "UsageTracker") Then
        Call SortSheetDown(ActiveCell.CurrentRegion.Column + 2)
        If sheetName = "" Then sheetName = GLBQueryName
        ActiveSheet.Name = sheetName
        SHDownload = sheetName
    End If
    Workbooks(WBDownload).Worksheets(SHDownload).Activate
    If (GLBQueryName = "Drilldown-Event") Then
        useCol = FindColumnHeader("Event_Start_Tm")
        If useCol <> -1 Then Call SortSheetUp(useCol)
    End If
    
    Call CalculationOn
    Workbooks(WBDownload).Worksheets(SHDownload).Activate
    Call StatusbarDisplay("DBQuery: Finished.")
    '
    ' Reset Defaults
    '
    GLBDownloadByColumn = False
    GLBDownloadShowTableName = False
    GLBManualPlacement = False
    GLBQueryName = ""
    If GLBQueryName = "ColumnNames" Then
        GLBColumnNamesWB = ActiveWorkbook.Name
        GLBColumnNamesSH = ActiveSheet.Name
    End If
    GLBDownloadWB = WBDownload
    GLBDownloadSH = SHDownload
    '
    Call ProgressMeterClose
    '
    Exit Sub
    
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="DBQuery ERROR"
    DBGlbAdodbError = True
    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    '
    ' Reset
    '
    Call CalculationOn
    'Stop
    Resume Next
End Sub  ' Sub Query

Sub QueryToClipboard()

    GLBUserQuery = QueryBuilder(ActiveSheet.Name)
    
    Call CopyToClipboard(GLBUserQuery)
End Sub


