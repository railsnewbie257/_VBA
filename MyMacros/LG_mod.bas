Attribute VB_Name = "lg_mod"
'
' Save all the Last Gasps, including dups
'
Sub LastGaspUpdate()
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

10  GLBUserQuery = "select max(rundate) from dl_oge_analytics." & TD_LASTGASP
    
20  Set DBDn = DBCheckConnection(DBCn)
30  Set DBRs = DBCheckRecordset(DBRs)
40  Set DBRs.ActiveConnection = DBCn
    
50  On Error GoTo gotError
60  DBRs.Open GLBUserQuery, DBCn

70  lastDate = format(DBRs.Fields(0).Value, "YYYY-MM-DD") ' the latest date in the database
    
80  If lastDate = format(Date - 1, "YYYY-MM-DD") Then
90      MsgBox "Last Gasp is Up-To-Date"
100     Exit Sub
110 End If
    
120 DBRs.Close

130 LOAD ChooseDateForm
    
140 ChooseDateForm.MonthView1.Value = lastDate
150 ChooseDateForm.Show
160 If formCancel Then Exit Sub
    
170 useDate = format(ChooseDateForm.MonthView1)
    Unload ChooseDateForm
    
180 SHUpdate = TD_LASTGASP & "Update"
190 Call QueryNewCondition(SHUpdate, MACROWORKBOOK, "RunDate =", "'" & format(useDate, "yyyy-mm-dd") & "'")
    ' Workbooks(MACROWORKBOOK).Worksheets(SHUpdate).Cells(98, 2) = "e.Event_Start_Dt = '" & useDate & "'"
    
200 GLBUserQuery = QueryBuilder(SHUpdate, MACROWORKBOOK)
    
    'debug_print GLBUserQuery
    
210 Set DBDn = DBCheckConnection(DBCn)
220 Set DBRs = DBCheckRecordset(DBRs)
230 If DBRs.ActiveConnection Is Nothing Then Set DBRs.ActiveConnection = DBCn
    
240 On Error GoTo gotError
250 DBRs.Open GLBUserQuery
    
260 MsgBox "Update Finished", Title:="LastGaspUpdate"
270  Exit Sub
    
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="LastGaspUpdate"
    Stop
    Resume Next

End Sub
'
' if the query changes, check QueryForm
'
Sub LastGaspDaily(Optional useDate)
Dim SHQuery As String

    On Error GoTo gotError
    
10    If IsMissing(useDate) Then
20        LOAD ChooseDateForm
30        ChooseDateForm.MonthView1.Value = format(Date - 1, "YYYY-MM-DD")
40        ChooseDateForm.Show
50        If formCancel Then
60            Unload ChooseDateForm
70            Exit Sub
80        End If
90        useDate = format(ChooseDateForm.MonthView1.Value, "YYYY-MM-DD")
100       Unload ChooseDateForm
110    End If
    
120    SHQuery = TD_LASTGASP & "Select"
130    Call QueryNewCondition(SHQuery, MACROWORKBOOK, "RunDate =", "'" & useDate & "'")
140    useQuery = QueryBuilder(SHQuery, MACROWORKBOOK)
    
150    GLBQueryName = "LastGasp"
160    Call Query(useQuery, False)
170    If formCancel Then Exit Sub
    
180    sortCol = FindColumnHeader("First_Event_Time_12007")
190    Call SortSheetUp(sortCol)
    
200    Exit Sub

gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="LastGaspDaily"
    Stop
    Resume Next
End Sub

Function VerifyDateFormat(s)
    b = StrConv(s, vbUnicode)
    c = Split(left(b, Len(b) - 1), vbNullChar)
    VerifyDateFormat = IsNumeric(c(0)) And IsNumeric(c(1)) And IsNumeric(c(2)) And _
        IsNumeric(c(3)) And (c(4) = "-") And IsNumeric(c(5)) And IsNumeric(c(6)) And _
        (c(7) = "-") And IsNumeric(c(8)) And IsNumeric(c(9))
End Function
