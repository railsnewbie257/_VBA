Attribute VB_Name = "DBQueryBuilder_MOD"
Sub AnalystQueryBuilder()
Attribute AnalystQueryBuilder.VB_ProcData.VB_Invoke_Func = "a\n14"

    QueryBuilderForm.Show False
    
End Sub

Function QueryBuilder(Optional SHQuery, Optional WBQuery)
Dim hasGroupBy As Boolean
Dim twoColumn As Boolean

    If IsMissing(SHQuery) Then SHQuery = ActiveSheet.Name
    If IsMissing(WBQuery) Then WBQuery = ActiveWorkbook.Name

    Call StatusbarDisplay("QueryBuilder: Start")
    'Workbooks(WBQuery).Worksheets(SHQuery).Activate
    
    botRow = LastRow(SHQuery, WBQuery)
    
    With Workbooks(WBQuery).Worksheets(SHQuery)
        
        GLBUserQuery = ""
        fieldCount = 0
        groupBy = ""
        getFields = True
        inSelect = False
        hasGroupBy = False
        twoColumn = False
        
        For i = 1 To botRow
            token1 = Trim(.Cells(i, 1))
            token1 = Replace(token1, Chr(160), " ")
            If Not token1 = "" Then
                If left(token1, 2) = "--" Then GoTo continue

                k = InStr(token1, "--")
                If k > 0 Then token1 = left(token1, k - 2)
                
                GLBUserQuery = GLBUserQuery & token1 & " "

                If UCase(Trim(token1)) = "SELECT" Then
                    inSelect = True
                Else
                    inSelect = False
                End If
                If UCase(Trim(token1)) = "GROUP BY" Then
                    inGroupBy = True
                    hasGroupBy = True
                Else
                    inGroupBy = False
                End If
            End If
            
            token2 = Trim(.Cells(i, 2))
            If Not token2 = "" Then
            
                twoColumn = True
                If Not IsEmpty(.Cells(i + 1, 2)) Then ' one line look ahead
                    c = "," & vbNewLine
                Else
                    c = " " & vbNewLine
                   End If
                GLBUserQuery = GLBUserQuery & token2 & c
                If inSelect Then  ' only group fields in the SELECT
                    '
                    ' Exclude aggregates
                    '
                    If UCase(left(token2, 5)) = "COUNT" Or _
                        UCase(left(token2, 3)) = "MIN" Or _
                        UCase(left(token2, 3)) = "MAX" Or _
                        UCase(left(token2, 3)) = "SUM" Then
                        
                        fieldCount = fieldCount + 1
                    Else
                        fieldCount = fieldCount + 1
                        If Len(groupBy) > 0 Then groupBy = groupBy & ","
                        groupBy = groupBy & format(fieldCount, "0")
                    End If
                End If
            End If
continue:
        Next i
        
    End With 'Workbooks(WBQuery).Worksheets(SHQuery)
    
    If Not hasGroupBy And twoColumn Then GLBUserQuery = GLBUserQuery & " GROUP BY " & groupBy
    Debug_Print GLBUserQuery
    
    GLBUserQuery = Replace(GLBUserQuery, ",,", ",")
    
    QueryBuilder = GLBUserQuery
    
    Call StatusbarDisplay("QueryBuilder: Done")
    'Workbooks(WBQuery).Worksheets(SHQuery).Activate
    
End Function

'
' Format:
' Column 1 = operator
' Column 2 = field
'
Function QueryBuilder_old(Optional SHQuery)

    If IsMissing(SHQuery) Then
        WBQuery = ActiveWorkbook.Name
        SHQuery = ActiveSheet.Name
    Else
        WBQuery = MACROWORKBOOK
    End If
    
    'Workbooks(WBQuery).Worksheets(SHQuery).Activate
    
    botRow = LastRow(SHQuery, WBQuery)
    
    With Workbooks(WBQuery).Worksheets(SHQuery)
        GLBUserQuery = ""
        fieldCount = 0
        groupBy = ""
        getFields = True
        inSelect = False
        For i = 1 To botRow
            If Not .Cells(i, 1) = "" Then
                If left(.Cells(i, 1), 2) = "--" Then GoTo continue

                GLBUserQuery = GLBUserQuery & .Cells(i, 1) & " "
                If UCase(Trim(.Cells(i, 1))) = "SELECT" Then
                    inSelect = True
                Else
                    inSelect = False
                End If
                If UCase(Trim(.Cells(i, 1))) = "GROUP BY" Then
                    inGroupBy = True
                Else
                    inGroupBy = False
                End If
            End If
            If Not .Cells(i, 2) = "" Then
                GLBUserQuery = GLBUserQuery & .Cells(i, 2) & " "
                If inSelect Then  ' only group fields in the SELECT
                    '
                    ' Exclude aggregates
                    '
                    If UCase(left(.Cells(i, 2), 5)) = "COUNT" Or _
                        UCase(left(.Cells(i, 2), 3)) = "MIN" Or _
                        UCase(left(.Cells(i, 2), 3)) = "MAX" Or _
                        UCase(left(.Cells(i, 2), 3)) = "SUM" Then
                        
                        fieldCount = fieldCount + 1
                    Else
                        fieldCount = fieldCount + 1
                           If Len(groupBy) > 0 Then groupBy = groupBy & ","
                     groupBy = groupBy & format(fieldCount, "0")
                    End If
                End If
            End If
continue:
        Next i
    End With 'Workbooks(WBQuery).Worksheets(SHQuery)
    
    GLBUserQuery = GLBUserQuery & " GROUP BY " & groupBy
    Debug_Print GLBUserQuery
    
    QueryBuilder = GLBUserQuery
    
    'Workbooks(WBQuery).Worksheets(SHQuery).Activate
    
End Function

Sub QueryNewCondition(SHQuery, WBQuery, useCondition, Optional newValue, Optional newCondition)
Dim sRange As range, fRange As range
Dim botRow As Long
Dim s As String
Dim dataCol As Integer

10  On Error GoTo gotError
20  dataCol = QUERYDATACOL
30  If IsColumnEmpty(dataCol, SHQuery, WBQuery) Then dataCol = 1  ' this is a cheat incase the query is only in the first column
40  botRow = ColumnLastRow(dataCol, SHQuery, WBQuery)
    
50  Set sRange = range(Workbooks(WBQuery).Worksheets(SHQuery).Cells(1, dataCol), _
                       Workbooks(WBQuery).Worksheets(SHQuery).Cells(botRow + 2, dataCol))
    
60  Set fRange = FindInRange(useCondition, sRange)
    
    If Not IsMissing(newCondition) Then
        s = newCondition
    Else
70      s = useCondition & " " & newValue
    End If
    
80  't = sRange.Cells(1, 1)
90  't = sRange.Cells(93, 1)
100 If Not fRange Is Nothing Then
110     fRange.Value = s
120 Else
130     Debug_Print useCondition & " not found"  ' warning message
140 End If
    
150 Exit Sub

gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl
    Stop
    Resume Next
    
End Sub
'
' quickly assemble the query on a spreadsheet
'
Sub QuickQuery()

    s = QueryBuilder(ActiveSheet.Name)
    
    Call CopyToClipboard(s)
    
End Sub
