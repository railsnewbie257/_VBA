Attribute VB_Name = "TableUtils_MOD"
Sub Table2Plain()
    useName = ActiveSheet.Name
    Cells.Copy
    newName = useName & "-Plain"
    Set addSheet = Sheets.Add
    addSheet.Name = newName
    Cells(1, 1).PasteSpecial xlPasteValues
    Call ClearClipboard
End Sub

Sub TableToSheet()
    ActiveSheet.ListObjects(1).Range.Copy
    Sheets.Add
    Cells(1, 1).PasteSpecial Paste:=xlValues
End Sub

Sub CountTables()
    Debug.Print ActiveSheet.ListObjects(1).Name
End Sub

Sub SetTableNames()
    GLBTableNameWorkbook = ActiveWorkbook.Name
    GLBTableNameSheet = ActiveSheet.Name
End Sub

Sub GotoTableNames()
    If Not GLBTableNameWorkbook = "" And Not GLBTableNameSheet = "" Then
        Workbooks(GLBTableNameWorkbook).Worksheets(GLBTableNameSheet).Activate
    End If
End Sub

Sub showDatabaseList()
    t = ""
    If IsArrayAllocated(GLBDatabaseNameList) Then
        For i = UBound(GLBDatabaseNameList) To 1 Step -1
            t = t & GLBDatabaseNameList(i) & vbNewLine
        Next i
    End If
    t = t & "----------------------------------" & vbNewLine
    If IsArrayAllocated(GLBTableNameList) Then
        For i = UBound(GLBTableNameList) To 1 Step -1
            t = t & GLBTableNameList(i) & vbNewLine
        Next i
    End If
    t = t & "----------------------------------"
    MsgBox t
End Sub

Sub PreloadNameLists()
    Call DatabaseNameListAdd("dl_oge_analytics")
    Call DatabaseNameListAdd("putlvw")
    Call DatabaseNameListAdd("da_customer_vw")
    Call DatabaseNameListAdd("dbc")
    Call DatabaseNameListAdd("putl_cert_data_mart_views")

    Call TableNameListAdd("EUL_METER_READ_ACTG_PERIOD_SUM")
    Call TableNameListAdd("Usage_Drop_2")
    Call TableNameListAdd("Last_Gasp_2")
    Call TableNameListAdd("Zero_KWH_2")
    Call TableNameListAdd("markie_revenue")
    Call TableNameListAdd("eul_pos_meters_d")
    Call TableNameListAdd("Event")
    Call TableNameListAdd("columns")
    Call TableNameListAdd("billing_statement_charge")
End Sub

Sub AddDatabaseTable()
    DatabaseTableNames.Show
End Sub
Sub testDatabaseNameListAdd()
    Call DatabaseNameListAdd("hello")
    Call DatabaseNameListAdd("goodbye")
    Call DatabaseNameListAdd("hello")
    t = GLBDatabaseNameList(2)
    
    QueryForm.Show
    t = GLBDatabaseNameList(2)
End Sub

Sub DatabaseNameListAdd(s)
    found = False
    If IsArrayAllocated(GLBDatabaseNameList) Then
        For i = 1 To UBound(GLBDatabaseNameList)
            If GLBDatabaseNameList(i) = s Then
                ' GLBDatabaseNameList(i) = GLBDatabaseNameList(i + 1)
                found = True
            End If
        Next i
        If Not found Then ReDim Preserve GLBDatabaseNameList(UBound(GLBDatabaseNameList) + 1)
        GLBDatabaseNameList(UBound(GLBDatabaseNameList)) = s
    Else
        ReDim Preserve GLBDatabaseNameList(1)
        GLBDatabaseNameList(1) = s
    End If
    
    GLBDatabaseNameList(UBound(GLBDatabaseNameList)) = s
End Sub

Sub TableNameListAdd(s)
    found = False
    If IsArrayAllocated(GLBTableNameList) Then
        For i = 1 To UBound(GLBTableNameList)
            If GLBTableNameList(i) = s Then
                'GLBTableNameList(i) = GLBTableNameList(i + 1)
                found = True
            End If
        Next i
        If Not found Then ReDim Preserve GLBTableNameList(UBound(GLBTableNameList) + 1)
        GLBTableNameList(UBound(GLBTableNameList)) = s
    Else
        ReDim Preserve GLBTableNameList(1)
        GLBTableNameList(1) = s
    End If
    
    GLBTableNameList(UBound(GLBTableNameList)) = s
End Sub

Function TableFromQuery(Optional q)
Dim k As Integer, k2 As Integer
Dim s As String

    s = Replace(q, vbNewLine, " ") & "  "
    k = InStr(1, UCase(s), "FROM ")
    If k > 0 Then
        k2 = InStr(k + 5, s, " ")
        If k2 = 0 Then k2 = InStr(k + 5, s, vbNewLine)
        TableFromQuery = Mid(s, k + 5, k2 - (k + 5))
    End If
End Function

Function DatabaseNameFromQuery(Optional q)
Dim k As Integer, k2 As Integer

    k = InStr(1, q, "FROM ")
    If k > 0 Then
        k2 = InStr(k + 5, q, ".")
        If k2 > 0 Then
            DatabaseNameFromQuery = Mid(q, k + 5, k2 - (k + 5))
        End If
    End If
End Function

Function TableNameFromQuery(Optional q)
Dim k As Integer: k = 1
Dim k2 As Integer
Dim s As String

    s = Replace(q, vbNewLine, " ")
    k = InStr(1, s, "FROM ")
    If k > 0 Then
        k = InStr(k + 5, s, ".")
        k2 = InStr(k, s, vbNewLine)
        If k2 = 0 Then k2 = InStr(k, s, " ")
        If k2 > 0 Then
            TableNameFromQuery = Mid(s, k + 1, k2 - (k + 1))
        End If
    End If
End Function
