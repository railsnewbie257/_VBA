Attribute VB_Name = "Fastload_MOD"
'
' This function replaces invalid characters in table names
'
Function TrimReplace(t)
Dim s As String
    s = Trim(t)
    s = Replace(s, ".", "")
    s = Replace(s, " ", "_")
    s = Replace(s, "(", "_")
    s = Replace(s, ")", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, "-", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, ">", "_")
    s = Replace(s, "<", "_")
    s = Replace(s, "'", "")
    s = Replace(s, """", "")
    s = Replace(s, "%", "pct")
    t1 = Len(s)
    s = Replace(s, "__", "_")
    If Len(s) <> t1 Then s = Replace(s, "__", "_")
    t1 = Len(s)
    If Len(s) <> t1 Then s = Replace(s, "__", "_")
    t1 = Len(s)
    If Len(s) <> t1 Then s = Replace(s, "__", "_")
    TrimReplace = s
End Function

Sub FastLoad()
Dim wsh As Object
Dim userTableName As String     ' from user
Dim newTableName As String      ' table name after TrimReplace, user may have used invalid characters
Dim fullTableName As String     ' database table name

'Dim waitOnReturn As Boolean: waitOnReturn = True
'Dim windowStyle As Integer: windowStyle = 1 ' or whatever suits you best
Dim emptyColumnCount As Integer: emptyColumnCount = 1
Dim errorCode As Integer
Dim mergeTable As Boolean       ' whether a merge is necessary because the table already exists

On Error Resume Next
    MkDir "C:\Fastload"

On Error GoTo gotError
    
    'Set DBGlbConnection = Nothing

    Set wsh = CreateObject("WScript.Shell")

    WBOrig = ActiveWorkbook.Name
    SHOrig = ActiveSheet.Name
    Dim c As String
    
    On Error Resume Next
    userTableName = InputBox("Table Name?", Title:="FastLoad", Default:=GLBTableName)
    If IsEmpty(userTableName) Or userTableName = "" Then Exit Sub
    GLBTableName = userTableName
    newTableName = TrimReplace(userTableName)
    If newTableName <> userTableName Then
        retCode = MsgBox("Modifying Table Name:" & vbNewLine & vbNewLine & userTableName & "  ->  " & newTableName, vbOKCancel, Title:="FastLoad")
        If retCode = vbCancel Then Exit Sub
    End If
    
    Call UsageTracker("FastLoad", "Start: " & newTableName)
    '
    '
    DatabaseName = "dl_oge_analytics"
    fullTableName = DatabaseName & "." & newTableName
    Call UsageTracker("FastLoad", fullTableName)
    
    If TDTableExists(fullTableName) Then
        retCode = MsgBox("Table: " & newTableName & " already EXISTS, will APPEND this table.", vbOKCancel)
        If retCode = vbCancel Then Exit Sub
        mergeTable = True
        newTableName = newTableName & "_up"
        fullTableName = DatabaseName & "." & newTableName

    Else
        retCode = MsgBox("Creating Table: " & newTableName, vbOKCancel, Title:="Fastload")
        If retCode = vbCancel Then Exit Sub
        mergeTable = False
    End If
    '
    filePath = "C:\Fastload\" & newTableName & ".fl"
    Kill filePath
    
    Call StatusbarDisplay("Fastload: Setup")
    
    Call FastLoadWrite(filePath, "LOGMECH LDAP;")
    userName = LCase(Environ$("Username"))
    Call FastLoadWrite(filePath, "LOGON TD1/" & userName & "," & Password & ";")
    Call FastLoadWrite(filePath, "DATABASE dl_oge_analytics;")
    '
    ' DROP TABLES ------------------------------------------------------------------------------
    '
    Call FastLoadWrite(filePath, "DROP TABLE " & fullTableName & ";")
    Call FastLoadWrite(filePath, "DROP TABLE " & fullTableName & "_ET;")
    Call FastLoadWrite(filePath, "DROP TABLE " & fullTableName & "_UV;")
    '
    ' CREATE TABLE -----------------------------------------------------------------------------
    '
    Call StatusbarDisplay("Fastload: Create Table")
    Call FastLoadWrite(filePath, "")
    Call FastLoadWrite(filePath, "CREATE MULTISET TABLE " & fullTableName & ",")
    Call FastLoadWrite(filePath, "NO FALLBACK,")
    Call FastLoadWrite(filePath, "NO BEFORE JOURNAL,")
    Call FastLoadWrite(filePath, "NO AFTER JOURNAL,")
    Call FastLoadWrite(filePath, "CHECKSUM = DEFAULT,")
    Call FastLoadWrite(filePath, "DEFAULT MERGEBLOCKRATIO")
    Call FastLoadWrite(filePath, "(")
    rightCol = RowLastColumn(1, SHOrig, WBOrig)
    '
    ' COLUMN NAMES -----------------------------------------------------------------------------
    '
    Call FastLoadWrite(filePath, "LoadDate varchar(20),") ' load date column
    For i = 1 To rightCol
        t = Worksheets(SHOrig).Cells(1, i)
        If t = "" Then
            t = "EmptyColumn" & emptyColumnCount
            Worksheets(SHOrig).Cells(1, i) = t
            emptyColumnCount = emptyColumnCount + 1
        End If
        t = CheckReservedWord(t)
        c = ","
        If i = rightCol Then c = ")"
        Call FastLoadWrite(filePath, t & " varchar(300)" & c)
    Next i

    t = CheckReservedWord(Worksheets(SHOrig).Cells(1, 1))
    Call FastLoadWrite(filePath, "PRIMARY INDEX(" & t & ");") 'set first column as primary index to spread processing
    '
    '
    Call FastLoadWrite(filePath, "")
    Call FastLoadWrite(filePath, "BEGIN LOADING " & fullTableName)
    Call FastLoadWrite(filePath, "ERRORFILES " & newTableName & "_ET, " & newTableName & "_UV;")
    Call FastLoadWrite(filePath, "SET RECORD VARTEXT delimiter " & "'|' QUOTE YES " & "'" & """" & "'" & ";")
    '
    ' DEFINE -------------------------------------------------------------------------------------
    '
    Call StatusbarDisplay("Fastload: Define")
    Call FastLoadWrite(filePath, "")
    Call FastLoadWrite(filePath, "DEFINE")
    Call FastLoadWrite(filePath, "in_LoadDate (varchar(20)),")
    For i = 1 To rightCol
        t = Worksheets(SHOrig).Cells(1, i)
        t = "in_" & TrimReplace(t)
        c = ","
        If i = rightCol Then c = ""
        Call FastLoadWrite(filePath, t & " (varchar(300))" & c)
    Next i
    Call FastLoadWrite(filePath, "FILE= " & newTableName & ".txt;")
    '
    ' INSERT --------------------------------------------------------------------------------------
    '
    Call FastLoadWrite(filePath, "")
    Call FastLoadWrite(filePath, "INSERT INTO " & fullTableName & " (")
    Call FastLoadWrite(filePath, "LoadDate,")
    For i = 1 To rightCol
        t = Worksheets(SHOrig).Cells(1, i)
        t = CheckReservedWord(t)
        c = ","
        If i = rightCol Then c = ")"
        Call FastLoadWrite(filePath, t & c)
    Next i
    '
    ' VALUES --------------------------------------------------------------------------------------
    '
    Call FastLoadWrite(filePath, "")
    Call FastLoadWrite(filePath, "VALUES (")
    Call FastLoadWrite(filePath, ": in_LoadDate,")
    For i = 1 To rightCol
        t = Worksheets(SHOrig).Cells(1, i)
        t = ": in_" & TrimReplace(t)
        c = ","
        If i = rightCol Then c = ");"
        Call FastLoadWrite(filePath, t & c)
    Next i
    '
    ' END LOADING ---------------------------------------------------------------------------------
    '
    Call FastLoadWrite(filePath, "END LOADING;")
    Call FastLoadWrite(filePath, "LOGOFF;")
    '
    ' DATA FILE -----------------------------------------------------------------------------------
    '
    Call StatusbarDisplay("Fastload: Create Data File")
    filePath = "C:\Fastload\" & newTableName & ".txt"
    Kill filePath
    
    Set foundRange = FindRangeErrors(Cells)
    foundRange.Value = ""  ' clear all #N/As
    
    botRow = ColumnLastRow(1, SHOrig)
    For j = 2 To botRow
        aline = """" & format(Now(), "mm/dd/yyyy") & """" & "|"
        'aline = ""
        For i = 1 To rightCol
            If j = botRow + 10 Then
                ' get the format frm the first line of data, header may only be "General"
                aline = aline & """" & Worksheets(SHOrig).Cells(j + 1, i).NumberFormat & """"
            Else
                t = Worksheets(SHOrig).Cells(j, i).Value
                t = Replace(t, """", "")
                aline = aline & """" & t & """"
            End If
            c = "|"
            If i = rightCol Then c = ""
            aline = aline & c
        Next i
        Call FastLoadWrite(filePath, aline)
    Next j
    Call StatusbarDisplay("Fastload: Shell Run")
    '
    ' Shell DOS command ---------------------------------------------------------------------------
    '
    t = "cmd.exe /c cd /d C:\Fastload && fastload < " & newTableName & ".fl"
    output = ShellRun("cmd.exe /c cd /d C:\Fastload && fastload < " & newTableName & ".fl")
    
    filePath = "C:\Fastload\" & newTableName & ".log"
    Kill filePath
    Call FastLoadWrite(filePath, output)
    '
    ' Need to MERGE?
    If mergeTable Then Call FastloadMerge(fullTableName)
    '
    ' Extract return code
    '
    LOAD TextForm
    TextForm.txtBody = output
    If InStr(output, "Highest return code encountered = '0'") > 0 Then
        TextForm.txtHeader = "Success"
    Else
        TextForm.txtHeader = "Failed"
    End If
    TextForm.Show
    Unload TextForm

    Call UsageTracker("FastLoad", "Finished")
    Exit Sub
    
gotError:
    t = Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl
    MsgBox t, Title:="Fastload"
    Call UsageTracker("FastLoad", t)
    Stop
    Resume Next
    
End Sub

Sub FastloadMerge(fullTableName)
    
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

On Error GoTo gotError

10    qq = "INSERT INTO " & left(fullTableName, Len(fullTableName) - 3) & " SELECT * from " & fullTableName

        Debug_Print qq
20    Set DBCn = DBCheckConnection(DBCn)
30    Set DBRs = DBCheckRecordset(DBRs)

40    With DBRs
50        .CursorLocation = adUseServer ' adUseClient
60        .CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
70        .LockType = adLockOptimistic  ' adLockReadOnly
80        Set .ActiveConnection = DBCn
90    End With

100   DBRs.Open qq, DBCn

110   Exit Sub
    
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="FastloadMerge"
    Stop
    Resume Next
End Sub

Function CheckReservedWord(word)

    word = TrimReplace(word)
    ThisWorkbook.Worksheets("SQLReservedWords").Cells(1, 1) = word
    If Not IsError(ThisWorkbook.Worksheets("SQLReservedWords").Cells(2, 1)) Then
        CheckReservedWord = "a_" & word
    Else
        CheckReservedWord = word
    End If
    
End Function

Function IsReservedWord(word)

    ThisWorkbook.Worksheets("SQLReservedWords").Cells(1, 1) = word
    If Not IsError(ThisWorkbook.Worksheets("SQLReservedWords").Cells(2, 1)) Then
        IsReservedWord = True
    Else
        IsReservedWord = False
    End If
    
End Function

Sub FastLoadWrite(filePath, str)
Dim fso As Object
Dim oFile As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    '
    ' 1 - readonly
    ' 2 - writing
    ' 8 - append
    '
    ' 0 - Ascii format
    Set oFile = fso.OpenTextFile(filePath, 8, True, 0)
    
    oFile.WriteLine str
    oFile.Close

    Set fso = Nothing  ' for garbage collector
    Set oFile = Nothing

End Sub

Function CheckFastLoadTable()
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    With DBRs
        .CursorLocation = adUseClient ' adUseServer, adUseClient
        .CursorType = adUseClient ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
        .LockType = adLockOptimistic ' adLockReadOnly
        Set .ActiveConnection = DBCn
    End With
    
    On Error GoTo gotError
        Debug_Print GLBUserQuery
        Set DBCn = DBCheckConnection(DBCn)
        If DBCn Is Nothing Then Exit Function
        Set DBRs = DBCheckRecordset(DBRs)
        
        'useQuery = "select count(_fl_id) from dl_oge_analytics.delete_me"
        'DBRs.Open useQuery
        
        'fieldCount = DBRs.Fields.count
        'For i = 1 To fieldCount
            Debug_Print DBRs.DataSource
gotError:
    MsgBox "DBQuery Error (" & Err.Number & "): " & Err.Description, vbOKOnly, Title:="DBQuery ERROR"
    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
End Function


