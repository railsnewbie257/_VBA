Attribute VB_Name = "Utillib_MOD"
Sub DefaultWorkbookAndSheet(SHUse, WBUse)
    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
End Sub

Sub HelpVideos()
    ActiveWorkbook.FollowHyperlink _
      Address:="https://ppihoge.github.io/Videos/", _
      NewWindow:=True
End Sub

Sub MyMacros()
    ThisWorkbook.Activate
End Sub

Function MacroTimestamp() As String
    MacroTimestamp = ThisWorkbook.Worksheets("Pallette").Cells(8, 1)
End Function

Sub ShowVersion()
    MsgBox "TimeStamp: " & MacroTimestamp & vbNewLine & vbNewLine & "WorkBook Type: " & IdentifyWorkbookType()
    
End Sub

Sub ClearSelection()
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub ClearClipboard()
    Application.CutCopyMode = False
End Sub

Sub AutoFilterOff()
'
' if autofilter on a table
'
ActiveSheet.ListObjects(1).AutoFilter.ShowAllData
ActiveSheet.ListObjects(1).ShowAutoFilter = False
End Sub
Sub NumberOfRows()
Dim t As Long
    MsgBox "Number of rows:  " & format(ActiveCell.CurrentRegion.Rows.count - 1, "#,##0")
    'MsgBox "Number of rows:  " & format(LastRow, "#,##0")
End Sub
Function NextRow(Optional SHUse, Optional WBUse)
On Error GoTo Err1:
    NextRow = LastRow(SHUse, WBUse) + 1
    Exit Function
Err1:
    LastRow = 0
    
End Function

Sub ItemCount()
    n = ColumnCountA(ActiveCell.Column)
    nRows = ColumnLastRow(ActiveCell.Column) - 1 ' -1 for header
    MsgBox (n & " in " & nRows & " rows")
    
End Sub

Function HasHeader()
    retCode = MsgBox("Has Headers?", vbYesNo)
    If retCode = vbYes Then
        HasHeader = True
    Else
        HasHeader = False
    End If
    
End Function

Function RangeToText(aRange, Optional rel_or_abs, Optional format)
Dim t2 As range
    'Set t2 = aRange
    'debug_print aRange.Parent.Name
    'debug_print aRange.Address(external:=False)
    RangeToText = "'[" & aRange.Parent.Parent.Name & "]" & aRange.Parent.Name & "'!" & aRange.Address(False, False, xlA1)
End Function

Function DBTableNextAlias()

    LastCol = RowLastColumn(1)
    max_c = Chr(Asc("a") - 1)
    For i = 1 To LastCol
        c = GetTableAlias(Cells(1, i))
        If c > max_c Then max_c = c
    Next i
    
    DBTableNextAlias = Chr(Asc(max_c) + 1)
    
End Function

Function GetTableAlias(s)
    GetTableAlias = Mid(s, Len(s), 1)
End Function

Sub AreaConcat()
Dim useRange As range, destRange As range
Dim i As Long, j As Long, k As Integer
Dim botRow As Long
Dim colCount As Integer
Dim columnList() As String
Dim columnLen() As Long
Dim destCol As Long
Dim maxRows As Long

    WBOrig = ActiveWorkbook.Name
    SHOrig = ActiveSheet.Name
   
    On Error Resume Next
    Set useRange = Nothing
    Set useRange = Application.InputBox("Please Select Range(s) to Concatenate", Title:="DoConcat", Default:=Selection.Address, Type:=8)
    If useRange Is Nothing Then Exit Sub
    On Error GoTo 0
    
    On Error Resume Next
    Set destRange = Nothing
    Set destRange = Application.InputBox("Please Select Destination Column", Title:="DoConcat", Type:=8)
    If destRange Is Nothing Then Exit Sub
    On Error GoTo 0
    destCol = destRange.Column
    
    SHConcat = "Concat"
    
    Call DeleteSheet("ConCat")

    Sheets.Add
    ActiveSheet.Name = "ConCat"
    
    'debug_print UBound(columnList)
    maxRows = 0
    For i = 1 To useRange.Areas.count
        If useRange.Areas(i).Rows.count > maxRows Then maxRows = useRange.Areas(i).Rows.count
    Next i
    
    colCount = 0
    For i = 1 To useRange.Areas.count
        For j = 1 To useRange.Areas(i).Columns.count
            baseCol = useRange.Areas(i).Column
            Debug_Print "rows > " & useRange.Areas(i).Row & "," & useRange.Areas(i).Rows.count
            t = left(Cells(1, baseCol + j - 1).Address(False, False), 1)
            '
            '
            nRows = useRange.Areas(i).Rows.count
            useRange.Areas(i).Cells.Copy Destination:=Worksheets(SHConcat).Cells(1, i * j)
            '
            ' Autextend on scratch sheet
            '
            With Worksheets(SHConcat)
            Set aRange = .range(.Cells(1, i * j), .Cells(nRows, i * j))
            Set bRange = .range(.Cells(1, i * j), .Cells(maxRows, i * j))
            On Error Resume Next
            aRange.AutoFill Destination:=bRange, Type:=xlFillDefault
            End With
            
            colCount = colCount + 1
            
            If (Not columnList) = True Then
                ReDim Preserve columnList(1)
            Else
                ReDim Preserve columnList(UBound(columnList) + 1)
            End If
            columnList(UBound(columnList)) = t
            Debug_Print "columnList >" & UBound(columnList)
        Next j
    Next i
    
    fff = "="
    For i = 1 To colCount - 1
        fff = fff & Cells(1, i).Address(False, False, xlA1) & "&"
    Next i
    fff = fff & Cells(1, colCount).Address(False, False, xlA1)
        
    With Worksheets(SHConcat)
        Set aRange = .range(.Cells(1, newCol), .Cells(maxRows, newCol))
    End With
    
    aRange.Formula = fff
    Call RangeToValues(aRange)
    

    botRow = ColumnLastRow(destCol, destRange.Parent.Name, destRange.Parent.Parent.Name)
    If botRow <> 0 Then newCol = ColumnInsertRight(destCol - 1, destRange.Parent.Name, destRange.Parent.Parent.Name)
    aRange.Copy Destination:=destRange
    
End Sub

