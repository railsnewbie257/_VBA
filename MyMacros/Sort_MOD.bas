Attribute VB_Name = "Sort_MOD"
'Option Explicit

Sub ColumnMax()
    Set m = ThisColumnRange
    w = m.MAX
    MsgBox w
End Sub

Sub SortSheetUp(Optional sortCol1, Optional sortCol2, Optional sortCol3, Optional SHSort)
Dim botRow As Long
Dim retCode As Integer
Dim rightColumn As Integer
Dim sortRange As range
Dim selectRange As range
Dim keyRange1 As range
Dim keyRange2 As range
Dim origRange As range
Dim nCols As Integer, useCol As Integer

On Error GoTo gotError

10    ActiveCell.Select
20    Set origRange = ActiveCell

30    If IsMissing(sortCol1) Then  ' called from menu pick
40        If Selection.Areas.count >= 1 Then sortCol1 = Selection.Areas(1).Column
50        If Selection.Areas.count >= 2 Then sortCol2 = Selection.Areas(2).Column
60        If Selection.Areas.count >= 3 Then sortCol3 = Selection.Areas(3).Column
70    End If
    
    'If IsMissing(sortCol1) Then sortCol1 = ActiveCell.Column
80    If IsMissing(SHSort) Then SHSort = ActiveSheet.Name
    
      If Not IsMissing(sortCol1) Then
90      Set sortRange = Cells(1, sortCol1).CurrentRegion
        Set sortRange = ActiveCell.CurrentRegion
        sortRange.Copy
        hRow = HeaderRow(sortRange.Cells(1, 1))
        Set sortRange = range(Cells(hRow, sortRange.Column), Cells(sortRange.Row + sortRange.Rows.count + sortRange.Row - hRow, sortRange.Column + sortRange.Columns.count - 1))
        sortRange.Copy

        Set selectRange = sortRange
      End If

100    botRow = selectRange.Rows.count ' use this instead of selectRange.rows.count incase empty rows
110    If botRow = 0 Then
120        retCode = MsgBox("No columns to sort.", Title:="SortSheetUp")
130        Exit Sub
140    End If
    '
    ' Default selection is entire Sheet
    '
150    useCol = 1
    
160    rightColumn = selectRange.Columns.count
170    If ActiveSheet.Name = "Scratch" Then
180        If selectRange.Columns.count > 1 Then
190            retCode = MsgBox("Sort Only This Column?(YES) or entire SHEET (NO)?", vbYesNoCancel)
200            If retCode = vbCancel Then Exit Sub
210            If retCode = vbYes Then rightColumn = useCol
220        End If
230    End If
    'Set sortRange = Range(selectRange.Cells(1, useCol), Cells(botRow, rightColumn))
    'sortRange.Select
240    Set sortRange = selectRange ' Range(Cells(1, useCol), Cells(botRow, rightColumn))

250    'Set keyRange1 = Range(selectRange.Cells(1, sortCol1 - selectRange.Column + 1), selectRange.Cells(botRow, sortCol1 - selectRange.Column + 1))
260    'keyRange1.Select
    '
    ' Entire spreadsheet
    '
    ' Cells.Select
    '
270    Call ClearClipboard
    '
280    ActiveWorkbook.Worksheets(SHSort).Sort.SortFields.Clear

    '
    ' ActiveWorkbook.Worksheets(SHSort).Sort.SortFields.Add Key:=keyRange1, _

310    'ActiveWorkbook.Worksheets(SHSort).Sort.SortFields.Add Key:=keyRange1, _
       ' SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
320    If Not IsMissing(sortCol2) Then
        'botRow = Application.WorksheetFunction.Max(ColumnLastRow(sortCol2), botRow)
330        Set keyRange2 = range(selectRange.Cells(1, sortCol2 - selectRange.Column + 1), selectRange.Cells(botRow, sortCol2 - selectRange.Column + 1))
340        ActiveWorkbook.Worksheets(SHSort).Sort.SortFields.Add Key:=keyRange2, _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
350    End If

    If Not IsMissing(sortCol1) Then
        'botRow = Application.WorksheetFunction.Max(ColumnLastRow(sortCol2), botRow)
        Set keyRange1 = range(selectRange.Cells(1, sortCol1 - selectRange.Column + 1), selectRange.Cells(botRow, sortCol1 - selectRange.Column + 1))
        ActiveWorkbook.Worksheets(SHSort).Sort.SortFields.Add Key:=keyRange1, _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    End If


    
360    If Not IsMissing(sortCol3) Then
        'botRow = Application.WorksheetFunction.Max(ColumnLastRow(sortCol3), botRow)
370        Set keyRange3 = range(selectRange.Cells(1, sortCol3 - selectRange.Column + 1), selectRange.Cells(botRow, sortCol3 - selectRange.Column + 1))
380        ActiveWorkbook.Worksheets(SHSort).Sort.SortFields.Add Key:=keyRange3, _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
390    End If
    ' Finally do the sort
    '
400    With ActiveWorkbook.Worksheets(SHSort).Sort
410        .SetRange sortRange
420        .Header = xlYes
430        .MatchCase = False
440        .Orientation = xlTopToBottom
450        .SortMethod = xlPinYin
460        .Apply
470    End With

480    ActiveWorkbook.Worksheets(SHSort).Sort.SortFields.Clear
490    origRange.Select
        
        Exit Sub
        
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="SortSheetUp"
    Stop
    t = selectRange.Columns.count
    Resume Next
    
End Sub

Sub SortSheetDown(Optional sortCol1, Optional sortCol2, Optional sortCol3, Optional SHSort)
Dim botRow As Long
Dim retCode As Integer
Dim rightColumn As Integer
Dim selectRange As range
Dim sortRange As range
Dim keyRange1 As range
Dim keyRange2 As range
Dim origRange As range
Dim nCols As Integer, useCol As Integer

    Set origRange = ActiveCell
    useCol = ActiveCell.Column
    
    If IsMissing(sortCol1) Then
        If Selection.Areas.count >= 1 Then sortCol1 = Selection.Areas(1).Column
        If Selection.Areas.count >= 2 Then sortCol2 = Selection.Areas(2).Column
        If Selection.Areas.count >= 3 Then sortCol3 = Selection.Areas(3).Column
    End If
    
    'If IsMissing(sortCol1) Then sortCol1 = ActiveCell.Column
    If IsMissing(SHSort) Then SHSort = ActiveSheet.Name
    
    Set selectRange = ActiveCell.CurrentRegion
    
    botRow = selectRange.Rows.count
    If botRow = 0 Then
        retCode = MsgBox("No columns to sort.", Title:="SortSheetDown")
        Exit Sub
    End If
    '
    ' Default selection is entire Sheet
    '
    useCol = 1
    rightColumn = selectRange.Columns.count
    '
    If ActiveSheet.Name = "Scratch" Then
        If selectRange.Columns.count > 1 Then
            retCode = MsgBox("Sort Only This Column?(YES) or entire SHEET (NO)?", vbYesNoCancel)
            If retCode = vbCancel Then Exit Sub
            If retCode = vbYes Then rightColumn = useCol
        End If
    End If
    sortCol1 = sortCol1 - selectRange.Column + 1
    Set sortRange = range(selectRange.Cells(1, useCol), selectRange.Cells(botRow, rightColumn))
    Set keyRange1 = range(selectRange.Cells(1, sortCol1), selectRange.Cells(botRow, sortCol1))

    '
    ' Entire spreadsheet
    '
    ' Cells.Select
    '
    Call ClearClipboard
    '
    ActiveWorkbook.Worksheets(SHSort).Sort.SortFields.Clear
    '
    ActiveWorkbook.Worksheets(SHSort).Sort.SortFields.Add Key:=keyRange1, _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    
    If Not IsMissing(sortCol2) Then
        botRow = Application.WorksheetFunction.MAX(ColumnLastRow(sortCol1), botRow)
        Set keyRange2 = range(selectRange.Cells(1, sortSol2 - selectRange.Column + 1), selectRange.Cells(botRow, sortSol2 - selectRange.Column + 1))
        ActiveWorkbook.Worksheets(SHSort).Sort.SortFields.Add Key:=keyRange2, _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    End If
    
    If Not IsMissing(sortCol3) Then
        botRow = Application.WorksheetFunction.MAX(ColumnLastRow(sortCol3), botRow)
        Set keyRange3 = range(selectRange.Cells(1, sortSol3 - selectRange.Column + 1), selectRange.Cells(botRow, sortSol3 - selectRange.Column + 1))
        ActiveWorkbook.Worksheets(SHSort).Sort.SortFields.Add Key:=keyRange3, _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    End If
    ' Finally do the sort
    '
    With ActiveWorkbook.Worksheets(SHSort).Sort
        .SetRange sortRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ActiveWorkbook.Worksheets(SHSort).Sort.SortFields.Clear
    origRange.Select
End Sub
Sub Comment()
'
    Cells.Select
    ActiveWorkbook.Worksheets("Proximity").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Proximity").Sort.SortFields.Add Key:=range( _
        "A2:A17904"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Proximity").Sort.SortFields.Add Key:=range( _
        "B2:B17904"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Proximity").Sort.SortFields.Add Key:=range( _
        "D2:D17904"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Proximity").Sort
        .SetRange range("A1:K17904")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

