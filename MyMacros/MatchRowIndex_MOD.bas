Attribute VB_Name = "MatchRowIndex_MOD"
Option Explicit

Sub MatchRowIndex(Optional fromRange, Optional toRange, Optional fromColRange, Optional toColRange)
Dim i As Long, j As Long
Dim lookupRange As Range, formulaRange As Range, col As Range
Dim WBTo As String, SHTo As String
Dim toCol As Long
Dim WBFrom As String, SHFrom As String
Dim fromCol As Long, botRowFrom As Long
Dim firstfromDataCol As Long, lastfromDataCol As Long
Dim fromHeader As String, fromHeaderColumn As Long
Dim toAddr As String
Dim useCol As Long, newCol As Long
Dim f As String
Dim fromLookBackCol As Long
Dim botRow As Long, botRowTo As Long
Dim firstToDataCol As Long
Dim fromIndex As Long


    If IsMissing(toRange) Then
        On Error Resume Next
            Set toRange = Application.InputBox("To IDs", Type:=8)
            Err.Clear
            On Error GoTo 0
        If IsMissing(toRange) Then Exit Sub
    End If
    
    If IsMissing(fromRange) Then
        On Error Resume Next
            Set fromRange = Application.InputBox("From IDs", Type:=8)
            Err.Clear
            On Error GoTo 0
        If IsMissing(fromRange) Then Exit Sub
    End If
    
    If IsMissing(fromColRange) Then
        On Error Resume Next
        Set fromColRange = Application.InputBox("Select Copy FROM Columns", Title:="MatchRowIndex", Type:=8)
        Err.Clear
        On Error GoTo 0
        If fromColRange Is Nothing Then Exit Sub
    End If
    
    If IsMissing(toColRange) Then
        On Error Resume Next
        Set toColRange = Application.InputBox("Select Copy TO Columns", Title:="MatchRowIndex", Type:=8)
        Err.Clear
        On Error GoTo 0
        If toColRange Is Nothing Then Exit Sub
    End If
    '
    '----------------------------------------------------------------------------------------------------------
    '
    WBTo = toRange.Parent.Parent.Name
    SHTo = toRange.Parent.Name
    toCol = toRange.Column
    
    WBFrom = fromRange.Parent.Parent.Name
    SHFrom = fromRange.Parent.Name
    fromCol = fromRange.Column
    firstfromDataCol = fromCol + 2 ' first col is key, second col is row index
    lastfromDataCol = LastColumn(SHFrom, WBFrom)
    '
    ' Use the column header to get the rownums
    fromHeader = Workbooks(WBFrom).Worksheets(SHFrom).Cells(1, fromCol)
    fromHeaderColumn = FindColumnHeader(fromHeader & "-RowIndex", SHFrom, WBFrom)
    If (fromHeaderColumn < 0) Then
        MsgBox "Can not find RowNums for: " & fromHeader, Title:="MatchRowIndex"
    End If
    botRowFrom = ColumnLastRow(fromCol, SHFrom, WBFrom)
    Set lookupRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(2, fromCol), _
                            Workbooks(WBFrom).Worksheets(SHFrom).Cells(botRowFrom, fromHeaderColumn))
    
    toAddr = Cells(2, toRange.Column).Address(False, False)
    toCol = toRange.Column
    useCol = fromHeaderColumn - fromCol + 1
    f = "=VLOOKUP(" & toAddr & "," & FullAddressPath(lookupRange) & "," & useCol & ",0)"

    newCol = ColumnInsertRight(toCol, SHTo, WBTo)
    fromLookBackCol = newCol  ' column with index IDs from WBTo
    botRow = ColumnLastRow(toCol, SHTo, WBTo)
    botRowTo = botRow
    '
    ' Lookback Index Column
    '
    Workbooks(WBTo).Worksheets(SHTo).Cells(1, newCol) = "" & WorkbookSheetPath(fromRange) & ""
    Call ColorRange(Workbooks(WBTo).Worksheets(SHTo).Cells(1, newCol), ORANGE)
    
    Set formulaRange = Range(Workbooks(WBTo).Worksheets(SHTo).Cells(2, newCol), _
                            Workbooks(WBTo).Worksheets(SHTo).Cells(botRow, newCol))
    formulaRange.NumberFormat = "General"
    formulaRange.Formula = f
    Call RangeToValues(formulaRange)
    formulaRange.NumberFormat = "0"
    
    firstToDataCol = NextColumn(SHTo, WBTo)
    
    Workbooks(WBTo).Worksheets(SHTo).Activate
    '
    ' Copy the FROM Header
    '
    'Set fromRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(1, firstFromDataCol),
    '                      Workbooks(WBFrom).Worksheets(SHFrom).Cells(1, lastFromDataCol))
    'fromRange.Copy Destination:=Workbooks(WBTo).Worksheets(SHTo).Cells(1, firstToDataCol)
    If toColRange Is Nothing Then Exit Sub
    
    firstToDataCol = toColRange.Column
    j = fromColRange.count
    j = firstToDataCol
    For Each col In fromColRange
        fromCol = col.Column
            
        Workbooks(WBFrom).Worksheets(SHFrom).Cells(1, fromCol).Copy _
        Destination:=Workbooks(WBTo).Worksheets(SHTo).Cells(1, j)
        Call ColorRange(Workbooks(WBTo).Worksheets(SHTo).Cells(1, j), LIGHTBLUE)
        j = j + 1
    Next col
    '
    ' Copy the FROM data
    '
    Call CalculationOff
    j = botRowTo
    On Error GoTo gotError
    For i = 2 To botRowTo
        fromIndex = Cells(i, fromLookBackCol) ' get the lookback index
        j = firstToDataCol
        For Each col In fromColRange
            fromCol = col.Column
            
            If IsError(fromIndex) Then
                Workbooks(WBTo).Worksheets(SHTo).Cells(i, j) = fromIndex
            Else
                'Workbooks(WBFrom).Worksheets(SHFrom).Cells(fromIndex, fromCol).Copy _
                'Destination:=Workbooks(WBTo).Worksheets(SHTo).Cells(i, j)
                Workbooks(WBTo).Worksheets(SHTo).Cells(i, j) = Workbooks(WBFrom).Worksheets(SHFrom).Cells(fromIndex, fromCol)
            End If
            j = j + 1
        Next col
    
        'Set fromRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(fromIndex, firstFromDataCol), _
        '                      Workbooks(WBFrom).Worksheets(SHFrom).Cells(fromIndex, lastFromDataCol))
        'fromRange.Copy Destination:=Workbooks(WBTo).Worksheets(SHTo).Cells(i, firstToDataCol)
        
    Next i
    '
    ' cleanup the LookBack column
    '
    Columns(newCol).Delete
    
    Call CalculationOn
    Call ClearClipboard
    Exit Sub
gotError:
    Resume Next
End Sub


