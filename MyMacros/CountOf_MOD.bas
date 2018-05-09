Attribute VB_Name = "CountOf_MOD"
Sub AddCountCol(Optional useCol)
Dim aRange As Range
Dim countCol As Long
Dim botRow As Long
Dim i As Long, count As Long
Dim firstRow As Long
Dim oldValue As String
    
    If IsMissing(useCol) Then
        On Error Resume Next
        Set aRange = Nothing
        Set aRange = Application.InputBox("Choose Column", Title:="Count Of", Default:=Selection.Address, Type:=8)
        If aRange Is Nothing Then Exit Sub
        useCol = aRange.Column
    Else
        
    End If
    
    Call SortSheetUp(useCol)
    
    countCol = ColumnInsertRight(useCol)
    Cells(1, countCol) = "Count of " & Cells(1, useCol)
    botRow = LastRow()
    count = 0
    oldValue = ""
    For i = botRow To 1 Step -1  ' slight cheat since uses the column headers row=1 as the end point
        If Cells(i, useCol) <> oldValue Then
            If oldValue <> "" Then
                Cells(startRow, countCol) = count
                If (startRow - 1) >= (i + 1) Then Range(Rows(startRow - 1), Rows(i + 1)).Delete
            End If
            startRow = i
            oldValue = Cells(startRow, useCol)
            count = 1
        Else
            count = count + 1
        End If
    Next i
        
    
End Sub
