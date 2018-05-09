Attribute VB_Name = "Header_MOD"
'
' Determine which row the column headers are on
'
' The format for headers are:
'   Tablename header is in BLUE and ORANGE and 3 cells wide, and name has a "."
'   Column headers are in bold and black
'
'----------------------------------------------------------------------------------------
'
' Check for DBTableName
'
Function IsTableNameFormat(aCell) As Boolean
aCell.Copy
    With aCell
        If .Font.Bold = True _
            And .Font.color = BLUE _
            And .Interior.color = ORANGE _
            And InStr(aCell, ".") > 0 _
            And .Offset(0, 1).Interior.color = LIGHTGREEN _
            And aCell.Offset(0, 2) = "<<< QUERY" Then
                IsTableNameFormat = True
        Else
            IsTableNameFormat = False
        End If
    End With
End Function
'
' Check if the row being passed is bolded which means it's a header row
'
Function IsRowOfBold(aCell) As Boolean
Dim nCols As Integer
Dim count As Integer

    nCols = RowLastColumn(aCell.Row)
    For i = aCell.Column To nCols
        If Cells(aCell.Row, i).Font.Bold = True Then count = count + 1
    Next i
    If count = nCols - aCell.Column + 1 Then
        IsRowOfBold = True
    Else
        IsRowOfBold = False
    End If
End Function
'
' Returns the row with column headers, there may be a DBTableName row above it
'
Function HeaderRow(useRange) As Long
Dim aRegion As range
Dim t As range
    
    Set thisRegion = useRange.CurrentRegion
    
    HeaderRow = 1  ' default
    
    If IsTableNameFormat(thisRegion.Cells(1, 1)) Then
        If IsRowOfBold(thisRegion.Cells(2, 1)) Then HeaderRow = thisRegion.Row + 1
    Else
        If IsRowOfBold(thisRegion.Cells(1, 1)) Then HeaderRow = thisRegion.Row
    End If

End Function

Sub q()
    t = HeaderRow(ActiveCell)
End Sub
