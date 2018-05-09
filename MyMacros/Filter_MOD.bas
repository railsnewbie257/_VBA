Attribute VB_Name = "Filter_MOD"
Sub FilterProximity()
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$U$12521").AutoFilter Field:=3, Criteria1:="5"
End Sub

Sub filterNas()

    FilterNaForm.Show
    
End Sub
'
' This will rollup multiple workorders on a meter to one line
'
Sub filterMultipleWorkOrders()
Dim aRange As Range
Dim meterCol As Integer, workorderCol As Integer
Dim botRow As Long

    meterCol = FindColumnHeader("METER_SERIAL_NUM")
    workorderCol = FindColumnHeader("WORK_ORDER_TYPE_CD")
    botRow = ColumnLastRow(meterCol)
    
    Call SortSheetUp(meterCol)
    
    oldValue = ""
    i = 2
    While i < botRow
        While (Cells(i, meterCol) = Cells(i + 1, meterCol))  ' same meter on 2 lines
         
                Range(Rows(i), Rows(i + 1)).Copy
                If Cells(i, workorderCol) <> "" Then
                    Cells(i, workorderCol) = Cells(i, workorderCol) & "," & Cells(i + 1, workorderCol)
                End If
            
                Rows(i + 1).Delete
                botRow = botRow - 1
        Wend
        i = i + 1
    Wend
    '
    ' highlight workorders in red
    '
    Set aRange = RangeHasValues(Columns(workorderCol))
    If Not aRange Is Nothing Then aRange.Interior.color = RED
    
    aCol = FindColumnHeader("Work_Order_Type_Desc")
    Set aRange = RangeHasValues(Columns(aCol))
    If Not aRange Is Nothing Then aRange.Interior.color = RED
End Sub
