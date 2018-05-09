Attribute VB_Name = "Callbacks_MOD"
'
'Some Callbacks after Query() is run and rows downloaded
'

Sub MeterKeepsCallback()
    
    Call SortSheetDown(1)
    
End Sub

Sub BruceCallback()
    ActiveSheet.Rows.RowHeight = 15
End Sub

Sub MyTablesCallback()
    databaseCol = FindColumnHeader("DatabaseName")
    tableCol = FindColumnHeader("TableName")
    
    Call SortSheetUp(databaseCol, tableCol)
End Sub

