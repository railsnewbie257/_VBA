Attribute VB_Name = "Split_MOD"
Sub ColumnValuesToTabs(Optional useCol)
Dim colRange As Range
Dim sortCol As Long
Dim i As Long, oldColor As Long

    If IsMissing(useCol) Then
        On Error Resume Next
        Set colRange = Nothing
        Set colRange = Application.InputBox("Select Column To Split", Default:=Selection.Address, Title:="SplitOnColumn", Type:=8)
        If colRange Is Nothing Then Exit Sub
    
        sortCol = colRange.Column
    Else
        sortCol = useCol
    End If
    
    Call ScreenOff
    
    Call SortSheetUp(sortCol)
    
    SHOrig = ActiveSheet.Name
    
    botRow = LastRow()
    oldValue = ""
    startRow = 2
    For i = 2 To botRow + 1
        If Cells(i, sortCol) <> oldValue Then
            '
            If oldValue <> "" Then
                Set headerRange = Range(Rows(1), Rows(1))
                Set moveRange = Range(Rows(startRow), Rows(i - 1))

                Call DeleteSheet(oldValue)

                Sheets.Add.Name = LegalSheetName(oldValue)
                ActiveSheet.Tab.color = oldColor
                    headerRange.Copy Destination:=Cells(1, 1)
                    moveRange.Copy Destination:=Cells(2, 1)
                Worksheets(SHOrig).Activate
                oldValue = Cells(i, sortCol)
                oldColor = Cells(i, sortCol).Interior.color
                startRow = i
            Else
                oldValue = Cells(i, sortCol)
            End If
        End If
    Next i
    
    Call ScreenOn
End Sub

Sub test()

    Set aRange = Range(Rows(1), Rows(1))
    Sheets.Add
    aRange.Copy Destination:=Cells(1, 1)
    
End Sub
