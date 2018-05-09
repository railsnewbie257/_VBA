Attribute VB_Name = "Highlight_MOD"
Sub HeaderBold()
    Rows(1).Font.Bold = True
End Sub

Sub CopyHighlights()
Dim fromRange As Range
Dim toRange As Range

    On Error Resume Next
    Set fromRange = Application.InputBox("Select Column FROM Workbook", Type:=8)
    If fromRange Is Nothing Then Exit Sub
    
    SHFrom = fromRange.Parent.Name
    WBFrom = fromRange.Parent.Parent.Name
    useColFrom = fromRange.Column
    botRowFrom = ColumnLastRow(useColFrom, SHFrom, WBFrom)
    
    fromHeader = Workbooks(WBFrom).Worksheets(SHFrom).Cells(1, useColFrom)
    
    Set toRange = Application.InputBox("Select Column TO Workbook", Type:=8)
    If toRange Is Nothing Then Exit Sub
    
    SHTo = toRange.Parent.Name
    WBTo = toRange.Parent.Parent.Name
    useColTo = toRange.Column
    botRowTo = ColumnLastRow(useColTo, SHTo, WBTo)
    Set sRange = Range(Workbooks(WBTo).Worksheets(SHTo).Cells(2, useColTo), _
                       Workbooks(WBTo).Worksheets(SHTo).Cells(botRowTo, useColTo))
                       
    ToHeader = Workbooks(WBTo).Worksheets(SHTo).Cells(1, useColTo)
    
    If ToHeader <> fromHeader Then
        retCode = MsgBox("Column Names Do NOT Match, Proceed?", vbYesNo)
        If retCode = vbNo Then Exit Sub
    End If
    
    For i = 2 To botRowFrom
    
        If Workbooks(WBFrom).Worksheets(SHFrom).Cells(i, useColFrom).Interior.Pattern = 1 Then
            useValue = Workbooks(WBFrom).Worksheets(SHFrom).Cells(i, useColFrom).Value
            useColor = Workbooks(WBFrom).Worksheets(SHFrom).Cells(i, useColFrom).Interior.color
            
            Set fRange = FindInRange(useValue, sRange)
            
            Debug.Print fRange.Row
            Call ColorRange(Workbooks(WBTo).Worksheets(SHTo).Rows(fRange.Row), useColor)
        End If
    
    Next i
End Sub

Sub FindLastGasps()

    t = ColumnNumToLetter(3)

    eventCol = FindColumnHeader("event_external_event_cd")
    rundateCol = FindColumnHeader("rundate")
    timeCol = FindColumnHeader("event_start_tm")
    
    newCol = ColumnInsertRight(eventCol)
    botRow = ColumnLastRow(eventCol)
    
    For i = 2 To botRow
        If Cells(i, eventCol) = 12007 Or Cells(i, eventCol) = 15035 Then
            Cells(i, newCol) = "Off"
        ElseIf Cells(i, eventCol) = 100007 Or Cells(i, eventCol) = 15036 Then
            Cells(i, newCol) = "On"
        Else
            Cells(i, newCol) = "Unknown"
        End If
    Next i
    Cells(1, newCol) = "MeterStatus"
    
    dayCol = ColumnInsertRight(newCol)
    Range(Cells(2, dayCol), Cells(botRow, dayCol)).Formula = "=WEEKDAY(DATEVALUE(TEXT(A2,""mm/dd/yyyy"")), 1)"
    Cells(1, dayCol) = "Weekday"
    
    hourCol = ColumnInsertRight(dayCol)
    timeCol = FindColumnHeader("event_start_tm")
    Range(Cells(2, hourCol), Cells(botRow, hourCol)).Formula = "=VALUE(TEXT(" & ColumnNumToLetter(timeCol) & "2,""hh""))"
    Cells(1, hourCol) = "Hour"
    Call SortSheetUp(newCol, rundateCol, timeCol)

End Sub

Sub HighlightEvents()
Dim aRange As Range
Dim useCol As Long, botRow As Long

    On Error Resume Next
    Set aRange = Nothing
    Set aRange = Application.InputBox("Select Events Column", Default:=Selection.Address, Title:="HighligtsEvents", Type:=8)
    If aRange Is Nothing Then Exit Sub
    
    useCol = aRange.Column
    botRow = ColumnLastRow(useCol)
    For i = 2 To botRow
        Select Case Cells(i, useCol)
        
            Case 12007
                Call ColorRange(Cells(i, useCol), RED)
            Case 100007
                Call ColorRange(Cells(i, useCol), GREEN)
            Case 15035
                Call ColorRange(Cells(i, useCol), LIGHTBLUE)
            Case 15036
                Call ColorRange(Cells(i, useCol), BLUE)
            Case 15105
                Call ColorRange(Cells(i, useCol), PURPLE)
            Case Else
                Call ColorRange(Cells(i, useCol), LIGHTGREY)
        End Select
    Next i
    
End Sub

