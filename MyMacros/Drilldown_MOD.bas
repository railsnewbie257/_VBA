Attribute VB_Name = "Drilldown_MOD"
Sub EventsDrilldown()

    lasteventCol = FindColumnHeader("Last_Event_Time")
    rundateCol = FindColumnHeader("RunDate")
    
    Call SortSheetUp(rundateCol, lasteventCol)

    eventCol = FindColumnHeader("Event_External_Event_Cd")
    botRow = ColumnLastRow(eventCol)
    
    For i = 2 To botRow
        eventCode = Cells(i, eventCol)
        
        Select Case eventCode
            
            Case 12007
                useColor = RED
            Case 15035, 15036
                useColor = LIGHTGREEN
            Case 100007
                useColor = GREENfo
            Case Else
                useColor = NOCOLOR
        End Select
        
        Call ColorRange(Cells(i, eventCol), useColor)
    Next i
    
End Sub
