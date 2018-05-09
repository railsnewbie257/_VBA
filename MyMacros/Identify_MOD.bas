Attribute VB_Name = "Identify_MOD"
'
' Identify the sheet type based on the column headers
'
Function IdentifySheet(Optional SHUse, Optional WBUse)

    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name

    With Workbooks(WBUse).Worksheets(SHUse)
        If (UCase(.Cells(1, 1)) = "EVENT_LOG_ID" And _
            UCase(.Cells(1, 2)) = "EVENT_ID" And _
            UCase(.Cells(1, 3)) = "EVENT_NAME") Then
            IdentifySheet = "SSN"
        
        ElseIf (UCase(.Cells(1, 1)) = "RUNDATE" And _
            UCase(.Cells(1, 2)) = "METER_SERIAL_NUM" And _
            UCase(.Cells(1, 3)) = "NUM_OF_12007") Then
            IdentifySheet = "LASTGASP"
            
        ElseIf (UCase(.Cells(1, 1)) = "_FL_ID") Then
            IdentifySheet = "FASTLOAD"
        
        ElseIf (UCase(.Cells(2, 1)) = "REQUEST TEXT") Then
            IdentifySheet = "SHOWTABLE"
            
        ElseIf LastColumn() = 0 And LastRow() = 0 Then
            IdentifySheet = "EMPTY"
            
        ElseIf Cells(1, 1).Interior.color = ORANGE And SheetExists("ColumnNames") Then
            IdentifySheet = "ColumnNames"
            
        ElseIf left(ActiveSheet.Name, 5) <> "Sheet" Then
            IdentifySheet = UCase(ActiveSheet.Name)
        Else
            IdentifySheet = "UNKNOWN"
        End If
    End With
End Function

Function IdentifyWorkbookType(Optional WBName) As String

    If IsMissing(WBName) Then WBName = ActiveWorkbook.Name

    If SheetExists("LastGasp", WBName) Then
        IdentifyWorkbookType = "LastGasp"
    ElseIf SheetExists("UsageDrop", WBName) Then
        IdentifyWorkbookType = "UsageDrop"
    ElseIf SheetExists("PhaseAngleAlarm", WBName) Then
        IdentifyWorkbookType = "PhaseAngleAlarm"
    ElseIf SheetExists("UnderVoltage", WBName) Then
        IdentifyWorkbookType = "UnderVoltage"
    ElseIf SheetExists("ReceivedEnergy", WBName) Then
        IdentifyWorkbookType = "ReceivedEnergy"
    ElseIf SheetExists("ZeroKWH", WBName) Then
        IdentifyWorkbookType = "ZeroKWH"
    Else
        If ActiveWorkbook.Worksheets.count = 1 Then
            IdentifyWorkbookType = Worksheets(1).Name
        Else
            IdentifyWorkbookType = "?? Unknown ??"
        End If
    End If
        
End Function
