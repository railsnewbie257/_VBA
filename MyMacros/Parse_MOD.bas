Attribute VB_Name = "Parse_MOD"
Sub ParseEventTime()
Dim eventTimeRange As Range
Dim botRow As Long
    
    eventCol = FindColumnHeader("event_time")
    timeCol = ColumnInsertRight(eventCol)
    Cells(1, timeCol) = "EventDay"
    botRow = LastRow
    
    For i = 2 To botRow
        Cells(i, timeCol) = left(Cells(i, eventCol), 10)
    Next i
   
End Sub
'
' Extracts NIC timestamp
'
Sub Parse_Event_Desc()
    Application.DisplayStatusBar = True
    Application.StatusBar = "Parse_Event_Desc"
    
    fromCol = FindColumnHeader("Event_Desc")
    toCol = ColumnInsertRight(fromCol)
    toCol = fromCol + 1
    botRow = ColumnLastRow(fromCol)
    Application.StatusBar = "Parse_Event_Desc: Parse"
    For i = 2 To botRow
        timePos = (InStr(50, Cells(i, fromCol), "NIC timestamp: ")) + 15
        Cells(i, toCol) = (Mid(Cells(i, fromCol), timePos, 26))
    Next i
    
    Application.StatusBar = "Parse_Event_Desc: Done"
End Sub

Sub ParseEventIds()

    useCol = ActiveCell.Column
    useHeader = Cells(1, useCol)
    
    useCol = FindColumnHeader("event_log_id")
    If (useCol > 0) Then
        Call Parse_Event_Log_Id
    Else
        useCol = FindColumnHeader("event_external_id")
        If (useCol > 0) Then
            Call Parse_Event_External_Id
        Else
            MsgBox ("No Event ID column")
        End If
    End If
End Sub


Sub Parse_Event_External_Id()

    Application.DisplayStatusBar = True
    Application.StatusBar = "Parse_Event_External_Id"
    
    fromCol = FindColumnHeader("Event_External_Id")
    botRow = ColumnLastRow(fromCol)
    Columns(fromCol).AutoFit
    toCol = ColumnInsertRight(fromCol)
    
    Set toRange = Range(Cells(2, toCol), Cells(botRow, toCol))
    toRange.NumberFormat = "General"
    Set fromRange = Range(Cells(2, fromCol), Cells(2, fromCol))
    'fromRange.NumberFormat = "@"
    Debug.Print toRange.Address
    ' toRange.Formula = "=TEXT(" & "B2" & ",""00000000000"")"
    'toRange.Formula = "=TEXT(" & fromRange.Address(False, False) & ",""00000000000"")"
    Application.StatusBar = "Parse_Event_External_Id: Parse"
    toRange.Formula = "=TRIM(RIGHT(" & fromRange.Address(False, False) & ",LEN(" & fromRange.Address(False, False) & ")-2))"
    Call RangeToValues(toRange)
    toRange.NumberFormat = "@"
    
    Cells(1, toCol) = "Parse-" & Cells(1, fromCol)
    Call ColorRange(Cells(1, toCol), LIGHTGREEN)
    Call AddRowNumbers(toCol)
    Application.StatusBar = "Parse_Event_External_Id: Done"
    Application.DisplayStatusBar = False
End Sub
'
' event_log_id is from the LG dumps
'
Sub Parse_Event_Log_Id()

    Application.DisplayStatusBar = True
    Application.StatusBar = "Parse_Event_Log_Id"
    
    fromCol = FindColumnHeader("event_log_id")
    If (fromCol < 0) Then
        retCode = MsgBox("Wrong spreadsheet type ""eventlog_id"" not found, ABORTING", vbOKOnly)
        Exit Sub
    End If
    botRow = ColumnLastRow(fromCol)
    Columns(fromCol).AutoFit
    toCol = ColumnInsertRight(fromCol)
    
    Set toRange = Range(Cells(2, toCol), Cells(botRow, toCol))
    Set fromRange = Range(Cells(2, fromCol), Cells(2, fromCol))
    Debug.Print toRange.Address
    ' toRange.Formula = "=TEXT(" & "B2" & ",""00000000000"")"
    Application.StatusBar = "Parse_Event_External_Id: Parse"
    toRange.Formula = "=TEXT(" & fromRange.Address(False, False) & ",""00000000000"")"
    Call RangeToValues(toRange)
    Cells(1, toCol) = "Parse-" & Cells(1, fromCol)
    Call ColorRange(Cells(1, toCol), LIGHTGREEN)
    
    Call AddRowNumbers(toCol)
    
    Application.StatusBar = "Parse_Event_External_Id: Done"
    Application.DisplayStatusBar = False
End Sub

