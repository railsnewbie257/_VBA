Attribute VB_Name = "Duplicates_Mod"
Sub FilterDuplicateMeter()
Debug.Print "Start: " & Now()
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    Application.ScreenUpdating = False
    
    col1 = FindColumnHeader("meter_serial_num")
    col2 = FindColumnHeader("event_start_tm")
    
    Call SortSheetUp(col1, col2)
    
    botRow = ColumnLastRow(col1)
    For i = botRow - 1 To 2 Step -1
        If (Cells(i, col1) = Cells(i + 1, col1) And _
            Cells(i, col2) = Cells(i + 1, col2)) Then
            'Rows(i).Copy
            Rows(i).EntireRow.Delete
        End If
    Next i
Debug.Print "End: " & Now()
End Sub

Sub FilterRemovalDate()
    Debug.Print Now()
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    Application.ScreenUpdating = False
    
    useCol = FindColumnHeader("meter_removal_date")
    botRow = ColumnLastRow(useCol)
    For i = 2 To botRow
        If Not (Cells(i, useCol) = "12/31/9999") Then
            Rows(i).Delete
        End If
    Next i
    Call ClearClipboard
    Debug.Print Now()
End Sub

Sub DupByCol()
Dim aRange As Range
Dim useCol As Integer
Dim oldValue As String
Dim i As Long, startRow As Long
Dim count As Long
Dim useColor As Long

t = Selection.Address
t = ActiveCell.Address


    On Error Resume Next
    Set aRange = Selection
    Set aRange = Application.InputBox("Choose Column", Title:="DupByCol", Default:=Selection.Address, Type:=8)
    If aRange Is Nothing Then Exit Sub
    
    useCol = aRange.Column
    
    'Call SortSheetUp(useCol)
    
    oldValue = ""
    botRow = LastRow()
    startRow = 0
    Set aRange = Nothing
    For i = 2 To botRow + 1
        If Cells(i, useCol) <> oldValue Then
            If count > 1 Then
                useColor = ChooseColor(useColor)
                Range(Cells(startRow, useCol), Cells(i - 1, useCol)).Interior.color = useColor
            End If
            oldValue = Cells(i, useCol).Text
            startRow = i
            count = 1
        Else
            If aRange Is Nothing Then Set aRange = Range(Cells(startRow, useCol), Cells(startRow, useCol))
            count = count + 1
        End If

    Next i
    
    aRange.Select
    
    Exit Sub

    useCol = ActiveCell.Column
    runCol = ColumnInsertRight(useCol)
    botRow = ColumnLastRow(useCol)
    dups = 0
    useColor = 0
    runLength = 0
    For i = 2 To botRow - 1
        If Cells(i, useCol) = Cells(i + 1, useCol) Then
             If Not (old_dup = Cells(i, useCol)) Then
                useColor = ChooseColor(useColor)
                old_dup = Cells(i + 1, useCol)
                Cells(i - 1, runCol) = runLength
                runLength = 1
            End If
            Call ColorRange(Cells(i, useCol), useColor)
            Call ColorRange(Cells(i + 1, useCol), useColor)
            Call ColorRange(Cells(i, useCol).Offset(0, -1), useColor)
            Call ColorRange(Cells(i + 1, useCol).Offset(0, -1), useColor)
            dups = dups + 1
            runLength = runLength + 1
        End If
    Next i
    
    MsgBox dups & " duplicates found."
    
End Sub

Sub DupByRow()
Dim aRange As Range
Dim range1 As Range, range2 As Range
Dim useRow1 As Long, useRow2 As Long
Dim rightCol As Integer
Dim count As Integer: count = 0

    On Error Resume Next
    Set aRange = Nothing
    Set aRange = Application.InputBox("Select Rows To Compare", Title:="DupByRow", Default:=Selection.Address, Type:=8)
    If aRange Is Nothing Then Exit Sub
    
    Call RangeChooseTwo(aRange, range1, range2)
    
    useRow1 = range1.Row
    useRow2 = range2.Row
    rightCol = LastColumn()
    For i = 1 To rightCol
        If Cells(useRow1, i) = Cells(useRow2, i) Then
            Cells(useRow1, i).Interior.color = LIGHTBLUE
            Cells(useRow2, i).Interior.color = LIGHTBLUE
            count = count + 1
        End If
    Next i
    
    MsgBox count & " duplicates found out of " & rightCol & " possible."
End Sub


