Attribute VB_Name = "Goto_MOD"
Sub GotoRow()
    useCol = ActiveCell.Column
    useRow = ActiveCell.Row
    currentColor = ChooseColor(currentColor)
    Call ColorRange(Rows(useRow), currentColor)
    
    'direct = Left(Cells(useRow, 1))
    
    SHTarget = Right(Cells(1, 1), Len(Cells(1, 1)) - 5)
    
    If (useCol = 1) Then
        targetRow = val(Cells(useRow, 1))
        targetCol = 1
    Else
        targetRow = val(Right(Cells(1, useCol), Len(Cells(1, useCol)) - 7))
        targetCol = val(Right(Cells(useRow, 1), Len(Cells(useRow, 1)) - 7))
    End If
    
    Worksheets(SHTarget).Activate
    
    rightColumn = LastColumn(SHTarget)
    Set targetRange = Range(Worksheets(SHTarget).Cells(targetRow, 1), Worksheets(SHTarget).Cells(targetRow, rightColumn))
    Call DrawBoxRange(targetRange)
    targetRange.Select
    Call ColorRange(targetRange, currentColor)

End Sub

Sub GotoWorksheet()
Dim aRange As Range, origRange As Range
Dim WBOrig As String, SHGoto As String
Dim WBUse As String, SHUse As String

    Call SpeedupOn
    Set origRange = ActiveCell
    
    WBOrig = ActiveWorkbook.Name
    
    SHGoto = "GotoSheet"
    On Error Resume Next
        Call AlertsOff
            Worksheets(SHGoto).Delete
        Call AlertsOn
        Worksheets.Add.Name = SHGoto
    
    With Worksheets(SHGoto).Cells(1, 2)
        .Value = "Workbook Name"
        .Font.Bold = True
        .Interior.color = LIGHTBLUE
    End With
    
    With Worksheets(SHGoto).Cells(1, 3)
        .Value = "Sheet Name"
        .Font.Bold = True
        .Interior.color = LIGHTBLUE
    End With
    
    k = 2
    For i = 1 To Workbooks.count
        For j = 1 To Workbooks(i).Worksheets.count
            If Not ((Workbooks(i).Name = WBOrig) And (Workbooks(i).Worksheets(j).Name = SHGoto)) Then
                Worksheets(SHGoto).Cells(k, 2) = Workbooks(i).Name
                Worksheets(SHGoto).Cells(k, 3) = Workbooks(i).Worksheets(j).Name
                useColor = Workbooks(i).Worksheets(j).Tab.color
                If useColor <> 0 Then Worksheets(SHGoto).Cells(k, 3).Interior.color = useColor
                Debug.Print "Tab color: " & Workbooks(i).Worksheets(j).Tab.color
                If Workbooks(i).Name = WBOrig Then
                    Worksheets(SHGoto).Cells(k, 2).Interior.color = YELLOW  ' current workbook is yellow
                    Worksheets(SHGoto).Cells(k, 1) = 0  ' current worksheet sorts to top
                ElseIf Workbooks(i).Name = MACROWORKBOOK Then
                    Worksheets(SHGoto).Cells(k, 2).Interior.color = 14540253  ' macro workbook is purple
                    Worksheets(SHGoto).Cells(k, 1) = 999
                Else
                    Worksheets(SHGoto).Cells(k, 1) = i
                End If
                k = k + 1
            End If
            
        Next j
    Next i
    
    Call SortSheetUp(1, 3, SHSort:=SHGoto)
    Worksheets(SHGoto).Columns(2).AutoFit
    Worksheets(SHGoto).Columns(3).AutoFit
    Worksheets(SHGoto).Columns(1).Delete
    Call SpeedupOff
    
    Set aRange = Nothing
    Set aRange = Application.InputBox("Select Sheet To GoTo", Type:=8, Title:="Goto Worksheet")
    If aRange Is Nothing Then ' cancelled
        Call AlertsOff
            Worksheets(SHGoto).Delete
        Call AlertsOn
        origRange.Activate
        Exit Sub
    End If
    
    Call ScreenOff
    WBUse = Cells(aRange.Row, 1)
    SHUse = Cells(aRange.Row, 2)
    Workbooks(WBUse).Worksheets(SHUse).Activate
    Call ScreenOn
    'Exit Sub
    Call AlertsOff
        Workbooks(WBOrig).Worksheets(SHGoto).Delete
    Call AlertsOn
End Sub

