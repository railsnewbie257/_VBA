Attribute VB_Name = "Lookup_MOD"
Sub Lookup_Event_Id()
    Set toRange = Application.InputBox("Click on Target Column", Type:=8)
    toAddr = toRange.Address(False, False)
    toCol = toRange.Column
    SHTo = toRange.Parent.Name
    WBTo = toRange.Parent.Parent.Name
    
    Set fromRange = Application.InputBox("Click on From Column", Type:=8)
    fromAddr = fromRange.Address(False, False)
    fromCol = fromRange.Column
    SHFrom = fromRange.Parent.Name
    WBFrom = fromRange.Parent.Parent.Name
    
    ' check the From column header and see if there are row numbers next to it
    Debug.Print Workbooks(WBFrom).Worksheets(SHFrom).Cells(1, fromCol + 1).Value
    Debug.Print Workbooks(WBFrom).Worksheets(SHFrom).Cells(1, fromCol).Value & "-RowNums"
    If (Workbooks(WBFrom).Worksheets(SHFrom).Cells(1, fromCol + 1).Value = Workbooks(WBFrom).Worksheets(SHFrom).Cells(1, fromCol).Value & "-RowNums") Then
        hasRowNums = True
    End If
    
    If (hasRowNums) Then
        botRow = ColumnLastRow(fromCol, SHFrom, WBFrom)
        Set lookupRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(2, fromCol), _
                            Workbooks(WBFrom).Worksheets(SHFrom).Cells(botRow, fromCol + 1))
        Debug.Print lookupRange.Address
    End If
    
    f = "=VLOOKUP(" & toAddr & "," & FullAddressPath(lookupRange) & ",2,0)"
    Debug.Print f
    
    formulaCol = ColumnInsertRight(toCol)
    botRow = ColumnLastRow(toCol)
    Set formulaRange = Range(Cells(2, formulaCol), Cells(botRow, formulaCol))
    formulaRange.NumberFormat = "General"
    formulaRange.Formula = f
    Call RangeToValues(formulaRange)
    formulaRange.NumberFormat = "0"

End Sub

Function FullAddressPath(inRange)

    addrName = inRange.Address
    SHName = inRange.Parent.Name
    WBName = inRange.Parent.Parent.Name
    FullAddressPath = "'[" & WBName & "]" & SHName & "'!" & addrName
End Function

Function WbShPath(inRange)

    addrName = inRange.Address
    SHName = inRange.Parent.Name
    WBName = inRange.Parent.Parent.Name
    WbShPath = "'[" & WBName & "]" & SHName & "'"
End Function

Sub CompareColumnValues()

    Set rownumRange = Application.InputBox("Row Number column", Type:=8)
    
    Set targetRange = Application.InputBox("Target Column", Type:=8)
    
    Set fromRange = Application.InputBox("From column", Type:=8)
    
    SHTarget = targetRange.Parent.Name
    WBTarget = targetRange.Parent.Parent.Name
    Workbooks(WBTarget).Worksheets(SHTarget).Activate
    newCol = ColumnInsertRight(targetRange.Column)
    
    WBFrom = fromRange.Parent.Parent.Name
    SHFrom = fromRange.Parent.Name
    fromCol = fromRange.Column
    
    rownumCol = rownumRange.Column
    botRow = ColumnLastRow(rownumCol)
    
    For i = 2 To botRow
        fromRow = Workbooks(WBTarget).Worksheets(SHTarget).Cells(i, rownumCol).Value
        If IsError(fromRow) Then
            Workbooks(WBTarget).Worksheets(SHTarget).Cells(i, newCol) = "#N/A"
        Else
            Workbooks(WBTarget).Worksheets(SHTarget).Cells(i, newCol) = Workbooks(WBFrom).Worksheets(SHFrom).Cells(fromRow, fromCol).Value
        End If
    Next i
    
    Call ClearClipboard
End Sub

Sub AddColumnValues(Optional indexRange, Optional fromRange, Optional targetRange)
Dim t2 As String

    If IsMissing(indexRange) Then
        On Error Resume Next
            Set indexRange = Application.InputBox("Row Index column", Type:=8)
            Err.Clear
            On Error GoTo 0
        If IsEmpty(indexRange) Then Exit Sub
    End If
    
    If IsMissing(fromRange) Then
        On Error Resume Next
            Set fromRange = Application.InputBox("From column", Type:=8)
            Err.Clear
            On Error GoTo -1
        If IsEmpty(fromRange) Then Exit Sub
    End If
    
    t = WbShPath(fromRange)
    t = Right(t, Len(t) - 1)
    WBindex = indexRange.Parent.Parent.Name
    SHindex = indexRange.Parent.Name

   Debug.Print Workbooks(WBindex).Worksheets(SHindex).Cells(1, indexRange.Column).Value
    If Not t = Workbooks(WBindex).Worksheets(SHindex).Cells(1, indexRange.Column) Then
        retCode = MsgBox("From source is incorrect, should be " & t)
        Exit Sub
    End If
        
    If IsMissing(targetRange) Then
        On Error Resume Next
            retCode = MsgBox("Select Column (NO = Append)?", vbYesNoCancel)
            If (retCode = vbCancel) Then Exit Sub
            If (retCode = vbYes) Then
                Set targetRange = Application.InputBox("Target Column", Type:=8)
                newCol = ColumnInsertRight(targetRange.Column - 1)
            Else
                newCol = NextColumn(ActiveSheet.Name)
                Set targetRange = Range(Cells(1, newCol), Cells(1, newCol))
            End If
    Else
        newCol = ColumnInsertRight(targetRange.Column - 1)
        ' Err.Clear
        ' On Error GoTo 0
        ' If IsEmpty(targetRange) Then Exit Sub
    End If

        
    SHTarget = targetRange.Parent.Name
    WBTarget = targetRange.Parent.Parent.Name
    Workbooks(WBTarget).Worksheets(SHTarget).Activate
    ' newCol = ColumnInsertRight(targetRange.Column)
    
    WBFrom = fromRange.Parent.Parent.Name
    SHFrom = fromRange.Parent.Name
    fromCol = fromRange.Column
    
    indexCol = indexRange.Column
    botRow = ColumnLastRow(indexCol, SHTarget, WBTarget)
    
    Workbooks(WBTarget).Worksheets(SHTarget).Cells(1, newCol) = Workbooks(WBFrom).Worksheets(SHFrom).Cells(1, fromCol).Value
    Call ColorRange(Workbooks(WBTarget).Worksheets(SHTarget).Cells(1, newCol), LIGHTBLUE)
    For i = 2 To botRow
        fromRow = Workbooks(WBTarget).Worksheets(SHTarget).Cells(i, indexCol).Value
        If IsError(fromRow) Then
            Workbooks(WBTarget).Worksheets(SHTarget).Cells(i, newCol) = ""
        Else
            Workbooks(WBTarget).Worksheets(SHTarget).Cells(i, newCol) = Workbooks(WBFrom).Worksheets(SHFrom).Cells(fromRow, fromCol).Value
        End If
    Next i
    
    Call ClearClipboard
End Sub

Sub MatchRowIndex(Optional fromRange, Optional toRange)

    If IsMissing(toRange) Then
        On Error Resume Next
            Set toRange = Application.InputBox("To IDs", Type:=8)
            Err.Clear
            On Error GoTo 0
        If IsEmpty(toRange) Then Exit Sub
    End If
    
    If IsMissing(fromRange) Then
        On Error Resume Next
            Set fromRange = Application.InputBox("From IDs", Type:=8)
            Err.Clear
            On Error GoTo 0
        If IsEmpty(fromRange) Then Exit Sub
    End If
    
    WBTo = toRange.Parent.Parent.Name
    SHTo = toRange.Parent.Name
    toCol = toRange.Column
    
    WBFrom = fromRange.Parent.Parent.Name
    SHFrom = fromRange.Parent.Name
    fromCol = fromRange.Column
    
    FromHeader = Workbooks(WBFrom).Worksheets(SHFrom).Cells(1, fromCol)
    fromHeaderColumn = FindColumnHeader(FromHeader & "-RowNum", SHFrom, WBFrom)
    botRow = ColumnLastRow(fromCol, SHFrom, WBFrom)
    Set lookupRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(2, fromCol), _
                            Workbooks(WBFrom).Worksheets(SHFrom).Cells(botRow, fromHeaderColumn))
    Debug.Print lookupRange.Address
    
    toAddr = Cells(2, toRange.Column).Address(False, False)
    toCol = toRange.Column
    useCol = fromHeaderColumn - fromCol + 1
    f = "=VLOOKUP(" & toAddr & "," & FullAddressPath(lookupRange) & "," & useCol & ",0)"
    Debug.Print f
    
    newCol = ColumnInsertRight(toCol, SHTo, WBTo)
    botRow = ColumnLastRow(toCol, SHTo, WBTo)
    
    Workbooks(WBTo).Worksheets(SHTo).Cells(1, newCol) = "" & WbShPath(fromRange) & ""
    Call ColorRange(Workbooks(WBTo).Worksheets(SHTo).Cells(1, newCol), ORANGE)
    
    Set formulaRange = Range(Workbooks(WBTo).Worksheets(SHTo).Cells(2, newCol), _
                            Workbooks(WBTo).Worksheets(SHTo).Cells(botRow, newCol))
    formulaRange.NumberFormat = "General"
    formulaRange.Formula = f
    Call RangeToValues(formulaRange)
    formulaRange.NumberFormat = "0"
    
    Workbooks(WBTo).Worksheets(SHTo).Activate
End Sub



