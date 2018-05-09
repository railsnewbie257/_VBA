Attribute VB_Name = "Vlookup_MOD"
Sub CheckColumnMatches()

    On Error Resume Next
    Set targetRange = Application.InputBox("Select TARGET column", Title:="CheckMatches", Type:=8)
    If IsEmpty(targetRange) Then Exit Sub
    
    WBTarget = targetRange.Parent.Parent.Name
    SHTarget = targetRange.Parent.Name
    targetCol = targetRange.Column
    botRowTarget = LastRow(SHTarget, WBTarget)
    
    On Error Resume Next
    Set matchRange = Application.InputBox("Select column TO MATCH", Title:="CheckMatches", Type:=8)
    If IsEmpty(targetRange) Then Exit Sub
    
    WBMatch = matchRange.Parent.Parent.Name
    SHMatch = matchRange.Parent.Name
    matchCol = matchRange.Column
    botRowMatch = LastRow(SHMatch, WBMatch)
    
    Set lookupRange = Range(Workbooks(WBMatch).Worksheets(SHMatch).Cells(2, matchCol), _
                            Workbooks(WBMatch).Worksheets(SHMatch).Cells(botRowMatch, matchCol))
    
    formulaCol = ColumnInsertRight(targetCol, SHTarget, WBTarget)
    
    Set formulaRange = Range(Workbooks(WBTarget).Worksheets(SHTarget).Cells(2, formulaCol), _
                            Workbooks(WBTarget).Worksheets(SHTarget).Cells(botRowTarget, formulaCol))
    
    fff = "=VLOOKUP(" & Cells(2, targetCol).Address(False, False) & ",'[" & WBMatch & "]" & SHMatch & "'!" & lookupRange.Address & ",1,0)"
    
    formulaRange.Formula = fff
End Sub

Sub lkj()
Debug.Print Cells(5, 4).Address
End Sub
