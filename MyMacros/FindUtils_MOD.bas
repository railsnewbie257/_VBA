Attribute VB_Name = "FindUtils_MOD"
Function FindInRangeExact(searchValue, searchRange)
    Set FindInRangeExact = FindInRange(searchValue, searchRange, xlWhole)
End Function

Function FindInRange(searchValue, searchRange, Optional useLookAt)
On Error GoTo None:

    If (IsMissing(useLookAt)) Then useLookAt = xlPart

    searchRange.Copy
    
    ' following line incase searchValue is in the first cell fo the searchRange
    Set startRange = searchRange.Cells(searchRange.Rows.Count, searchRange.Columns.Count)
    startRange.Copy
   Set resultRange = searchRange.Find(What:=searchValue, _
                    after:=startRange, _
                    LookAt:=useLookAt, _
                    LookIn:=xlValues, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    searchformat:=False, _
                    MatchCase:=False)
    If Not (resultRange Is Nothing) Then
        startAddress = resultRange.Address
        Set findRange = resultRange
        Do
            'Debug.Print findRange.Address
            Set findRange = searchRange.FindNext(after:=findRange)
            If (findRange.Address = startAddress) Then Exit Do
            findRange.Copy
            Set resultRange = Application.Union(resultRange, findRange)
        Loop
    End If
    Set FindInRange = resultRange
    Exit Function
None:
    Set FindInRange = Nothing

End Function

Function FindColumnHeader(columnName, Optional SHUse, Optional WBUse)
    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    
    Set sRange = Workbooks(WBUse).Worksheets(SHUse).Rows(1)
    Set fRange = FindInRange(columnName, sRange)
    
    If Not fRange Is Nothing Then
        FindColumnHeader = fRange(1).Column
        Exit Function
    Else
        FindColumnHeader = -1
    End If
End Function

Sub FindTest()
    col = FindColumnHeader("First_event_time")
End Sub
