Attribute VB_Name = "GetValues_MOD"
'
' Returns a standarized date format "YYYY-MM-DD"
'
Function GetDate(Optional inDate)
    If IsMissing(inDate) Then inDate = format(Now(), "YYYY-MM-DD")
        
    inDate = InputBox(prompt:="Enter Date YYYY-MM-DD or MM/DD/YYYY", Default:=inDate, Title:="Enter Date")

    If Not inDate = "" Then
        GetDate = format(inDate, "yyyy-mm-dd")
    Else
        GetDate = ""
    End If
End Function


Function SelectValuesInColumnRange()
Dim resultRange As Range
Dim findRange As Range
Dim searchRange As Range

    On Error GoTo Out:
    searchValue = ActiveCell.Value
    SHUse = ActiveSheet.Name
    useCol = ActiveCell.Column
    Set searchRange = Range(Cells(1, useCol), Cells(LastRow(SHUse), useCol))
    
    Debug.Print searchRange.Address
    Set resultRange = FindInRange(searchValue, searchRange)
        
    resultRange.Copy
    Set SelectValuesInColumnRange = resultRange
    Exit Function

None:
    Set SelectValuesInColumnRange = Nothing
    Exit Function
Out:
    MsgBox (Err.Description)
End Function

Sub ExtractHeaders()

    If (Selection.count = 1) Then
        retCode = MsgBox("Only Extract 1 cell?", vbyesyno)
        If retCode = vbNo Or retCode = vbCancel Then Exit Sub
    End If
    Set headerRange = Selection
    headerRange.Copy
    
    If Not SheetExists(False, "Headers") Then
        Call MakeScratchSheet(False, "Headers")
     
    End If
    useCol = NextColumn("Headers")
    Worksheets("Headers").Cells(1, useCol).PasteSpecial Paste:=xlPasteAll, _
        Operation:=xlNone, SkipBlanks:=False, _
        Transpose:=True

End Sub

Sub ValuesToScratch()
Dim t As Range
    
    Call ScreenOff
    
    Call MakeScratchSheet(True)
    Set t = SelectValuesInColumnRange()
    t.Copy
    t.EntireRow.Copy
    Sheets("Scratch").Activate
    Cells(NextRow, 1).PasteSpecial Paste:=xlValues
    
    Call ScreenOn
    
End Sub

Sub ValuesToScratch2()
 useValue = active4cell.Value
 useCol = ActiveCell.Column
 
    Sheets.Add
 
 botRow = ColumnLastRow(useCol)
 
End Sub

Sub CopyRowsValue_del()

    SHFrom = ActiveSheet.Name
    useCol = ActiveCell.Column
    nRows = ColumnLastRow(useCol)
    
    Set NewSheet = Worksheets.Add
    SHTo = NewSheet.Name
    
    Worksheets(SHFrom).Activate
    useValue = ActiveCell.Value
    outCount = 0 ' starting output row
    '
    ' Copy header
    '
    Worksheets(SHFrom).Rows(1).Copy Destination:=Worksheets(SHTo).Cells(1, 1)
    '
    ' Copy rows
    '
    For i = 2 To nRows
        If Worksheets(SHFrom).Cells(i, useCol) = useValue Then
            outCount = outCount + 1
            Worksheets(SHFrom).Rows(i).EntireRow.Copy Destination:=Worksheets(SHTo).Cells(outCount + 1, 1)
            
        End If
    Next i
    
    Worksheets(SHTo).Activate
    Worksheets(SHTo).Name = useValue
    MsgBox "Copied " & outCount & " rows."
        
End Sub

Sub ValueToTab() ' searchValue, searchRange, Optional useLookAt) As Range
Dim fRange As Range
Dim useCol As Integer, useColor As Long
Dim botRow As Long

    Call ScreenOff
    
    SHOrig = ActiveSheet.Name
    WBOrig = ActiveWorkbook.Name
    
    useCol = ActiveCell.Column
    botRow = ColumnLastRow(useCol)
    Set sRange = Range(Cells(2, useCol), Cells(botRow, useCol))
    
    useValue = ActiveCell.Text
    useColor = ActiveCell.Interior.color
    Set fRange = FindInRangeExact(useValue, sRange)
    
    t = fRange.count
    
    Sheets.Add.Name = LegalSheetName(useValue)
    ActiveSheet.Tab.color = useColor
    
    Workbooks(WBOrig).Worksheets(SHOrig).Rows(1).Copy Destination:=Cells(1, 1)
    fRange.EntireRow.Copy Destination:=Cells(2, 1)
    
    Set fRange = Nothing

    Call ScreenOn
End Sub
