Attribute VB_Name = "Markie_MOD"
'Option Explicit

Function FullAddressPath(inRange As Range)

    addrName = inRange.Address
    SHName = inRange.Parent.Name
    WBName = inRange.Parent.Parent.Name
    FullAddressPath = "'[" & WBName & "]" & SHName & "'!" & addrName
End Function

Function WorkbookSheetPath(inRange) As String ' refactor ?
Dim addrName As String
Dim SHName As String, WBName As String

    addrName = inRange.Address
    SHName = inRange.Parent.Name
    WBName = inRange.Parent.Parent.Name
    WorkbookSheetPath = "'[" & WBName & "]" & SHName & "'"
End Function

Sub CopyLookupValues(Optional indexRange, Optional fromRange, Optional targetRange)
Dim t2 As String
Dim t As String

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
    
    t = WorkbookSheetPath(fromRange)
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
'
' event_external_id is from Vily's files
'
Sub MarkieFromIDs(Optional useCol, Optional SHUse, Optional WBUse)
Dim anchorRange As Range ' the column to anchor to
Dim WBAnchor As String ' workbook name for Anchor
Dim SHAnchor As String ' sheet name for Anchor
Dim anchorCol As Integer ' the anchor column
Dim anchorBotRow As Long ' last row in the Anchor column
Dim initVal As String
Dim indexCol As Long
    
    StatusbarDisplay ("SetAnchors")
    
    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    
    If IsMissing(useCol) Then
        On Error Resume Next
            Set anchorRange = Nothing
            initVal = ActiveCell.Address
            Set anchorRange = Application.InputBox("Select From ( KEY ) to Anchor.", Default:=Selection.Address, Title:="Markie From IDs", Type:=8)
            Err.Clear
            On Error GoTo 0
            If anchorRange Is Nothing Then Exit Sub
        WBAnchor = anchorRange.Parent.Parent.Name
        SHAnchor = anchorRange.Parent.Name
        anchorCol = anchorRange.Column
    End If
    
    Workbooks(WBAnchor).Worksheets(SHAnchor).Cells(1, anchorCol).Interior.color = YELLOW
    Workbooks(WBAnchor).Worksheets(SHAnchor).Cells(1, anchorCol).Font.Bold = True
    Workbooks(WBAnchor).Worksheets(SHAnchor).Tab.color = GREEN
    anchorBotRow = ColumnLastRow(anchorCol, SHAnchor, WBAnchor)
    Columns(anchorCol).AutoFit
    indexCol = AddRowNumbers(anchorCol, SHAnchor, WBAnchor)
    Workbooks(WBAnchor).Worksheets(SHAnchor).Cells(1, indexCol) = Cells(1, anchorCol).Text & "-FromIndex"
    
    Application.StatusBar = "SetAnchors: Done"
    Application.DisplayStatusBar = False
End Sub
'
' First run SetAnchors, Copies the Anchor RowIndex over
'
Sub MarkieToIDs(Optional fromRowIndexRange, Optional toRowIndexRange, Optional fromColRange, Optional toColRange)
Dim WBFrom As String
Dim SHFrom As String
Dim fromRowIdCol As Long
Dim fromHeader As String
Dim fromIndexCol As Long
Dim botRowIdFrom As Long
Dim fromLookBackCol As String
Dim fromKeyCol As Integer

Dim WBTo As String
Dim SHTo As String
Dim toRowIdcol As Long
Dim toRowIdAddr As String
Dim toRowIndexCol As Long
Dim botRowTo As Long

Dim lookupRange As Range
Dim lookupOffset As Integer
Dim fff As String
Dim t As String
Dim formulaRange As Range
Dim initVal As String
Dim fromColHeader As String


    If IsMissing(fromRowIndexRange) Then ' where the row ids are coming from
step1:
        On Error Resume Next
            Set fromRowIndexRange = Nothing
            initVal = ActiveCell.Address
            Set fromRowIndexRange = Application.InputBox("Select GREEN Anchor KEY Column", Title:="MarkieToIDs", Default:=initVal, Type:=8)
            On Error GoTo 0
        If fromRowIndexRange Is Nothing Then Exit Sub
    End If

    Set fromRowIndexRange = fromRowIndexRange.End(xlUp)
    WBFrom = fromRowIndexRange.Parent.Parent.Name
    SHFrom = fromRowIndexRange.Parent.Name
    '
    colHeader = fromRowIndexRange.End(xlUp).Value
    'If (InStr(1, colHeader, "-FromIndex") = 0) Then colHeader = Cells(1, FindColumnHeader(colHeader & "-FromIndex", SHFrom, WBFrom))
    If InStr(1, colHeader, "-FromIndex") > 0 Then
        fromRowIndexCol = FindColumnHeader(colHeader, SHFrom, WBFrom)
        keyHeader = left(colHeader, InStr(1, colHeader, "-FromIndex") - 1)
        fromKeyCol = FindColumnHeader(keyHeader, SHFrom, WBFrom)
    Else
        'MsgBox "Can not find -FromIndex for: " & colHeader, Title:="MarkieDest"
        MsgBox "Select GREEN column"
        GoTo step1
    End If
    
    
    If IsMissing(toRowIndexRange) Then ' where the row ids are going
        On Error Resume Next
            Set toRowIndexRange = Nothing
            Set toRowIndexRange = Application.InputBox("Destination KEY Column (corresponding Yellow)", Type:=8)
            Err.Clear
            On Error GoTo 0
        If toRowIndexRange Is Nothing Then Exit Sub
    End If
    '
    '----------------------------------------------------------------------------------------------------------
    '
    WBTo = toRowIndexRange.Parent.Parent.Name
    SHTo = toRowIndexRange.Parent.Name
    toRowIndexCol = toRowIndexRange.Column
    toRowIndexCol = ColumnInsertRight(toRowIndexCol, SHTo, WBTo) ' needs to do the insert incase moves the FROM column on same sheet
    
    '
    ' Use the column header to get the rownums
    ''fromHeader = Workbooks(WBFrom).Worksheets(SHFrom).Cells(1, fromRowIdCol) ' name of the From column
    ''fromIndexCol = FindColumnHeader(fromHeader, SHFrom, WBFrom)
    ''If (fromIndexCol < 0) Then
    ''    MsgBox "Can not find RowIndex for: " & fromHeader, Title:="MarkieDest"
    ''    Exit Sub
    ''End If
    botRowIndexCol = ColumnLastRow(fromRowIndexRange.Column, SHFrom, WBFrom)
    Set lookupRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(2, fromKeyCol), _
                            Workbooks(WBFrom).Worksheets(SHFrom).Cells(botRowIndexCol, fromRowIndexRange.Column))
    Debug.Print lookupRange.Address(True, True, xlA1, True)
    
    toRowIndexAddr = Workbooks(WBTo).Worksheets(SHTo).Cells(2, toRowIndexRange.Column).Address(False, False)
   
    fff = "=VLOOKUP(" & toRowIndexAddr & "," & lookupRange.Address(True, True, xlA1, True) & _
                                "," & fromRowIndexRange.Column - fromKeyCol + 1 & ",0)"

    'toRowIndexCol = ColumnInsertRight(toRowIdcol, SHTo, WBTo)
    fromLookBackCol = toRowIndexCol  ' column with index IDs from WBTo
    botRowToIndexCol = ColumnLastRow(toRowIndexRange.Column, SHTo, WBTo)
    

    Workbooks(WBTo).Worksheets(SHTo).Cells(1, toRowIndexCol) = "[" & WBFrom & "]" & SHFrom & "-ToIndex"
    Call ColorRange(Workbooks(WBTo).Worksheets(SHTo).Cells(1, toRowIndexCol), ORANGE)
    Workbooks(WBTo).Worksheets(SHTo).Tab.color = ORANGE
    
    Set formulaRange = Range(Workbooks(WBTo).Worksheets(SHTo).Cells(2, toRowIndexCol), _
                             Workbooks(WBTo).Worksheets(SHTo).Cells(botRowToIndexCol, toRowIndexCol))
    formulaRange.NumberFormat = "General"
    formulaRange.Formula = fff
    t = fromRowIndexRange.Address(False, False, xlA1)
    Call RangeToValues(formulaRange)
    formulaRange.NumberFormat = "0"
    
End Sub

Sub MarkieCopy_old(Optional indexRange, Optional fromColRange, Optional toColRange)
Dim t As String
Dim fff As String
Dim k As Long
Dim i As Long
Dim j As Long
Dim s As String
Dim indexCol As String
Dim useCol As Long
Dim fromLetCol As String ' used to the the column letter for copying
Dim botRow As Long
Dim columnName As String
Dim toWBRef As String
Dim WBTo As String
Dim SHTo As String
Dim toCol As Integer
Dim WBFrom As String
Dim SHFrom As String
Dim fromWBRef As String
Dim fRange As Range
Dim top As Integer
Dim t_left As Integer
Dim rightCol As Integer
Dim retCode As Integer
Dim fromIndexHeader As String
Dim toLookBackCol As Integer ' the column with row() indexes from the Source sheet
Dim toLookBackHeader As String
Dim cRange As Range
    
    On Error Resume Next
    If IsMissing(fromColRange) Then
        Set fromColRange = Nothing
        Set fromColRange = Application.InputBox("Select Source Column(s)", Default:=Selection.Address, Title:="AnchorGetColumn", Type:=8)
        If fromColRange Is Nothing Then Exit Sub
    End If
    WBFrom = fromColRange.Parent.Parent.Name
    SHFrom = fromColRange.Parent.Name
    fromWBRef = "[" & WBFrom & "]" & SHFrom

    If IsMissing(toColRange) Then
        Set toColRange = Nothing
        Set toColRange = Application.InputBox("Select Destination Column", Title:="AnchorGetColumn", Type:=8)
        If toColRange Is Nothing Then Exit Sub
    End If
    SHTo = toColRange.Parent.Name
    WBTo = toColRange.Parent.Parent.Name
    toCol = toColRange.Column
    '
    ' Search for Look Back Index
    '
    rightCol = RowLastColumn(1, SHTo, WBTo)
    toLookBackHeader = fromWBRef & "-RowIndex"
    toLookBackCol = FindColumnHeader(toLookBackHeader, SHTo, WBTo) ' column with row index to Source sheet
    If toLookBackCol < 0 Then
        MsgBox "Source Spreadsheet has no Index, EXITING"
        Exit Sub
    End If
    '
    ' To loop ---------------------------------------------------------------------------------------------
    '
    
    For i = 1 To fromColRange.Areas.count
        For j = 1 To fromColRange.Areas(i).Columns.count
            't = fromColRange.End(xlUp).Address(False, False, xlA1) ' top row so address row number is "1"
            s = fromColRange.Areas(i).Columns(j).End(xlUp).Address(False, False, xlA1)
            fromLetCol = left(s, Len(s) - 1)
            '
            ' Check if Destination Column is blank
            '
            botRow = ColumnLastRow(toCol, SHTo, WBTo)
            If botRow > 0 Then
                useCol = ColumnInsertRight(toCol, SHTo, WBTo)
            Else
                useCol = toCol
            End If
            'useCol2 = indexRange.Column
    
            toLookBackHeader = fromWBRef & "-RowIndex"
            toLookBackCol = FindColumnHeader(toLookBackHeader, SHTo, WBTo) ' column with row index to Source sheet
            If toLookBackCol < 0 Then
                MsgBox "Source Spreadsheet has no Index, EXITING"
                Exit Sub
            End If
    
            botRow = ColumnLastRow(toLookBackCol, SHTo, WBTo)
    
            Set fRange = Range(Workbooks(WBTo).Worksheets(SHTo).Cells(2, useCol), _
                               Workbooks(WBTo).Worksheets(SHTo).Cells(botRow, useCol))

            fff = "='" & fromWBRef & "'!" & fromLetCol & Workbooks(WBTo).Worksheets(SHTo).Cells(2, Cells(2, toLookBackCol).Column)
            Debug.Print fff
    
            fRange.Copy
            fRange.Formula = fff
            Call RangeToValues(fRange)
            Workbooks(WBTo).Worksheets(SHTo).Cells(1, useCol).Formula = "='" & fromWBRef & "'!" & fromLetCol & "1"
            Call ColorRange(Workbooks(WBTo).Worksheets(SHTo).Cells(1, useCol), LIGHTBLUE)
            Workbooks(WBTo).Worksheets(SHTo).Columns(useCol).AutoFit
    
            toCol = toCol + 1 ' iterate to the right
        Next j
    Next i
End Sub

Sub MarkieCopy(Optional indexRange, Optional fromColRange, Optional toColRange)
Dim t As String
Dim fff As String
Dim k As Long
Dim i As Long
Dim j As Long
Dim s As String
Dim indexCol As String
Dim useCol As Long
Dim fromLetCol As String ' used to the the column letter for copying
Dim botRow As Long
Dim columnName As String, lookBackColLet As String
Dim toWBRef As String
Dim WBTo As String, SHTo As String
Dim toCol As Integer
Dim WBFrom As String, SHFrom As String
Dim fromWBRef As String, fromColLet As String
Dim fRange As Range
Dim top As Integer
Dim t_left As Integer
Dim rightCol As Integer
Dim retCode As Integer
Dim fromIndexHeader As String
Dim toLookBackCol As Integer ' the column with row() indexes from the Source sheet
Dim toLookBackHeader As String
Dim cRange As Range
    
    On Error Resume Next
    If IsMissing(fromColRange) Then
        Set fromColRange = Nothing
        Set fromColRange = Application.InputBox("Select Source DATA Column(s) (on GREEN column sheet)" & vbNewLine, Default:=Selection.Address, Title:="AnchorGetColumn", Type:=8)
        If fromColRange Is Nothing Then Exit Sub
    End If
    WBFrom = fromColRange.Parent.Parent.Name
    SHFrom = fromColRange.Parent.Name
    fromWBRef = "[" & WBFrom & "]" & SHFrom
    fromCol = fromColRange.Column

    If IsMissing(toColRange) Then
        Set toColRange = Nothing
        Set toColRange = Application.InputBox("Select Destination Column (on ORANGE column sheet)  Where to Put It?", Title:="AnchorGetColumn", Type:=8)
        If toColRange Is Nothing Then Exit Sub
    End If
    SHTo = toColRange.Parent.Name
    WBTo = toColRange.Parent.Parent.Name
    toCol = ColumnInsertLeft(toColRange.Column, SHTo, WBTo)
    '
    ' Search for Look Back Index
    '
    rightCol = RowLastColumn(1, SHTo, WBTo)
    toLookBackHeader = fromWBRef & "-ToIndex"
    toLookBackCol = FindColumnHeader(toLookBackHeader, SHTo, WBTo) ' column with row index to Source sheet
    lookBackRange = Cells
    If toLookBackCol < 0 Then
        MsgBox "Source Spreadsheet has no Index, EXITING"
        Exit Sub
    End If
    '
    ' To loop ---------------------------------------------------------------------------------------------
    '
    
    For i = 1 To fromColRange.Areas.count
        For j = 1 To fromColRange.Areas(i).Columns.count
            't = fromColRange.End(xlUp).Address(False, False, xlA1) ' top row so address row number is "1"
            s = fromColRange.Areas(i).Columns(j).End(xlUp).Address(False, False, xlA1)
            fromColLet = left(s, Len(s) - 1)
            '
            ' Check if Destination Column is blank
            '
            'botRow = ColumnLastRow(toCol, SHTo, WBTo)
            'If botRow > 0 Then
            '    useCol = ColumnInsertRight(toCol, SHTo, WBTo)
            'Else
            '    useCol = toCol
            'End If
            'useCol2 = indexRange.Column
    
            toLookBackHeader = fromWBRef & "-ToIndex"
            toLookBackCol = FindColumnHeader(toLookBackHeader, SHTo, WBTo) ' column with row index to Source sheet
            If toLookBackCol < 0 Then
                MsgBox "Destination Spreadsheet has no Index, EXITING"
                Exit Sub
            End If
    
            botRow = ColumnLastRow(toLookBackCol, SHTo, WBTo)
    
            Set fRange = Range(Workbooks(WBTo).Worksheets(SHTo).Cells(2, toCol), _
                               Workbooks(WBTo).Worksheets(SHTo).Cells(botRow, toCol))

            t = Cells(1, toLookBackCol).Address(False, False, xlA1)
            lookBackColLet = left(t, Len(t) - 1)
            '
            ' =INDIRECT("'[customers.xlsx]Oil Field'!C" & D2)
            '
            fff = "=indirect(" & """'" & fromWBRef & "'!" & fromColLet & """" & "&" & lookBackColLet & "2)"
            Debug.Print fff
            fRange.Formula = fff
            fRange.NumberFormat = "General"
            Call RangeToValues(fRange)
            Workbooks(WBTo).Worksheets(SHTo).Cells(1, toCol).Formula = "='" & fromWBRef & "'!" & fromColLet & "1"
            Call ColorRange(Workbooks(WBTo).Worksheets(SHTo).Cells(1, toCol), LIGHTBLUE)
            Workbooks(WBTo).Worksheets(SHTo).Columns(useCol).AutoFit
    
            toCol = toCol + 1 ' iterate to the right
        Next j
    Next i
    Workbooks(WBTo).Worksheets(SHTo).Activate
End Sub



