Attribute VB_Name = "SSN_MOD"
Sub SSNDirectory()
Dim t As String
Dim useRow As Long

    Workbooks.Add
    WBNew = ActiveWorkbook.Name
    SHNew = "Directory"

    Call DeleteSheet(SHNew)

    Sheets.Add
    SHNew = "Directory"
    ActiveSheet.Name = SHNew
    
    filePath = SSNPATH & "*"
    t = Dir(filePath)
    useRow = 0
    Do Until t = ""
        useRow = useRow + 1
        Cells(useRow, 1) = t
        If Not (left(t, 4) = "SSN-") Then
            MsgBox "Non SSN file detected " & t & vbNewLine & vbNewLine & "EXITING", Title:="SSNDirectory"
            Exit Do
        End If
        
        t = Dir()
    Loop
    
    Call SortSheetUp(1)
End Sub

Sub SsnMerge()
Dim i As Long, botRowDir As Long
Dim botRowFrom As Long, botRowMerge As Long
Dim fso As Object
    
    Call SSNDirectory
    WBDir = ActiveWorkbook.Name
    SHDir = ActiveSheet.Name
    
    datesMerged = False
    WBDir = ActiveWorkbook.Name
    SHDir = "Directory"
    botRowDir = ColumnLastRow(1, SHDir, WBDir)
    oldDate = "" 'Mid(Workbooks(WBDir).Worksheets(SHDir).Cells(1, 1), 5, 10)
    filesMerged = False
    i = 1
    While i <= botRowDir - 1 ' assumes the last file is the first partial for the next day
        filesMerged = False
        Debug.Print Workbooks(WBDir).Worksheets(SHDir).Cells(i, 1)
        firstDate = Mid(Workbooks(WBDir).Worksheets(SHDir).Cells(i, 1), 5, 10) ' get the date portion
        firstTime1 = Mid(Workbooks(WBDir).Worksheets(SHDir).Cells(i, 1), 16, 6) ' begin time
        firstTime2 = Mid(Workbooks(WBDir).Worksheets(SHDir).Cells(i, 1), 23, 6) ' end time
        secondDate = Mid(Workbooks(WBDir).Worksheets(SHDir).Cells(i + 1, 1), 5, 10)
        secondTime1 = Mid(Workbooks(WBDir).Worksheets(SHDir).Cells(i + 1, 1), 16, 6) ' begin time
        secondTime2 = Mid(Workbooks(WBDir).Worksheets(SHDir).Cells(i + 1, 1), 23, 6) ' end time
        '
        ' Start merging if two consecutive files with same date have not been processed
        '
        If (firstDate = secondDate) Then
            If (Len(Workbooks(WBDir).Worksheets(SHDir).Cells(i, 1)) > 20) And _
               (Len(Workbooks(WBDir).Worksheets(SHDir).Cells(i + 1, 1)) > 20) Then
                '
                ' Load first file ------------------------------------------------------------------------------------
                '
                Workbooks.Add
                WBMerge = ActiveWorkbook.Name
                SHMerge = ActiveSheet.Name
                loadedDate = Mid(Workbooks(WBDir).Worksheets(SHDir).Cells(i, 1), 5, 10)
                Workbooks.Open SSNPATH & Workbooks(WBDir).Worksheets(SHDir).Cells(i, 1)
                Debug.Print "Loading: "; Workbooks(WBDir).Worksheets(SHDir).Cells(i, 1)
                
                WBFrom = ActiveWorkbook.Name
                SHFrom = ActiveSheet.Name
                '
                ' Copy the header
                '
                Workbooks(WBFrom).Worksheets(SHFrom).Rows(1).Copy Destination:=Workbooks(WBMerge).Worksheets(SHMerge).Cells(1, 1)
                topRow = DATASTARTROW
                
                botRowFrom = ColumnLastRow(DATAFIRSTCOL, SHFrom, WBFrom)
                rightColumn = LastColumn(SHFrom, WBFrom)
                Set fromRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(topRow, 1), _
                                      Workbooks(WBFrom).Worksheets(SHFrom).Cells(botRowFrom, rightColumn))
                
                botRowMerge = ColumnLastRow(DATAFIRSTCOL, SHMerge, WBMerge) + 1
                Set mergeRange = Range(Workbooks(WBMerge).Worksheets(SHMerge).Cells(botRowMerge, 1), _
                                Workbooks(WBMerge).Worksheets(SHMerge).Cells(botRowMerge, 1))
                '
                ' Copy first contents to merge workbook ------------------------------------------------------------------
                '
                fromRange.Copy Destination:=mergeRange
            
                Workbooks(WBFrom).Close
                Call MoveSSNProcessedFile(WBFrom)
                i = i + 1
                '
                ' Loop over the next date files
                '
                While Mid(Workbooks(WBDir).Worksheets(SHDir).Cells(i, 1), 5, 10) = loadedDate And _
                        i <= botRowDir And _
                        Len(Workbooks(WBDir).Worksheets(SHDir).Cells(i, 1)) > 20
                
                    Debug.Print "Merging: " & Workbooks(WBDir).Worksheets(SHDir).Cells(i, 1)
                    Workbooks.Open SSNPATH & Workbooks(WBDir).Worksheets(SHDir).Cells(i, 1)
                    WBFrom = ActiveWorkbook.Name
                    WHFrom = ActiveSheet.Name
                    botRowFrom = ColumnLastRow(DATAFIRSTCOL, SHFrom, WBFrom)
                    rightColumn = LastColumn(SHFrom, WBFrom)
                    Set fromRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(DATASTARTROW, 1), _
                                          Workbooks(WBFrom).Worksheets(SHFrom).Cells(botRowFrom, rightColumn))
                    botRowMerge = ColumnLastRow(DATAFIRSTCOL, SHMerge, WBMerge)
                    Set mergeRange = Range(Workbooks(WBMerge).Worksheets(SHMerge).Cells(botRowMerge, DATAFIRSTCOL), _
                                           Workbooks(WBMerge).Worksheets(SHMerge).Cells(botRowMerge, DATAFIRSTCOL))
                              
                    fromRange.Copy Destination:=mergeRange
                    '
                    ' Close and Move the processed file
                    '
                    Workbooks(WBFrom).Close
                    Call MoveSSNProcessedFile(WBFrom)

                    i = i + 1
                    filesMerged = True
                Wend
                '
                ' Maybe save the merged file
                '
                If Not filesMerged Then
                    Call AlertsOff
                        Workbooks(WBMerge).Close
                        On Error Resume Next
                        Workbooks(WBFrom).Close
                    Call AlertsOn
                Else
                    fileName = SSNPATH & "SSN-" & firstDate & ".xlsx"
                    Call SaveAndCloseWorkbook(WBMerge, fileName)
                End If
            Else
                If (Len(Workbooks(WBDir).Worksheets(SHDir).Cells(i + 1, 1)) > 20) Then
                    fileName = Workbooks(WBDir).Worksheets(SHDir).Cells(i + 1, 1)
                    Call MoveSSNProcessedFile(fileName)
                    Workbooks(WBDir).Worksheets(SHDir).Rows(i + 1).Delete
                    botRowDir = botRowDir - 1
                End If
            End If
        Else
        
            fileMoved = False
            If (Len(Workbooks(WBDir).Worksheets(SHDir).Cells(i, 1)) > 20) Then
                fileName = Workbooks(WBDir).Worksheets(SHDir).Cells(i, 1)
                Call MoveSSNProcessedFile(fileName)
                Workbooks(WBDir).Worksheets(SHDir).Rows(i).Delete
                botRowDir = botRowDir - 1
                fileMoved = True
            End If
            
            If Not fileMoved Then i = i + 1
        End If
    Wend
    Call ClearClipboard
    Workbooks(WBDir).Close False
    
End Sub

Sub MoveSSNProcessedFile(fileName)
    Dim fso As Object
    Set fso = CreateObject("Scripting.Filesystemobject")
    R = Dir(SSNPATH & Cells(1, 1))
    'On Error Resume Next
    Call AlertsOff
    fso.CopyFile SSNPATH & fileName, SSNPATH & "processed SSN downloads\" & fileName, True
    fso.deleteFile SSNPATH & fileName
    Call AlertsOn
    Set fso = Nothing
End Sub

Sub SSNSplitFile()
Dim i As Long, botRow As Long, startRow As Long
Dim startTime As String
Dim endTime As String
Dim t As String

    If (IdentifySheet() <> "SSN") Then
        MsgBox "Not an SSN file format", Title:="SSNSplitFile"
        Exit Sub
    End If

    Debug.Print "Start: " & Now()
    WBFrom = ActiveWorkbook.Name
    SHFrom = ActiveSheet.Name

    sortCol = FindColumnHeader("event_time")
    Call SortSheetUp(sortCol)
    '
    ' Fist check top and bottom if only one date in this set
    '
    firstDate = left(Cells(DATASTARTROW, sortCol), 10)
    t = Mid(Workbooks(WBFrom).Worksheets(SHFrom).Cells(DATASTARTROW, sortCol), 12, 8)
    startTime = Replace(t, ":", "")
    botRow = ColumnLastRow(sortCol)
    lastDate = left(Cells(botRow, sortCol), 10)
    startRow = DATASTARTROW
    rightColumn = LastColumn(SHFrom, WBFrom)
    
    For i = DATASTARTROW To botRow
        If Not (left(Workbooks(WBFrom).Worksheets(SHFrom).Cells(i, sortCol), 10) = firstDate) Then

            Workbooks.Add ' target workbook to save this section
            WBNew = ActiveWorkbook.Name
            SHNew = ActiveSheet.Name
            
            endTime = Replace(Mid(Workbooks(WBFrom).Worksheets(SHFrom).Cells(i - 1, sortCol), 12, 8), ":", "")
            '
            ' copy to target save workbook
            '
            Workbooks(WBFrom).Worksheets(SHFrom).Rows(1).Copy Destination:=Workbooks(WBNew).Worksheets(SHNew).Rows(1)
            
            Set fromRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(startRow, DATAFIRSTCOL), _
                                  Workbooks(WBFrom).Worksheets(SHFrom).Cells(i - 1, rightColumn))
                                  
            Set toRange = Range(Workbooks(WBNew).Worksheets(SHNew).Cells(DATASTARTROW, DATAFIRSTCOL), _
                                Workbooks(WBNew).Worksheets(SHNew).Cells(DATASTARTROW, DATAFIRSTCOL))
                                
            fromRange.Copy Destination:=toRange
            Call ClearClipboard
            
            fileName = SSNPATH & "SSN-" & firstDate & "-" & startTime & "-" & endTime & ".xlsx"
            Call AlertsOff
                Call SaveAndCloseWorkbook(WBNew, fileName)
            Call AlertsOn
            Debug.Print "Finished(" & oldDate & ") " & Now()
            '
            ' reset section info
            '
            firstDate = left(Workbooks(WBFrom).Worksheets(SHFrom).Cells(i, sortCol), 10)
            startTime = Replace(Mid(Workbooks(WBFrom).Worksheets(SHFrom).Cells(i, sortCol), 12, 8), ":", "")
            startRow = i
        End If
        
    Next i
    Call ClearClipboard
    
    '
    ' Save the rest
    '
    Workbooks.Add ' target workbook to save this section
    WBNew = ActiveWorkbook.Name
    SHNew = ActiveSheet.Name
    
    Workbooks(WBFrom).Worksheets(SHFrom).Rows(1).Copy Destination:=Workbooks(WBNew).Worksheets(SHNew).Rows(1)
    
    endTime = Replace(Mid(Workbooks(WBFrom).Worksheets(SHFrom).Cells(i - 1, sortCol), 12, 8), ":", "")
    Set fromRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(startRow, DATAFIRSTCOL), _
                          Workbooks(WBFrom).Worksheets(SHFrom).Cells(i - 1, rightColumn))
                                  
    Set toRange = Range(Workbooks(WBNew).Worksheets(SHNew).Cells(DATASTARTROW, DATAFIRSTCOL), _
                        Workbooks(WBNew).Worksheets(SHNew).Cells(DATASTARTROW, DATAFIRSTCOL))
                        
    fromRange.Copy Destination:=toRange
    Call ClearClipboard
    
    fileName = SSNPATH & "SSN-" & firstDate & "-" & startTime & "-" & endTime & ".xlsx"
    Call AlertsOff
        Call SaveAndCloseWorkbook(WBNew, fileName)
    Call AlertsOn
    Call ClearClipboard
    Debug.Print "Finished(" & oldDate & ") " & Now()
    Call AlertsOff
        Workbooks(WBFrom).Close
    Call AlertsOn
    Call ClearClipboard
    
    MsgBox "Finished."
End Sub

