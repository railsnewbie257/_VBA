Attribute VB_Name = "File_MOD"
Function LatestFile(Optional usePath)
Dim objFSO As Object
Dim objFolder As Object
Dim objFile As Object
Dim i As Long
Dim oldfilename As String
Dim old_dt As String

'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder object
Set objFolder = objFSO.GetFolder(usePath)
i = 1
'loops through each file in the directory and prints their names and path
For Each objFile In objFolder.Files
    dt = format(FileDateTime(objFile.Path), "YYYYMMDDhhmmss")  ' get the midification date
    If dt > old_dt Then
        old_dt = dt
        old_filename = objFile.Name
    End If
Next objFile
Set objFSO = Nothing
Set objFolder = Nothing
LatestFile = old_filename
End Function

Sub LoadFile()
Dim txtFilename As Variant
Dim t As String

    LoadFileDirectory.Show
    If formCancel Then Exit Sub
    
    lastFile = LatestFile(GLBFilePath)
    t = InStr(1, GLBFilePath, "SSN", vbTextCompare)
    'If InStr(1, GLBFilePath, "SSN", vbTextCompare) > 0 Then
    '    lastFile = "SSN" & lastFile
    'End If

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .AllowMultiSelect = False
        .Title = "Please Select File To Load"
        .Filters.Clear
        .Filters.Add "Excel", "*.xls?"
        .InitialFileName = GLBFilePath & lastFile
        
        If .Show = False Then Exit Sub
        
        useFile = .SelectedItems(1)

        If GLBOpenReadOnly Then
            Workbooks.Open useFile, ReadOnly:=True
            Debug.Print "readonly"
        Else
            Workbooks.Open useFile
        End If
    End With

    Set fd = Nothing
End Sub

Sub SaveFile(Optional useName)
    WBOrig = ActiveWorkbook.Name
    SHOrig = ActiveSheet.Name
    Dim intChoice As Integer
    Dim strPath As String
    
    If IsMissing(useName) Then
        WBName = ""
        If (left(WBOrig, 4) = "Book") Then ' not a Last Gasp workbook
            If SheetExists("LastGasp") Then
                GLBFilePath = LASTGASPPATH
                useCol = HeaderToColumnNum("RunDate", "LastGasp")
                If useCol <> -1 Then ' not on a Last Gasp sheet
                    eventDate = format(Worksheets("LastGasp").Cells(2, useCol), "mmddyy")
                    WBName = eventDate & ".xlsx"
                End If
            ElseIf SheetExists("ZeroKWH") Then
                GLBFilePath = ZEROKWHPATH
                WBName = format(Now(), "mmddyy") & ".xlsx"
            ElseIf SheetExists("KV2CUnderVoltage") Then
                GLBFilePath = KV2CUNDERVOLTAGEPATH
                WBName = format(Now(), "mmddyy") & ".xlsx"
            ElseIf SheetExists("UsageDrop") Then
                GLBFilePath = USAGEDROPPATH
                WBName = format(Now(), "mmddyy") & ".xlsx"
            End If
            usePath = GLBFilePath & WBName
            GLBSaveFilename = WBName
        Else
            usePath = LASTGASPPATH & WBOrig
            GLBSaveFilename = WBOrig
        End If
    Else
        usePath = useName
    End If

    SaveFileAs.txtFilename = usePath
    SaveFileAs.Show
    
    If Not formCancel Then
        On Error Resume Next
        ActiveWorkbook.SaveAs fileName:=GLBSaveFilename
        On Error GoTo 0
        Call StatusbarDisplay("Saved File")
    Else
        Call StatusbarDisplay("Cancelled Save File")
    End If
    
End Sub

Sub SplitLgDump()
Dim t As String

    WBUse = ActiveWorkbook.Name

    useCol = FindColumnHeader("event_time")
    If useCol < 0 Then
        MsgBox "Not an LG Dump file format"
        Exit Sub
    End If
    botCol = ColumnLastRow(useCol)
    StartDate = left(Cells(2, useCol), 10)
    startTime = Mid(Cells(2, useCol), 12, 2)
    i = 2
    While left(Cells(i, useCol), 10) = StartDate
        i = i + 1
    Wend
    endTime = Mid(Cells(i - 1, useCol), 12, 2)
    
    Set copyRange = Range(Rows(1), Rows(i - 1))
    
    Workbooks.Add
    copyRange.Copy Destination:=Cells(1, 1)
    
    fileName = "C:\oge\Last Gasp - " & StartDate & " " & startTime & " " & endTime & " part.xlsx"
    Call SaveLastGaspFile(fileName)
    
    '=====================================================================
    startRow = i
    StartDate = left(Cells(startRow, useCol), 10)
    If StartDate = "" Then Exit Sub
    startTime = Mid(Cells(startRow, useCol), 12, 2)
    While left(Cells(i, useCol), 10) = StartDate
        i = i + 1
    Wend
    endTime = Mid(Cells(i - 1, useCol), 12, 2)
    
    Set copyRange = Range(Rows(startRow), Rows(i - 1))
    
    Workbooks.Add
    copyRange.Copy Destination:=Cells(1, 1)
    
    fileName = "C:\oge\Last Gasp - " & StartDate & " " & startTime & " " & endTime & " part.xlsx"
    Call SaveLastGaspFile(fileName)
    
    If (i > botRow) Then Exit Sub
    
    t = left(Cells(i, useCol), 10)
    Debug.Print t
    Debug.Print i

End Sub

Sub SaveAndCloseWorkbook(WBName, filePath)
        Workbooks(WBName).Activate
        Workbooks(WBName).SaveAs fileName:=filePath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        ActiveWorkbook.Close
End Sub

