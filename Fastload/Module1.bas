Attribute VB_Name = "Module1"
Private Sub Workbook_Open()
Dim fullDate As String, errorText As String
Dim usePath As String  ' location of deployed version
    
    '
    ' Check if this is the deployed version
    '
10    If ThisWorkbook.Name = MACROWORKBOOK Then  ' for daily use, check NOT timestamped version
20        'Workbooks.Add
30        Call UsageTracker("Workbook_Open", "Using Version: " & MacroTimestamp())
          Workbooks(MACROWORKBOOK).Sheets("Pallette").Activate
40        Exit Sub
50    Else
60        On Error Resume Next
70        Workbooks(MACROWORKBOOK).Close False  ' close if a My_Macros version is running to avoid conflict
80    End If
    '
    ' Fulldate for tagging
    '
90    fullDate = MacroTimestamp()
    '
100    On Error Resume Next
110        usePath = ThisWorkbook.Path
120        fromMacro = usePath & "\My_Macros_" & fullDate & ".xlsm"  ' timestamped versions
130        fromUI = usePath & "\Excel.officeUI_" & fullDate
    '
    ' check if Macro file is there
    '
140    errorText = ""
150    If ThisWorkbook.Name <> "My_Macros_" & fullDate & ".xlsm" Then errorText = "Bad Excel Macro Workbook: " & ThisWorkbook.Name
160    If Dir(fromUI) = "" Then errorText = errorText & vbNewLine & "Missing UI File: " & fromUI
    
170    If errorText <> "" Then
180        MsgBox errorText
185        Call UsageTracker("Workbook_Open", "ERROR: " & errorText)
190        ThisWorkbook.Close False
200        Exit Sub
210    End If
        
    '
    '-------------  Deploy Macros  ---------------------------------------------------------------------------
    '
220    Set fso = CreateObject("Scripting.Filesystemobject")
    '
    ' Deployed My_Macros and UI file will have timestamp in the filenames
    '
    ' check if UI file is there
    '
    ' Copy My_Macros to C:\OGE
    '
230    toMacro = "C:\OGE\My_Macros.xlsm"
240    On Error GoTo DeployErr
250        Call fso.CopyFile(fromMacro, toMacro) ' make a local copy of Macros

    '
    ' Deploy UI
    '
260    userName = LCase(Environ$("Username"))
270    toUI = "C:\Users\" & userName & "\AppData\Local\Microsoft\Office\Excel.officeUI"
280    Call fso.CopyFile(fromUI, toUI)
    
290    Workbooks.Open toMacro
300    MsgBox "New Macro Deploy Finished."
310    Call UsageTracker("Workbook_Open", "Deploying Version: " & MacroTimestamp())
        
'320    Set fso = Nothing
    '
    '
330    Set fso = Nothing
    
340    ThisWorkbook.Close False
350    Exit Sub
    
DeployErr:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="Workbook_Open"
    Stop
    Resume Next
End Sub


