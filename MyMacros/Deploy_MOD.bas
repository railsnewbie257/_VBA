Attribute VB_Name = "Deploy_MOD"

Sub DoDeploy()
Dim fso As Object
Dim deployName As String
Dim serialDate As String
Dim nowTime As Date
Dim WBMacro As String
    
    '
    ' Get the UI file
    '
    userName = LCase(Environ$("Username"))
    '
    ' Save the current Marcos
    '
    ThisWorkbook.Save
    WBMacro = ThisWorkbook.Name
    '
    ' make the timestamp
    '
    deployName = format(Now(), "[$-F800]ddd, mmm dd, yyyy HH-MM-SS")
    deployName = Replace(deployName, ", ", "_")
    deployName = Replace(deployName, " ", "_")
    
    deployName = InputBox("Deploy Name:", Default:=deployName)
    deployName = Replace(deployName, " ", "_")
    
    ThisWorkbook.Worksheets("Pallette").Activate
    ThisWorkbook.Worksheets("Pallette").Cells(8, 1) = deployName
    '
    ThisWorkbook.Worksheets("Pallette").Cells(1, 1) = "" ' blank out TD password
    ThisWorkbook.Save
    '
    '
    Set fso = CreateObject("Scripting.Filesystemobject")
    '
    ' Copy the Macro to the timstamp version
    '
    fromFile = "C:\OGE\My_Macros.xlsm"
    toFile = "C:\OGE\My_Macros_" & deployName & ".xlsm"
    Call fso.CopyFile(fromFile, toFile)
    '
    ' Get the UI file
    '
    fromFile = "C:\Users\" & userName & "\AppData\Local\Microsoft\Office\Excel.officeUI"
    toFile = "C:\OGE\Excel.officeUI_" & deployName
    Call fso.CopyFile(fromFile, toFile)
    '
    ThisWorkbook.Close False
End Sub

Sub DeployUI()
Dim fso As Object
Dim userName As String
Dim fullDate As String
Dim toFile As String
Dim fromFile As String

    Set fso = CreateObject("Scripting.Filesystemobject")
    userName = LCase(Environ$("Username"))

    fullDate = Workbooks(MACROWORKBOOK).Worksheets("Pallette").Cells(8, 1)
    toFile = "C:\Users\" & userName & "\AppData\Local\Microsoft\Office\Excel.officeUI"
    fromFile = "C:\OGE\Excel.officeUI_" & fullDate
    
    If Dir(toFile) <> "" Then
        Call fso.CopyFile(toFile, toFile & "_old")
    End If

    If Dir(fromFile) <> "" Then
    Debug.Print toFile
        Call fso.CopyFile(fromFile, toFile, True)
        Call fso.deleteFile(fromFile)
        MsgBox "New Menus Deployed: " & fullDate
    End If
    Set fso = Nothing
End Sub
