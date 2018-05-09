Attribute VB_Name = "Mailer_MOD"
'
' Send the ActiveSheet in an email using Outlook as an attachment
'
Sub EmailCurrentSheet()
'Working in Excel 2000-2016
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim Sourcewb As Workbook
    Dim Destwb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim OutApp As Object
    Dim OutMail As Object
    Dim sendTo As String   ' sending to emails
    Dim subjectLine As String
    Dim emailMsg As String

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set Sourcewb = ActiveWorkbook

    'Copy the ActiveSheet to a new workbook
    ActiveSheet.Copy
    Set Destwb = ActiveWorkbook

    'Determine the Excel version and file extension/format
    With Destwb
        If val(Application.Version) < 12 Then
            'You use Excel 97-2003
            FileExtStr = ".xls": FileFormatNum = -4143
        Else
            'You use Excel 2007-2016
            Select Case Sourcewb.FileFormat
            Case 51: FileExtStr = ".xlsx": FileFormatNum = 51
            Case 52:
                If .HasVBProject Then
                    FileExtStr = ".xlsm": FileFormatNum = 52
                Else
                    FileExtStr = ".xlsx": FileFormatNum = 51
                End If
            Case 56: FileExtStr = ".xls": FileFormatNum = 56
            Case Else: FileExtStr = ".xlsb": FileFormatNum = 50
            End Select
        End If
    End With

    '    'Change all cells in the worksheet to values if you want
    '    With Destwb.Sheets(1).UsedRange
    '        .Cells.Copy
    '        .Cells.PasteSpecial xlPasteValues
    '        .Cells(1).Select
    '    End With
    '    Application.CutCopyMode = False

    'Save the new workbook/Mail it/Delete it
    TempFilePath = Environ$("temp") & "\"
    TempFileName = "Part of " & Sourcewb.Name & " " & format(Now, "dd-mmm-yy h-mm-ss")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    sendTo = InputBox("Email to: ", Title:="MailCurrentSheet", Default:=LCase(Environ$("Username")) & "@oge.com")
    subjectLine = InputBox("Subject: ", Title:="MailCurrentSheet", Default:="Email from " & LCase(Environ$("Username")))
    
    emailMsg = InputBox("Email to: ", Title:="MailCurrentSheet")

    With Destwb
        .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
        On Error Resume Next
        With OutMail
            .to = sendTo
            .CC = ""
            .BCC = ""
            .Subject = subjectLine
            .Body = emailMsg
            .Attachments.Add Destwb.FullName
            'You can add other files also like this
            '.Attachments.Add ("C:\test.txt")
            .Send   'or use .Display
        End With
        On Error GoTo 0
        .Close savechanges:=False
    End With

    'Delete the file you have send
    Kill TempFilePath & TempFileName & FileExtStr

    Set OutMail = Nothing
    Set OutApp = Nothing

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub
