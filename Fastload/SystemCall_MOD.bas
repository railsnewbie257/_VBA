Attribute VB_Name = "SystemCall_MOD"
Sub SystemCall()
Dim s As String

    s = ShellRun("cmd.exe /c cd C:\oge\fastload && fastload < ssn_test.fl")
    Debug_Print s
    
    
    startCursor = InStr(1, s, "Total Records Read") - 5
    endCursor = InStr(startCursor, s, "****")

    t = Mid(s, startCursor, endCursor - startCursor)
    
    w2 = InStr(s, "Highest return code encountered")
    startLine = InStrRev(s, ".", w2)
    endLine = InStr(startLine + 1, s, ".")
    t2 = Mid(s, startLine, endLine - startLine)
    
    If InStr(t2, "'0'") > 0 Then t = t & vbNewLine & vbNewLine & "SUCCESS"
    Debug_Print "~" & t & "~"
End Sub

Public Function ShellRun2(sCmd As String) As String

    'Run a shell command, returning the output as a string'

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    'run command'
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec(sCmd)
    Set oOutput = oExec.StdOut

    'handle the results as they are written to and read from the StdOut object'
    Dim s As String
    Dim sLine As String
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then s = s & sLine & vbCrLf
    Wend

    ShellRun = s

End Function
