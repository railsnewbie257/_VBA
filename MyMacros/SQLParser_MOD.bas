Attribute VB_Name = "SQLParser_MOD"
Sub SQLParser(s)

    sLen = Len(s)
    t_start = 0
    t_len = 0
    t = ""
    useRow = 1
    For i = 1 To sLen
        t = Mid(s, i, 1)
        If t <> " " Then
            If t_start = 0 Then t_start = i
            t_len = t_len + 1
        End If
            
        If t = " " Or t = "(" Or t = ")" Then
            token = Mid(s, t_start, t_len)
            If IsReservedWord(token) Then
                useCol = 1
            Else
                useCol = 2
            End If
            Cells(useRow, useCol) = token
            useRow = useRow + 1
            t_start = 0
            t_len = 0
            Debug.Print token
        End If
        
    Next i
    token = Mid(s, t_start, t_len)
    Debug.Print token
            
End Sub

Sub test_SQLParser()
    Call SQLParser("SELECT UNIQUE(RunDate) FROM dl_oge_analytics." & TD_LASTGASP & " ORDER BY RunDate DESC;")
End Sub
