Attribute VB_Name = "PowerOn_MOD"
Sub EventTimeHilite()
Dim first12007 As Long, last12007 As Long, last15036 As Long, last15035 As Long, last10007 As Long
Dim SHUse As String, WBUse As String
Dim i As Long

    WBUse = ActiveWorkbook.Name
    SHUse = ActiveSheet.Name
    '
    ' Get the columns
    '
    first12007 = FindColumnHeader("First_Event_Time_12007")
    last12007 = FindColumnHeader("Last_Event_Time_12007")
    last15036 = FindColumnHeader("Last_Event_Time_15036")
    last15035 = FindColumnHeader("Last_Event_Time_15035")
    last10007 = FindColumnHeader("Last_Event_Time_100007")
    '
    ' Format the column
    '
    Columns(first12007).NumberFormat = "h:mm;@"
    Columns(last12007).NumberFormat = "h:mm;@"
    Columns(last15036).NumberFormat = "h:mm;@"
    Columns(last15035).NumberFormat = "h:mm;@"
    Columns(last10007).NumberFormat = "h:mm;@"
    
    
    botRow = LastRow(SHUse)
    For i = 2 To botRow
        If Cells(i, last12007) < Cells(i, last15036) Or _
            Cells(i, last12007) < Cells(i, last15036) Or _
            Cells(i, last12007) < Cells(i, last15035) Or _
            Cells(i, last12007) < Cells(i, last10007) Then
                useColor = LIGHTBLUE
        Else
                useColor = ORANGE
        End If
        
        Call ColorRange(Cells(i, first12007), useColor)
        Call ColorRange(Cells(i, last12007), useColor)
        Call ColorRange(Cells(i, last15035), useColor)
        Call ColorRange(Cells(i, last15036), useColor)
        Call ColorRange(Cells(i, last10007), useColor)
    Next i
End Sub
