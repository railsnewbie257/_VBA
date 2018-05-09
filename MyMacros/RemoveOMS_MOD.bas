Attribute VB_Name = "RemoveOMS_MOD"
Sub RemoveOMS()
Dim sRange As range
Dim omsRange As range, woRange As range
Dim copyRange As range
Dim SHOutage As String

10    SHOrig = ActiveSheet.Name

On Error GoTo gotError
20   SHOutage = "Outage"
30   If SheetExists(SHOutage, ActiveWorkbook.Name) Then
40       Call AlertsOff
50            Worksheets(SHOms).Delete
60        Call AlertsOn
70    End If
80    Worksheets.Add.Name = SHOutage
90    Worksheets(SHOrig).Rows(1).Copy Destination:=Worksheets(SHOutage).Cells(1, 1)
100   Worksheets(SHOrig).Activate
    '
    ' Get outage meters
    '
110   omsCol = FindColumnHeader("Outage_Event_Id")
120   botRow = ColumnLastRow(omsCol)
130   Set sRange = range(Cells(2, omsCol), Cells(botRow, omsCol))

140    Set fRange = FindRangeNonEmpty(sRange)

150    fRange.EntireRow.Copy Destination:=Worksheets(SHOutage).Cells(2, 1)
160    fRange.EntireRow.Delete

170    woCol = FindColumnHeader("Work_Order_Id")
180   botRow = ColumnLastRow(woCol)
190   Set sRange = range(Cells(2, woCol), Cells(botRow, woCol))
    
200   Set fRange = FindRangeNonEmpty(sRange)

210    k = LastRow(SHOutage) + 1
220    fRange.EntireRow.Copy Destination:=Worksheets(SHOutage).Cells(k, 1)
230    fRange.EntireRow.Delete
    
240    Exit Sub
    
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="RemoveOMS"
    Stop
    Resume Next
End Sub
