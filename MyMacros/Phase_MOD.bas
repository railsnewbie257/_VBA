Attribute VB_Name = "Phase_MOD"
Sub PhaseText()
Dim textCol As Long, botRow As Long
Dim i As Long
Dim descriptCol As Integer

    textCol = FindColumnHeader("event_text")
    descriptCol = ColumnInsertLeft(textCol)
    Cells(1, descriptCol) = "PhaseDescription"
    textCol = FindColumnHeader("event_text")
    botRow = ColumnLastRow(textCol)
    For i = 2 To botRow
        k = InStr(Cells(i, textCol), "Phase")
        Cells(i, descriptCol) = Mid(Cells(i, textCol), k)
    Next i
End Sub
