Attribute VB_Name = "Singletons_MOD"
Sub MakeSingletons()
Dim i As Long, botRow As Long

    botRow = LastRow()
    
    dateCol = FindColumnHeader("RunDate")
    
    ptransformerCol = ColumnInsertRight(dateCol)
        Cells(1, ptransformerCol) = "p_transformer"
    pcircuitCol = ColumnInsertRight(dateCol)
        Cells(1, pcircuitCol) = "p_circuit"
    pzipCol = ColumnInsertRight(dateCol)
        Cells(1, pzipCol) = "p_zip"
    pcityCol = ColumnInsertRight(dateCol)
        Cells(1, pcityCol) = "p_city"
    ptimeCol = ColumnInsertRight(dateCol)
        Cells(1, ptimeCol) = "p_time"
    psumCol = ColumnInsertRight(dateCol)
        Cells(1, psumCol) = "p_sum"
    
    psumCol = FindColumnHeader("p_sum")
    ptimeCol = FindColumnHeader("p_time")
    pcityCol = FindColumnHeader("p_city")
    pzipCol = FindColumnHeader("p_zip")
    pcircuitCol = FindColumnHeader("p_circuit")
    ptransformerCol = FindColumnHeader("p_transformer")
    
    Columns(psumCol).columnWidth = "3.5"
    Columns(ptimeCol).columnWidth = "3.5"
    Columns(pcityCol).columnWidth = "3.5"
    Columns(pzipCol).columnWidth = "3.5"
    Columns(pcircuitCol).columnWidth = "3.5"
    Columns(ptransformerCol).columnWidth = "3.5"
    
    timeCol = FindColumnHeader("first_event_time")
    cityCol = FindColumnHeader("pos_city_name")
    zipCol = FindColumnHeader("proximity_zip_code")
    circuitCol = FindColumnHeader("circuit_number")
    transformerCol = FindColumnHeader("transformer_number")
    
    For i = 11 To botRow - 10
        Cells(i, pcircuitCol) = ProximityCount(circuitCol, i)
        Cells(i, ptransformerCol) = ProximityCount(transformerCol, i)
        Cells(i, pcityCol) = ProximityCount(cityCol, i)
        Cells(i, pzipCol) = ProximityCount(zipCol, i)
        Cells(i, ptimeCol) = ProximityCount(timeCol, i)
        
        If i Mod 100 = 0 Then Call StatusbarDisplay(format(i, "#,##0"))
    Next i

End Sub

Sub SingletonsHilite()
Dim i As Long, botRow As Long

    psumCol = FindColumnHeader("p_sum")
    botRow = LastRow
    
    If botRow = 1 Then Exit Sub ' only column headers
    
    Set fRange = Range(Cells(2, psumCol), Cells(2, psumCol))
    Set sumRange = Range(fRange.Offset(0, 1), fRange.Offset(0, 5))
    Set fRange = Range(Cells(2, psumCol), Cells(botRow, psumCol))
    fRange.Formula = "=SUM(" & sumRange.Address(False, False) & ")"
    Call RangeToValues(fRange)
    
    psumCol = FindColumnHeader("p_sum")
    Columns(psumCol).columnWidth = 5
    ptimeCol = FindColumnHeader("p_time")
    Columns(ptimeCol).columnWidth = 5
    pcityCol = FindColumnHeader("p_city")
    Columns(pcityCol).columnWidth = 5
    pzipCol = FindColumnHeader("p_zip")
    Columns(pzipCol).columnWidth = 5
    pcircuitCol = FindColumnHeader("p_circuit")
    Columns(pcircuitCol).columnWidth = 5
    ptransformerCol = FindColumnHeader("p_transformer")
    Columns(ptransformerCol).columnWidth = 5
    
    For i = 11 To botRow - 10
        If Cells(i, psumCol) = 5 Then
            useColor = ORANGE
        Else
            useColor = LIGHTBLUE
        End If
        Call ColorRange(Cells(i, psumCol), useColor)
        Call ColorRange(Cells(i, ptimeCol), useColor)
        Call ColorRange(Cells(i, pcityCol), useColor)
        Call ColorRange(Cells(i, pzipCol), useColor)
        Call ColorRange(Cells(i, pcircuitCol), useColor)
        Call ColorRange(Cells(i, ptransformerCol), useColor)
    Next i
    
End Sub

Sub GetSingletons()
Dim i As Long, botRow As Long

    SHActive = "A-Single"
    SHDisconnect = "D-Single"
    SHOrig = ActiveSheet.Name
    
    SHUse = MakeSheet(True, SHActive)
    SHUse = MakeSheet(True, SHDisconnect)
    
    Sheets(SHOrig).Activate
    psumCol = FindColumnHeader("p_sum")
    statusCol = FindColumnHeader("src_ops_state")
    botRow = LastRow(SHOrig)
    
    Worksheets(SHOrig).Rows(1).Copy Destination:=Worksheets(SHActive).Rows(1)
    Worksheets(SHOrig).Rows(1).Copy Destination:=Worksheets(SHDisconnect).Rows(1)
    
    
    botRowActive = 2
    botRowdisconnect = 2
    For i = 2 To botRow
        If Cells(i, psumCol) = 5 Then
            If Cells(i, statusCol) = "Active" Then
                Rows(i).EntireRow.Copy Destination:=Worksheets(SHActive).Rows(botRowActive)
                botRowActive = botRowActive + 1
            ElseIf Cells(i, statusCol) = "Disconnected" Then
                'Rows(i).EntireRow.Copy Destination:=Worksheets(SHDisconnect).Rows(botRowdisconnect)
                Rows(i).EntireRow.Copy
                Worksheets(SHDisconnect).Rows(botRowdisconnect).PasteSpecial Paste:=xlAll
                botRowdisconnect = botRowdisconnect + 1
            End If
        End If
    Next i
    
    
End Sub
