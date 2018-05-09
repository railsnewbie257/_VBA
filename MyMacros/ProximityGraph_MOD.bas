Attribute VB_Name = "ProximityGraph_MOD"
Sub ProximityGraph()
    
    WBOrig = ActiveWorkbook.Name
    SHOrig = ActiveSheet.Name

    useCol = FindColumnHeader("pos_city_name", SHOrig)
    Call CopyColumnToScratch(SHOrig, useCol)
    
    useCol = FindColumnHeader("pos_latitude", SHOrig)
    Call CopyColumnToScratch(SHOrig, useCol)
    
    useCol = FindColumnHeader("pos_longitude", SHOrig)
    Call CopyColumnToScratch(SHOrig, useCol)
    
    botRow = LastRow()
    For i = 2 To botRow
        y = format(Cells(i, 2), "0")
        If Not y = 0 Then
            y = y - 20
            x = format(Cells(i, 3), "0")
            If Not x = 0 Then
                x = x + 108
                Cells(y, x) = Cells(y, x) + 1
            End If
        End If
    Next i

End Sub
