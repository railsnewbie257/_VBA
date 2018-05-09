Attribute VB_Name = "VKWH_MOD"
Sub VKWH()
Dim fRange As range

10   SHVkwh = "VKWHSelect"

20       LOAD ChooseDateForm
         ChooseDateForm.caption = "Start Date"
30       ChooseDateForm.MonthView1.Value = Now()
40       ChooseDateForm.Show
50       If formCancel Then Exit Sub
        
60       StartDate = format(ChooseDateForm.MonthView1, "YYYY-MM-DD")
61       Unload ChooseDateForm

62       LOAD ChooseDateForm
63       ChooseDateForm.caption = "End Date"
64       ChooseDateForm.MonthView1.Value = StartDate
65       ChooseDateForm.Show
66       If formCancel Then Exit Sub
67       EndDate = format(ChooseDateForm.MonthView1, "YYYY-MM-DD")
70       Unload ChooseDateForm
    
80   Call QueryNewCondition(SHVkwh, MACROWORKBOOK, "WHERE Reading_Start_Dt BETWEEN", "'" & StartDate & "' AND '" & EndDate & "'")
    
90   meter_id = InputBox("Meter_ID Number: ", Title:="VKWH")
100  If meter_id = "" Then Exit Sub
    
     meter_id = Trim(meter_id)
110  Call QueryNewCondition(SHVkwh, MACROWORKBOOK, "AND Meter_Id =", meter_id)

120  useQuery = QueryBuilder(SHVkwh, MACROWORKBOOK)
    
     If StartDate <> EndDate Then
130     GLBQueryName = meter_id & "_" & StartDate & "_" & EndDate
     Else
        GLBQueryName = meter_id & "_" & StartDate
     End If
140  Call Query(useQuery, False)
150  If formCancel Then Exit Sub

    
     Set fRange = FindInRange("Reading_Meas", Rows(1))
     fRange.Copy
     fRange = meter_id & " Reading_Meas"

160  Call VKWH_Graph
    
170  Exit Sub

End Sub

Sub VKWH_Graph()
    '
    ' Sum KWH
    '
    botRow = ColumnLastRow(3)
    
    range("C1").Select
    range(Selection, Selection.End(xlDown)).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    range("C2").FormulaR1C1 = "=RC[-1]"
    range("C3").FormulaR1C1 = "=RC[-1]+R[-1]C"
    range("C1") = "KWH"
    
    range("C3").AutoFill Destination:=range("C3:C" & botRow)
    'range("C3:C" & botRow).Select
    
    range("E2").Select  ' Voltage
    range(Selection, Selection.End(xlDown)).Select
    
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=range("$E$2:$E$" & botRow)
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    ActiveChart.FullSeriesCollection(1).Name = "=""Volts"""
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "=""KWH"""
    ActiveChart.FullSeriesCollection(2).Values = "='" & ActiveSheet.Name & "'!$C$2:$C$" & botRow
    ActiveChart.ChartArea.Select
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).AxisGroup = 1
    ActiveChart.FullSeriesCollection(2).ChartType = xlLine
    ActiveChart.FullSeriesCollection(2).AxisGroup = 1
    ActiveChart.FullSeriesCollection(1).ChartType = xlLine
    ActiveChart.FullSeriesCollection(2).AxisGroup = 2
    chartName = Mid(ActiveChart.Name, InStr(ActiveChart.Name, " ") + 1) ' = ActiveSheet.ChartObjects.count
    ActiveSheet.ChartObjects(chartName).Activate
    ActiveChart.FullSeriesCollection(2).Select
    ActiveChart.FullSeriesCollection(2).Trendlines.Add
    ActiveChart.FullSeriesCollection(2).Trendlines(1).Select
    Selection.DisplayEquation = True
    ActiveChart.FullSeriesCollection(2).Trendlines(1).DataLabel.Select
    Selection.left = 38.939
    Selection.top = 7.089
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Trendlines.Add
    ActiveChart.FullSeriesCollection(1).Trendlines(1).Select
    Selection.DisplayEquation = True
    ActiveChart.FullSeriesCollection(1).Trendlines(1).DataLabel.Select
    Selection.left = 238.019
    Selection.top = 10.008
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes(chartName).IncrementLeft -15
    ActiveSheet.Shapes(chartName).IncrementTop -262.5
    ActiveSheet.Shapes(chartName).IncrementLeft 53.25
    ActiveSheet.Shapes(chartName).IncrementTop -748.5
    ActiveSheet.Shapes(chartName).ScaleWidth 1.4625, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes(chartName).ScaleHeight 1.3993055556, msoFalse, _
        msoScaleFromTopLeft
    'ActiveChart.ChartTitle.Select
    'ActiveChart.ChartTitle.Text = "VKWH"
    Selection.format.TextFrame2.TextRange.Characters.Text = "VKWH"
    'With Selection.format.TextFrame2.TextRange.Characters(1, 4).ParagraphFormat
    '    .TextDirection = msoTextDirectionLeftToRight
    '    .Alignment = msoAlignCenter
    'End With
    'With Selection.format.TextFrame2.TextRange.Characters(1, 4).Font
    '    .BaselineOffset = 0
    '    .Bold = msoFalse
    '    .NameComplexScript = "+mn-cs"
    '    .NameFarEast = "+mn-ea"
    '    .Fill.Visible = msoTrue
    '    .Fill.ForeColor.RGB = RGB(89, 89, 89)
    '    .Fill.Transparency = 0
    '    .Fill.Solid
    '    .Size = 14
    '    .Italic = msoFalse
    '    .Kerning = 12
    '    .Name = "+mn-lt"
    '    .UnderlineStyle = msoNoUnderline
    '    .Spacing = 0
    '    .Strike = msoNoStrike
    'End With
    ActiveChart.ChartArea.Select
    'ActiveChart.SetElement (msoElementLegendRight)
    ActiveChart.SetElement (msoElementLegendBottom)
    range("O25").Select
    ActiveSheet.ChartObjects(chartName).Activate
    ActiveChart.FullSeriesCollection(2).Trendlines(1).DataLabel.Select
    Selection.left = 82.535
    Selection.top = 10
    ActiveChart.FullSeriesCollection(1).Trendlines(1).DataLabel.Select
    Selection.left = 388.536
    Selection.top = 12.42
End Sub

Sub VKWHSetup()

    meter_id = InputBox("Meter ID: ", Title:="VKWHSetup")
    If meter_id = "" Then Exit Sub
          
    LOAD ChooseDateForm
        
    ChooseDateForm.MonthView1.Value = Now()
    ChooseDateForm.Show
    If formCancel Then Exit Sub
        
    useDate = format(ChooseDateForm.MonthView1, "YYYY-MM-DD")
    Unload ChooseDateForm
    
    nDays = InputBox("Number of Days: ", Title:="VKWHSetup")
    If nDays = "" Then Exit Sub

    Cells(1, 1) = "Meter"
    Cells(2, 1) = "Date"
    For i = 1 To nDays
        Cells(1, i + 1) = meter_id
        Cells(2, i + 1) = format(CDate(useDate) + i - 1, "YYYY-MM-DD")
    Next i
End Sub

Sub VKWHDownload()
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)

    With DBRs
        .CursorLocation = adUseClient ' adUseServer
        .CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
        .LockType = adLockReadOnly ' adLockOptimistic
        Set .ActiveConnection = DBCn
    End With

    nCols = LastColumn()
    For i = 2 To nCols

        selectString = "Select Reading_Meas from PUTL_CERT_DATA_MART_VIEWS.v_meter_reading "
        dateString = "WHERE Reading_Start_Dt = '" & format(Cells(2, i), "YYYY-MM-DD") & "' "
        meterString = "AND Meter_Id = " & Cells(1, i)
        restString = " AND Service_Channel_Num = 1 order by Reading_Dttm"
    
        useQuery = selectString & dateString & meterString & restString
        Debug_Print useQuery
        
        Set DBRs = DBCheckRecordset(DBRs)
        DBRs.Open useQuery, DBCn
        
        fieldCount = DBRs.Fields.count
        recordCount = DBRs.recordCount
        k = 3
        For j = 0 To recordCount - 1
            Cells(k, i) = DBRs.Fields(0).Value
            DBRs.MoveNext
            k = k + 1
        Next j
        DBRs.Close
    Next i
    Debug_Print useQuery
    For j = 1 To recordCount
        Cells(2 + j, 1) = j
    Next j
End Sub

Sub VKWHSum()
    
    botRow = LastRow
    LastCol = LastColumn
    
    For i = 2 To LastColumn
        runSum = Cells(3, i)
        For j = 4 To botRow
            runSum = runSum + Cells(j, i)
            Cells(j, i) = runSum
        Next j
    Next i
End Sub
