Attribute VB_Name = "Pivot_MOD"
Sub Do_Pivot_no_use()
Dim useRange As Range
Dim q As String

    useCol = ActiveCell.Column
    Set aRange = Worksheets("Scratch").Range(Cells(1, useCol), Cells(LastRow, useCol))
    aRange.Select
    sourceAddr = RangeToText(aRange)

    Call MakeScratchSheet(False, "Pivot")
    pivotRow = LastRow("Pivot") + 3
    destAddr = RangeToText(ActiveWorkbook.Worksheets("Pivot").Cells(pivotRow, 1))
    pivotName = "PivotTable6"
    ' Set endRow = Range(Cells(1, LastRow("Pivot") + 3))
    ' Debug.Print useRange.Address
    
    Sheets("Pivot").Select
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
            SourceData:=sourceAddr, _
            Version:=xlPivotTableVersion15).CreatePivotTable _
            TableDestination:=destAddr, _
            tableName:=pivotName, _
            DefaultVersion:=xlPivotTableVersion15
 
    '
    
    With Worksheets("Pivot")
        .Cells(LastRow("Pivot"), 1).Select
    End With
    ActiveSheet.PivotTables("PivotTable9").AddDataField ActiveSheet.PivotTables( _
        "PivotTable9").PivotFields("I-210+C-RD HAN"), "Count of I-210+C-RD HAN", _
        xlCount
    With ActiveSheet.PivotTables("PivotTable9").PivotFields("I-210+C-RD HAN")
        .Orientation = xlRowField
        .Position = 1
    End With

End Sub

Sub Do_Pivot()
    PTName = "PivotTable6"
    useCol = ActiveCell.Column
    varName = Cells(1, useCol)
    Set aRange = Range(Worksheets("Scratch").Cells(1, useCol), Worksheets("Scratch").Cells(LastRow, useCol))
    Debug.Print aRange.Parent.Name
    aRange.Select
    sourceAddr = RangeToText(aRange)
    
    Call MakeScratchSheet(False, "Pivot")

    pivotRow = LastRow("Pivot") + 3
    destAddr = RangeToText(ActiveWorkbook.Worksheets("Pivot").Cells(pivotRow, 1))

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=sourceAddr, _
        Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:=destAddr, tableName:=PTName, DefaultVersion _
        :=xlPivotTableVersion15

    Worksheets("Pivot").Activate
    ActiveSheet.PivotTables(PTName).AddDataField _
        ActiveSheet.PivotTables(PTName).PivotFields(varName), "Count of " & varName, _
        xlCount
    With ActiveSheet.PivotTables(PTName).PivotFields(varName)
        .Orientation = xlRowField
        .Position = 1
    End With
End Sub


Sub t2_old()
    Columns("A:A").Select
    Sheets.Add
    ActiveSheet.Name = "Sheet3"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Scratch!R2C1:R1048576C1", Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:="Sheet3!R3C1", tableName:="PivotTable9", DefaultVersion _
        :=xlPivotTableVersion15
    Sheets("Sheet3").Select
    Cells(3, 1).Select
    ActiveSheet.PivotTables("PivotTable9").AddDataField ActiveSheet.PivotTables( _
        "PivotTable9").PivotFields("I-210+C-RD HAN"), "Count of I-210+C-RD HAN", _
        xlCount
    With ActiveSheet.PivotTables("PivotTable9").PivotFields("I-210+C-RD HAN")
        .Orientation = xlRowField
        .Position = 1
    End With
End Sub

