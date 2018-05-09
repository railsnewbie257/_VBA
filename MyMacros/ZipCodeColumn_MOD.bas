Attribute VB_Name = "ZipCodeColumn_MOD"
Option Explicit

Sub ProximityZipCodeColumn()
Dim zipCol As Long, newCol As Long
Dim botRow As Long
Dim aRange As Range

    zipCol = FindColumnHeader("pos_address_zip_code")
    botRow = ColumnLastRow(zipCol)
    
    newCol = ColumnInsertRight(zipCol)
    Set aRange = Range(Cells(2, newCol), Cells(botRow, newCol))
    
    aRange.Formula = "=LEFT(" & Cells(2, zipCol).Address(False, False) & ",5)"
    
    Call RangeToValues(aRange)
    
    Cells(1, newCol) = "proximity_zip_code"
    Call ColorRange(Cells(1, newCol), LIGHTBLUE)
    '
    ' cleanup old column
    '
    Columns(zipCol).Delete
    
End Sub

Sub tester()
Dim i As Integer
Dim t As Integer

    i = 4
    t = i + 5
    MsgBox t
End Sub

