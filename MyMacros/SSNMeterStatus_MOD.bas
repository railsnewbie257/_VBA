Attribute VB_Name = "SSNMeterStatus_MOD"
Option Explicit
'
' Assumes on the target workbook
'
Sub SSNMeterStatus()
Dim SHTo As String, WBTo As String

Dim SHFrom As String, WBFrom As String
Dim useDate As String
Dim useCol As Long, rundateCol As Long, rownumCol As Long
Dim meterColTo As Long, meterColFrom As Long
Dim useName As String, resultCode As String
Dim fromCol As Long, toCol As Long
Dim fromRange As Range, toRange As Range, zipCol As Long
Dim fromColRange As Range, toColRange As Range
Dim indexRange As Range

    rundateCol = FindColumnHeader("rundate")
    useDate = format(Cells(2, rundateCol).Value, "YYYY-MM-DD")
    
    WBTo = ActiveWorkbook.Name
    SHTo = ActiveSheet.Name
    
    meterColTo = FindColumnHeader("meter_serial_num", SHTo, WBTo)
    
    useName = "SSN-" & useDate & ".xlsx"
    resultCode = Dir(SSNPATH & useName)
    If (useName = resultCode) Then
        Workbooks.Open SSNPATH & useName, ReadOnly:=True
        WBFrom = ActiveWorkbook.Name
        SHFrom = ActiveSheet.Name
    Else
        MsgBox "SSN Meter file not found." & vbNewLine & vbNewLine & "Please process SNS Meters for " & useDate, Title:="SSN Meter Status"
        Exit Sub
    End If

    
    meterColFrom = FindColumnHeader("src_name", SHFrom, WBFrom)
    If (meterColFrom < 0) Then
        MsgBox "SSN Spreadsheet incorrect." & vbNewLine & vbNewLine & "Exiting", Title:="SSN Meter Status"
        Exit Sub
    End If
    
    Workbooks(WBFrom).Worksheets(SHFrom).Activate
    useCol = FindColumnHeader("src_name")
    Set fromRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(2, useCol), _
                          Workbooks(WBFrom).Worksheets(SHFrom).Cells(2, useCol))
    Call AddRowNumbers(useCol)
    
    fromCol = FindColumnHeader("src_ops_state", SHFrom, WBFrom)
    Set fromColRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(2, fromCol), _
                            Workbooks(WBFrom).Worksheets(SHFrom).Cells(2, fromCol))
    
    Workbooks(WBTo).Worksheets(SHTo).Activate
    useCol = FindColumnHeader("meter_serial_num", SHTo, WBTo)
    Set toRange = Range(Workbooks(WBTo).Worksheets(SHTo).Cells(2, useCol), _
                        Workbooks(WBTo).Worksheets(SHTo).Cells(2, useCol))

    useCol = FindColumnHeader("meter_active_status_code")
    Set toColRange = Range(Workbooks(WBTo).Worksheets(SHTo).Cells(2, useCol), _
                        Workbooks(WBTo).Worksheets(SHTo).Cells(2, useCol))
    
    Call MatchRowIndex(fromRange, toRange, fromColRange, toColRange)
    
    Workbooks(WBFrom).Close (False)
    
    Exit Sub
    '
    '==============================================================================================
    '
    Set indexRange = toRange.Offset(0, 1)
    
    toCol = FindColumnHeader("meter_serial_num", SHTo, WBTo)
    Set toRange = Range(Workbooks(WBTo).Worksheets(SHTo).Cells(2, toCol), _
                        Workbooks(WBTo).Worksheets(SHTo).Cells(2, toCol))
                        
    fromCol = FindColumnHeader("src_ops_state", SHFrom, WBFrom)
    Set fromRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(2, fromCol), _
                        Workbooks(WBFrom).Worksheets(SHFrom).Cells(2, fromCol))
                        
    Call CopyLookupValues(indexRange, fromRange, toRange)
    
    Workbooks(WBFrom).Close (False)
    
    zipCol = FindColumnHeader("proximity_zip_code")
    If zipCol < 0 Then Call ProximityZipCodeColumn
    
End Sub

