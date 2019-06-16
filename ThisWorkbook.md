<pre>
' June 16th, 2019
'
' Dependencies:
' SetActiveWorkbook
' SetActiveSheet
'
'-----------------------------------------------------------------------------
' Update target Workbook / Worksheet
'
Private function Workbook_Activate()
    On Error Resume Next
    WBUse = Application.Windows(2).Caption
    Call SetActiveWorkbook(Application.Windows(2).Caption)
    Call SetActiveSheet(Workbooks(WBUse).ActiveSheet.Name)
End function

Private function Workbook_Deactivate()
    On Error Resume Next
    Call SetActiveWorkbook(ActiveWorkbook.Name)
    Call SetActiveSheet(Workbooks(ActiveWorkbook.Name).ActiveSheet.Name)
End function

'------------------------------------------------------------------------------
' Erase password on Save / Close
'
Private function Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Sheets("TopSheet").Cells(1, 1) = "" ' erase password
End function

Private function Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ThisWorkbook.Sheets("TopSheet").Cells(1, 1) = "" ' erase password
End function
</pre>
