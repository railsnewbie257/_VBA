Attribute VB_Name = "DeleteRecord_MOD"
'
' Checks the table name is in the upper left corner
'
Function TOCformat(aRange)
    With aRange.CurrentRegion.Cells(1, 1)
        If .Font.Bold = True And _
           .Font.color = BLUE And _
           .Interior.color = ORANGE And _
            InStr(.Text, ".") > 0 Then
                TOCformat = True
        Else
                TOCformat = False
        End If
    End With
End Function
'
' This routine will delete a record which has been downloaded to a spreadsheet
' using Table by Rows, assumes the table name is in the upper left hand corner of the CurrentRegion
'
Sub DeleteRecord()
Dim currentRange As range
    '
    ' check the format of the CurrentRegion
    '
    Set currentRange = ActiveCell.CurrentRegion
    '
End Sub

Sub t()
    Call TOCformat(ActiveCell)
End Sub


