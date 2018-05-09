Attribute VB_Name = "QueryParser_MOD"
Sub QueryParser()

    Set selectRange = FindInRange("select", Cells)
    Set fromRange = FindInRange("from", Cells)
    
    MsgBox "Select X:" & selectRange.Row & " Y:" & selectRange.Column & vbNewLine & _
            "From X:" & fromRange.Row & " Y:" & fromRange.Column & vbNewLine
    
End Sub
