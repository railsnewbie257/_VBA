Attribute VB_Name = "QueryAnalyst_MOD"
Sub QueryAnalyst()

    DBQueryAnalystForm.Show vbModeless
    If formCancel Then Exit Sub
    Debug.Print "in QueryAnalyst"
    
    Exit Sub
    
    Columns("B:B").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
End Sub
