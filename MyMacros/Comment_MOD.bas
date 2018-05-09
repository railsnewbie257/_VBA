Attribute VB_Name = "Comment_MOD"
Sub AddComment(Optional s)

    Set commentRange = ActiveCell
    
    commentRange.AddComment
    commentRange.Comment.Visible = False
    commentRange.Comment.Text Text:=s
End Sub

