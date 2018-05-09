Attribute VB_Name = "Strings_MOD"
Function LegalSheetName(s) As String
    s = Replace(s, ":", "")
    s = Replace(s, "/", "")
    s = Replace(s, "\", "")
    s = Replace(s, "*", "")
    s = Replace(s, "[", "")
    LegalSheetName = Replace(s, "]", "")
End Function
