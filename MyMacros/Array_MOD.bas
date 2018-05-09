Attribute VB_Name = "Array_MOD"
'
' Check if an array has been allocated, used before addnig an element to an array
'
Function IsArrayAllocated(Arr As Variant) As Boolean
        On Error Resume Next
        IsArrayAllocated = IsArray(Arr) And _
                           Not IsError(LBound(Arr, 1)) And _
                           LBound(Arr, 1) <= UBound(Arr, 1)
End Function

