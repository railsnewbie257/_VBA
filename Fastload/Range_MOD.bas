Attribute VB_Name = "Range_MOD"
Sub RangeToValues(inRange)
    inRange.Copy
    inRange.PasteSpecial _
        Paste:=xlPasteValues, _
        Operation:=xlNone, _
        SkipBlanks:=False, _
        Transpose:=False
    Call ClearClipboard
End Sub
    
    Function Union2(ParamArray Ranges() As Variant) As range
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Union2
    ' A Union operation that accepts parameters that are Nothing.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim n As Long
        Dim RR As range
        For n = LBound(Ranges) To UBound(Ranges)
            If IsObject(Ranges(n)) Then
                If Not Ranges(n) Is Nothing Then
                    If TypeOf Ranges(n) Is Excel.range Then
                        If Not RR Is Nothing Then
                            Set RR = Application.Union(RR, Ranges(n))
                        Else
                            Set RR = Ranges(n)
                        End If
                    End If
                End If
            End If
        Next n
        Set Union2 = RR
    End Function
    
Function MyUnion(aRange, bRange) As range
    If Not (aRange Is Nothing) And Not (bRange Is Nothing) Then
        Set MyUnion = Application.Union(aRange, bRange)
    ElseIf aRange Is Nothing Then
        Set MyUnion = bRange
    Else
        Set MyUnion = aRange
    End If
End Function

Function RangeHasValues(inRange) As range
Dim numRange As range, txtRange As range
    Set RangeHasValues = inRange.SpecialCells(xlCellTypeConstants)
End Function
