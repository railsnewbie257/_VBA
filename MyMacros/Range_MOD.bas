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
'
' Returns two strings from a selection
'
Sub RangeChooseTwo(aRange As range, s1 As range, s2 As range)
Dim i As Integer
    
    t = aRange.count
    If aRange.Areas.count = 1 Then
    On Error GoTo Err:
        Set s1 = aRange(1)
        Set s2 = aRange(2)
    Else
        Set s1 = aRange.Areas(1)
        Set s2 = aRange.Areas(2)
    End If
    Exit Sub
    
Err:
    i = 2
    Resume Next
End Sub

     Function ProperUnion(ParamArray Ranges() As Variant) As range
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ProperUnion
    ' This provides Union functionality without duplicating
    ' cells when ranges overlap. Requires the Union2 function.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim ResR As range
        Dim n As Long
        Dim R As range
        
        If Not Ranges(LBound(Ranges)) Is Nothing Then
            Set ResR = Ranges(LBound(Ranges))
        End If
        For n = LBound(Ranges) + 1 To UBound(Ranges)
            If Not Ranges(n) Is Nothing Then
                For Each R In Ranges(n).Cells
                    If Application.Intersect(ResR, R) Is Nothing Then
                        Set ResR = Union2(ResR, R)
                    End If
                Next R
            End If
        Next n
        Set ProperUnion = ResR
    End Function
    
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
