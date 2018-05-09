Attribute VB_Name = "Box_MOD"
Sub DrawBoxRange(aRange)

    aRange.Borders(xlDiagonalDown).LineStyle = xlNone
    aRange.Borders(xlDiagonalUp).LineStyle = xlNone
    With aRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With aRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With aRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With aRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    aRange.Borders(xlInsideVertical).LineStyle = xlNone
    aRange.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
