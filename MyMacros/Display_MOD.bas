Attribute VB_Name = "Display_MOD"
Sub SpeedupOn()
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .DisplayStatusBar = False
        .EnableEvents = False
        '.DisplayPageBreaks = False
    End With
End Sub

Sub SpeedupOff()
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .DisplayStatusBar = True
        .EnableEvents = True
        '.DisplayPageBreaks = True
    End With
End Sub

Function ScreenWhat() As Boolean
    ScreenWhat = Application.ScreenUpdating
End Function

Function ScreenOff()
    Application.ScreenUpdating = False
End Function

Function ScreenOn()
    Application.ScreenUpdating = True
End Function

Function CalculationOff()
    Application.Calculation = xlManual
End Function

Function CalculationOn()
    Application.Calculation = xlAutomatic
End Function

Function StatusBarOn()
    Application.DisplayStatusBar = True
End Function

Function StatusBarOff()
    Application.DisplayStatusBar = False
End Function

Sub StatusbarDisplay(Optional s)
    Application.DisplayStatusBar = True
    If IsMissing(s) Then s = "testing..."
        Application.StatusBar = s
        'DoEvents
End Sub


Sub FreezeHeader(Optional SHUse)
Dim origCell As range

    'Call ScreenOff
    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    SHOrig = ActiveSheet.Name
    Set origCell = ActiveCell
    
    Sheets(SHUse).Activate
    'Cells(1, 1).Select
    Rows(2).Select
    
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    
    Sheets(SHOrig).Activate
    origCell.Select
    Call ScreenOn
End Sub

