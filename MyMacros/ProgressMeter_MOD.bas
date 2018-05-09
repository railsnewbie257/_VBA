Attribute VB_Name = "ProgressMeter_MOD"
Sub ProgressMeterShow(Optional numerator, Optional denominator)
Const PAD = "                         "
        pct = numerator / denominator
        LOAD ProgressMeter
        ProgressMeter.labPg1v.caption = PAD & format(pct, "0%")
        ProgressMeter.labPg1va.caption = ProgressMeter.labPg1v.caption
        ProgressMeter.labPg1va.width = ProgressMeter.labPg1.width

        ProgressMeter.labPg1.width = Int(ProgressMeter.labPg1.Tag * pct)
        If Not ProgressMeter.Visible Then ProgressMeter.Show False
        DoEvents
End Sub

Sub ProgressMeterClose()
    Unload ProgressMeter
End Sub

Sub ProgressStepsTest()
    Call ProgressStepCaption(1, "first")
    
    If Not ProgressCheckBox.Visible Then ProgressCheckBox.Show False
    DoEvents
End Sub

Sub ProgressStepCaption(stepNum, caption)
    LOAD ProgressCheckBox
    Select Case stepNum
        Case 1
            ProgressCheckBox.CheckBox1.caption = caption
            ProgressCheckBox.CheckBox1 = True
        Case 2
        Case 3
        Case 4
        Case 5
        Case 6
    End Select
    DoEvents
End Sub

Sub ProgressStepCheck(stepNum)
    Select Case stepNum
        Case 1
            ProgressCheckBox.CheckBox1 = True
        Case 2
        Case 3
        Case 4
        Case 5
        Case 6
    End Select
    If Not ProgressCheckBox.Visible Then ProgressCheckBox.Show False
End Sub

Sub ProgressStepClose()
    If ProgressCheckBox.Visible Then Unload ProgressCheckBox
End Sub

