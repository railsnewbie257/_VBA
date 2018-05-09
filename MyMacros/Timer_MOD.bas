Attribute VB_Name = "Timer_MOD"
Dim startTime As Date, stopTime As Date

Sub StartTimer()
    startTime = Now() ' global
    Debug.Print format(startTime, "HH:MM:SS")
End Sub

Sub StopTimer()
    stopTime = Now()  ' global
    Debug.Print format(stopTime, "HH:MM:SS")
End Sub

Function ElapsedTime()
    ElapsedTime = "Elapsed Time: " & format(stopTime - startTime, "HH:MM:SS")
End Function

Sub testTimer()
    Call StartTimer
    Debug.Print Now()
    Call Application.Wait(Now + TimeValue("0:01:00"))
    Debug.Print Now()
    Call StopTimer
    
    Debug.Print ElapsedTime & " " & format(startTime, "HH:MM:SS") & " " & format(stopTime, "HH:MM:SS")
End Sub

Sub FillData()
    Call StartTimer
    For i = 1 To 10000
        Cells(i, 1) = 100
    Next i
    Call StopTimer
    Debug.Print ElapsedTime
End Sub
