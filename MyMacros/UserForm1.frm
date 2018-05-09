VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Progress Meters"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const PI = 3.14159265358979
Sub DemoProgress1()
'
' Progress Bar
'
    Dim intIndex As Integer
    Dim sngPercent As Single
    Dim intMax As Integer
    
    intMax = 100
    For intIndex = 1 To intMax
        sngPercent = intIndex / intMax
        ProgressStyle1 sngPercent, chkPg1Value.Value
        DoEvents
        '------------------------
        ' Your code would go here
        '------------------------
        Sleep 100
    Next

End Sub
Sub DemoProgress6()
'
' Fancy2 Progress Bar
'
    Dim intIndex As Integer
    Dim sngPercent As Single
    Dim intMax As Integer
    
    intMax = 100
    For intIndex = 1 To intMax
        sngPercent = intIndex / intMax
        ProgressStyle8 sngPercent
        DoEvents
        '------------------------
        ' Your code would go here
        '------------------------
        Sleep 100
    Next

End Sub
Sub DemoProgress7()
'
' Fancy3 XP Startup
'
    Dim intIndex As Integer
    Dim sngPercent As Single
    Dim intMax As Integer
    Dim intRepeat As Integer
    
    ' this one can be a simple rolling one
    intMax = 100
    For intRepeat = 1 To 3
        For intIndex = 1 To intMax
            sngPercent = intIndex / intMax
            ProgressStyle9 sngPercent
            DoEvents
            '------------------------
            ' Your code would go here
            '------------------------
            Sleep 25
        Next
    Next
End Sub
Sub DemoProgress8()
'
' Fancy4 Rotary dials
'
    Dim intIndex As Integer
    Dim intMax As Integer
    
    ' this can go up to 999
    If IsNumeric(txtCount.Text) Then
        intMax = CInt(txtCount.Text)
    Else
        intMax = 100
    End If
    For intIndex = 1 To intMax
        ProgressStyle10 intIndex
        DoEvents
        '------------------------
        ' Your code would go here
        '------------------------
        Sleep 100
    Next

End Sub

Sub DemoProgress3()
'
' Progress Fancy
'
    Dim intIndex As Integer
    Dim sngPercent As Single
    Dim intMax As Integer
    
    intMax = 100
    For intIndex = 1 To intMax
        sngPercent = intIndex / intMax
        ProgressStyle4 sngPercent, chkPg3Value.Value
        DoEvents
        '------------------------
        ' Your code would go here
        '------------------------
        Sleep 100
    Next

End Sub
Sub DemoProgress4()
'
' Progress Blips
'
    Dim intIndex As Integer
    Dim sngPercent As Single
    Dim intMax As Integer
    
    intMax = 100
    For intIndex = 1 To intMax
        sngPercent = intIndex / intMax
        ProgressStyle5 sngPercent
        ProgressStyle6 sngPercent
        DoEvents
        '------------------------
        ' Your code would go here
        '------------------------
        Sleep 100
    Next

End Sub

Sub DemoProgress5()
'
' Progress Application Statusbar
'
    Dim intIndex As Integer
    Dim sngPercent As Single
    Dim intMax As Integer
    
    intMax = 100
    For intIndex = 1 To intMax
        sngPercent = intIndex / intMax
        ProgressStyle7 sngPercent
        DoEvents
        '------------------------
        ' Your code would go here
        '------------------------
        Sleep 100
    Next
    Application.StatusBar = False
End Sub
Sub DemoProgress2()
'
' Progress Cirlces
'
    Dim intIndex As Integer
    Dim sngPercent As Single
    Dim intMax As Integer
    
    intMax = 100
    For intIndex = 1 To intMax
        sngPercent = intIndex / intMax
        ProgressStyle2 sngPercent, chkPg2Value.Value
        ProgressStyle3 sngPercent, chkPg2Value.Value
        DoEvents
        '------------------------
        ' Your code would go here
        '------------------------
        Sleep 100
    Next

End Sub

Sub LoadHelp()

    Dim strMsg As String
    
    ' Progress Bar
    strMsg = "The progress bar is based on 4 label controls." & vbLf
    strMsg = strMsg & "One of the labels acts as a sunken container." & vbLf
    strMsg = strMsg & "One act as the coloured progress bar." & vbLf
    strMsg = strMsg & "The other 2 are used to display the percent value." & vbLf
    strMsg = strMsg & "The use of 2 allows the colour of the text to change as the progress bar moves underneath."
    labHelp1.Caption = strMsg
    
    ' Progress Circles
    strMsg = "The progress circle is based on 2 Image controls and a label control." & vbLf
    strMsg = strMsg & "One Image acts as a container." & vbLf
    strMsg = strMsg & "The other as a progress meter. The size and position of which is altered" & vbLf
    strMsg = strMsg & "as the percentage value changes." & vbLf
    strMsg = strMsg & "The label is used to display the percent value."
    labHelp2.Caption = strMsg
    
    ' Fancy Progress Bar
    strMsg = "The fancy progress bar is based on an image control and 2 label controls." & vbLf
    strMsg = strMsg & "The image has an transparent area through which to show the progress." & vbLf
    strMsg = strMsg & "One of the labels acts as the progress bar. Visible through the image." & vbLf
    strMsg = strMsg & "The other is used to display the percent value." & vbLf
    labHelp3.Caption = strMsg
    
    ' Blips Progress Bar
    strMsg = "The blips progress bars are both based on 2 label controls." & vbLf
    strMsg = strMsg & "One label act as the permanent empty blips." & vbLf
    strMsg = strMsg & "The other acts as the cover blip." & vbLf
    strMsg = strMsg & "The top bar increments from left to right." & vbLf
    strMsg = strMsg & "The bottom bar shuffles across from left to right ." & vbLf
    labHelp4.Caption = strMsg
    
    ' Status Bar
    strMsg = "The Status bar uses the Application.Statusbar." & vbLf
    strMsg = strMsg & "The text placed on the statusbar behaves like the Blip progress bar." & vbLf
    strMsg = strMsg & "Text is built using 2 characters." & vbLf
    strMsg = strMsg & "" & vbLf
    strMsg = strMsg & "Try different characters to give alternate looks."
    labHelp5.Caption = strMsg

    ' Turning hoop
    strMsg = "The turning hoop is based on 2 Image controls." & vbLf
    strMsg = strMsg & "One Image acts as a container." & vbLf
    strMsg = strMsg & "The other as a break in the hoop. The position of which is altered" & vbLf
    strMsg = strMsg & "as the percentage value changes."
    labHelp6.Caption = strMsg

    ' Fancy3 XP Startup Bar
    strMsg = "The XP Startup progress bar is based on 2 image controls and a label control." & vbLf
    strMsg = strMsg & "The top image has an transparent area through which to show the second image." & vbLf
    strMsg = strMsg & "The label provides the background colour. The second images movement is visible through the image." & vbLf
    strMsg = strMsg & "Although the example displays a percent movement it can also be played over and over to demonstrate activity."
    labHelp7.Caption = strMsg

    ' Rotary dial
    strMsg = "The Rotary dial is based on 3 image controls within a frame." & vbLf
    strMsg = strMsg & "The images top value is adjusted to give the effect of turning." & vbLf
    strMsg = strMsg & "Because all the dials turn, however slowly as the number changes it can be misleading." & vbLf
    labHelp8.Caption = strMsg

End Sub

Public Sub ProgressStyle5(Percent As Single)
'
' Progress Style 5
' Label Over Label changing text content
' Growth
'
    Dim strTemp As String
    Dim intIndex As Integer
    
    intIndex = Int(Len(labPg5a.Caption) * Percent)
    If intIndex > 0 Then
        strTemp = String(intIndex, "•") & String(Len(labPg5a.Caption) - intIndex, " ")
    Else
        strTemp = String(Len(labPg5a.Caption), " ")
    End If
    labPg5.Caption = strTemp
    
End Sub
Public Sub ProgressStyle6(Percent As Single)
'
' Progress Style 6
' Label Over Label changing text content
' Pulsing
'
    Dim strTemp As String
    Dim intIndex As Integer
    
    intIndex = Int((Percent * 100) Mod (Len(labPg6a.Caption) + 1))
    strTemp = String(Len(labPg6a.Caption), " ")
    If intIndex > 0 Then
        Mid(strTemp, intIndex, 1) = "•"
    End If
    labPg6.Caption = strTemp
    
End Sub
Public Sub ProgressStyle7(Percent As Single)
'
' Progress Style 7
' Application Status bar
' Pulsing
'
    Dim strTemp As String
    Dim intIndex As Integer
    Dim intLen As Integer
    
    intLen = 21
    intIndex = Int((Percent * 100) Mod intLen)
    strTemp = String(intLen, txtPg7a.Text)
    If intIndex > 0 Then
        Mid(strTemp, intIndex, 1) = txtPg7p.Text
    End If
    Application.StatusBar = "Processing " & strTemp
    
End Sub
Public Sub ProgressStyle8(Percent As Single)
'
' Progress Style 8
' Fancy2 Hoop with a break in it
' Pulsing
'
    Dim intAngle As Integer
    Dim sngX As Single
    Dim sngY As Single
    Dim sngRadius As Single
    Dim sngCenterX As Single
    Dim sngCenterY As Single
    Dim sngHalfBreak As Single
    
    sngHalfBreak = imgHoopBreak.Width / 2
    sngRadius = (imgHoop.Width / 2) - sngHalfBreak + 5
    sngCenterX = imgHoop.Left + (imgHoop.Width / 2) - sngHalfBreak
    sngCenterY = imgHoop.Top + (imgHoop.Height / 2) - sngHalfBreak
    intAngle = ((360 / 100) * (Percent * 100) + 90)
    sngX = sngCenterX + ((Cos((intAngle * (PI / 180)))) * sngRadius)
    sngY = sngCenterY + ((Sin((intAngle * (PI / 180)))) * sngRadius)
    imgHoopBreak.Move sngX, sngY
    
End Sub

Sub ProgressStyle1(Percent As Single, ShowValue As Boolean)
'
' Progress Style 1
' Label Over Label
'
    Const PAD = "                         "
    
    If ShowValue Then
        labPg1v.Caption = PAD & Format(Percent, "0%")
        labPg1va.Caption = labPg1v.Caption
        labPg1va.Width = labPg1.Width
    End If
    labPg1.Width = Int(labPg1.Tag * Percent)

End Sub
Sub ProgressStyle2(Percent As Single, ShowValue As Boolean)
'
' Progress Style 2
' Image growing Over Static Image
'
    Dim intWidth As Integer
    Dim intLeft As Integer
    Dim intTop As Integer
    
    intWidth = Int(imgPg2.Tag * Percent)
    intLeft = imgPg2a.Left + ((imgPg2a.Width - intWidth) / 2)
    intTop = imgPg2a.Top + ((imgPg2a.Height - intWidth) / 2)
    imgPg2.Move intLeft, intTop, intWidth, intWidth
    
    If ShowValue Then
        labPg2v.Caption = Format(Percent, "0%")
    End If

End Sub
Sub ProgressStyle3(Percent As Single, ShowValue As Boolean)
'
' Progress Style 3
' Image shrinking Over Static Image
'
    Dim intWidth As Integer
    Dim intLeft As Integer
    Dim intTop As Integer
    
    intWidth = Int(imgPg3.Tag * (1 - Percent))
    intLeft = imgPg3a.Left + ((imgPg3a.Width - intWidth) / 2)
    intTop = imgPg3a.Top + ((imgPg3a.Height - intWidth) / 2)
    imgPg3.Move intLeft, intTop, intWidth, intWidth
    
    If ShowValue Then
        labPg3v.Caption = Format(Percent, "0%")
    End If
    
End Sub


Sub ProgressStyle4(Percent As Single, ShowValue As Boolean)
'
' Progress Style 4
' Label under Image with Transparent section
'
    labPg4.Width = Int(labPg4.Tag * Percent)
    If ShowValue Then labPg4v.Caption = Format(Percent, "0%")

End Sub
Sub ProgressStyle9(Percent As Single)
'
' Progress Style 4
' Label under Image with Transparent section
'
    imgXPMarker.Left = imgXPCover.Left + (labXPBackground.Width * Percent)

End Sub
Sub ProgressStyle10(ByVal Value As Integer)
'
' Progress Style 10
' Rotary dial displaying values 0 to 999
'
    Dim sngHeight As Single
    Dim sngOffset As Single
    Dim intUnit As Integer
    Dim intTen As Integer
    Dim intHundred As Integer
    
    ' get values of each rotary wheel
    If Value \ 100 <> 0 Then
        intHundred = Value \ 100
        Value = Value - (100 * intHundred)
    End If
    If Value \ 10 <> 0 Then
        intTen = Value \ 10
        Value = Value - (10 * intTen)
    End If
    intUnit = Value

    sngOffset = -(imgUnits.Height / 12) * 1
    sngHeight = ((imgUnits.Height / 12) * 10) / 10
    
    ' units
    imgUnits.Top = sngOffset - (sngHeight * intUnit)
    ' tens
    imgTens.Top = sngOffset - (sngHeight * (intTen + (intUnit / 10)))
    ' hundreds
    imgHundreds.Top = sngOffset - (sngHeight * (intHundred + (intTen / 10) + (intUnit / 100)))
End Sub

Private Sub chkPg1Value_Click()

    labPg1v.Visible = chkPg1Value.Value
    labPg1va.Visible = chkPg1Value.Value

End Sub

Private Sub chkPg2Value_Click()

    labPg2v.Visible = chkPg2Value.Value
    labPg3v.Visible = chkPg2Value.Value
    
End Sub

Private Sub chkPg3Value_Click()
    Me.labPg3v.Visible = chkPg3Value.Value
End Sub

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub CommandButton2_Click()

    Application.Cursor = xlWait
    Select Case MultiPage1.Value
    Case 0  ' Bar
        DemoProgress1
    Case 1  ' Circle
        DemoProgress2
    Case 2  ' Fancy
        DemoProgress3
    Case 3  ' Blips
        DemoProgress4
    Case 4  ' Application Status
        DemoProgress5
    Case 5  ' Fancy 2
        DemoProgress6
    Case 6  ' XP Startup
        DemoProgress7
    Case 7  ' Rotary dials
        DemoProgress8
    End Select
    Application.Cursor = xlDefault

End Sub

Private Sub CommandButton3_Click()

    ' show progress bar at 50%
    ProgressStyle1 0.5, chkPg1Value.Value
    DoEvents
    
End Sub

Private Sub CommandButton4_Click()
        
    ' show 60%
    ProgressStyle2 0.6, chkPg2Value.Value
    ProgressStyle3 0.6, chkPg2Value.Value
    DoEvents

End Sub

Private Sub CommandButton5_Click()
        
    ' show 60%
    ProgressStyle4 0.6, chkPg3Value.Value
    DoEvents

End Sub

Private Sub CommandButton6_Click()
        
    ProgressStyle5 0.6
    DoEvents

End Sub

Private Sub CommandButton7_Click()
    
    ProgressStyle8 0.5
    DoEvents

End Sub

Private Sub CommandButton8_Click()

    If IsNumeric(txtValue.Text) Then ProgressStyle10 CInt(txtValue.Text)
    DoEvents

End Sub

Private Sub Label9_Click()

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub UserForm_Initialize()

    ' ProgressBar1
    labPg1.Tag = labPg1.Width
    labPg1.Width = 0
    labPg1v.Caption = ""
    labPg1va.Caption = ""
    
    'ProgressBar 2
    imgPg2.Tag = imgPg2.Width
    imgPg2.Width = 0
    labPg2v.Caption = ""
    
    'ProgressBar 3
    imgPg3.Tag = imgPg3.Width
    imgPg3.Width = 0
    labPg3v.Caption = ""
    
    'ProgressBar 4
    labPg4.Tag = labPg4.Width
    labPg4.Width = 0
    labPg4v.Caption = ""

    'ProgressBar 5
    labPg5.Caption = ""

    ProgressStyle8 0
    ProgressStyle9 0
        
    ProgressStyle10 0
    txtCount.Text = 100
    
    LoadHelp

End Sub
