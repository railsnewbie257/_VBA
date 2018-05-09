VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressMeter 
   Caption         =   "Progress Meter"
   ClientHeight    =   1830
   ClientLeft      =   2040
   ClientTop       =   2370
   ClientWidth     =   4710
   OleObjectBlob   =   "ProgressMeter.frx":0000
End
Attribute VB_Name = "ProgressMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub btnCancel_Click()
    formCancel = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        ' Your cod
        ' Tip: If you want to prevent closing UserForm by Close (×) button in the right-top corner of the UserForm, just uncomment the following line:
        Cancel = True
    End If
End Sub

Private Sub UserForm_Initialize()
Dim pct As Single
    Const PAD = "                         "
    
    ' ProgressBar1
    labPg1.Tag = labPg1.width
    labPg1.width = 0
    labPg1v.caption = ""
    labPg1va.caption = ""

    labPg1v.Visible = True
    labPg1va.Visible = True
    
    ProgressMeter.top = Application.top + Application.Height / 2 - ProgressMeter.Height / 2
    ProgressMeter.left = Application.left + Application.width / 2 - ProgressMeter.width / 2
    
End Sub

Private Sub UserForm_Terminate()
    Unload Me
End Sub
