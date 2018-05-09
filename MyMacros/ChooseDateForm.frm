VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChooseDateForm 
   Caption         =   "Choose A Date"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3855
   OleObjectBlob   =   "ChooseDateForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChooseDateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub monthview1_click()
    Debug.Print MonthView1.Value
End Sub
Private Sub btnCancel_Click()
    formCancel = True
    Hide
End Sub
Private Sub btnSubmit_Click()
    formCancel = False
    Debug.Print MonthView1.Value
    Hide
End Sub

Private Sub UserForm_Initialize()
    formCancel = False

    Me.top = Application.top + Application.Height / 2 - Me.Height / 2
    Me.left = Application.left + Application.width / 2 - Me.width / 2
End Sub
