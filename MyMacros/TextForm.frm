VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TextForm 
   Caption         =   "Text Form"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15750
   OleObjectBlob   =   "TextForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TextForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Me.top = Application.top + Application.Height / 2 - Me.Height / 2
    Me.left = Application.left + Application.width / 2 - Me.width / 2
End Sub

