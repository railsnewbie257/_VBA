VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Teradata Login"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnSubmit_Click()
    userName = txtUsername.Text
    Password = txtPassword.Text
    formCancel = False
    Unload LoginForm
End Sub
Private Sub btnCancel_Click()
    formCancel = True
    Unload LoginForm
End Sub

Private Sub UserForm_Initialize()
    txtUsername.Text = LCase(Environ$("Username"))
    'txtUsername.Enabled = False
    txtPassword.SetFocus
    
    Me.top = Application.top + Application.Height / 2 - Me.Height / 2
    Me.left = Application.left + Application.width / 2 - Me.width / 2
    
End Sub

