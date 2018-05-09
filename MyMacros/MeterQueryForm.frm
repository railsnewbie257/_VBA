VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MeterQueryForm 
   Caption         =   "Meter Query Form"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5535
   OleObjectBlob   =   "MeterQueryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MeterQueryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
    formCancel = True
    Unload Me
End Sub

Private Sub btnPreviousQuery_Click()

    With Workbooks(MACROWORKBOOK).Sheets("Pallette")
        txtDatabaseName = .Cells(3, 2)
        txtTableName = .Cells(3, 3)
        txtSelect = .Cells(3, 4)
    End With
    
End Sub

Private Sub btnSubmit_Click()
    formCancel = False
    
    With Workbooks(MACROWORKBOOK).Sheets("Pallette")
        .Cells(5, 3) = Replace(txtQuery, vbNewLine, "||")
        .Cells(5, 3).Font.color = NOCOLOR
        
        .Cells(5, 1) = txtInText
        .Cells(5, 1).Font.color = NOCOLOR
        
        .Cells(5, 2) = txtInRange
        .Cells(5, 2).Font.color = NOCOLOR

        .Cells(5, 3) = txtOutRange
        .Cells(5, 3).Font.color = NOCOLOR
        
        .Cells(5, 4) = txtSelect
        .Cells(5, 4).Font.color = NOCOLOR
        
        .Cells(5, 5) = txtWhere
        .Cells(5, 5).Font.color = NOCOLOR
    
        .Cells(3, 2) = txtDatabaseName
        .Cells(3, 2).Font.color = NOCOLOR
        .Cells(3, 3) = txtTableName
        .Cells(3, 3).Font.color = NOCOLOR
        .Cells(3, 4) = txtSelect
        .Cells(3, 4).Font.color = NOCOLOR
    End With
    
    MeterQueryForm.Hide
End Sub

Private Sub lblQuery_Click()

End Sub
Private Sub opt_da_customer_vw_Click()
    txtDatabaseName = "da_customer_vw"
End Sub

Private Sub opt_dl_oge_analytics_Click()
    txtDatabaseName = "dl_oge_analytics"
End Sub

Private Sub opt_putlvw_Click()
    txtDatabaseName = "putlvw"
End Sub

Private Sub txtDatabaseName_Change()
Dim tDBName As String, tTBName As String
    
    k = InStr(1, txtDatabaseName, ".")
    If k > 0 Then
        Call SplitFullTableName(txtDatabaseName, tDBName, tTBName)
        txtDatabaseName = tDBName
        txtTableName = tTBName
    End If
    
End Sub

Private Sub txtDatabaseName_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim tDBName As String, tTBName As String
    
    tName = txtDatabaseName
    
    Select Case tName
        Case "dl_oge_analytics"
            txtDatabaseName = "putlvw"

        Case "putlvw"
            txtDatabaseName = "dl_oge_analytics"

        Case Else
            txtDatabaseName = "dl_oge_analytics"
    End Select
        
    k = InStr(1, txtDatabaseName, ".")
    If k > 0 Then
        Call SplitFullTableName(txtDatabaseName, tDBName, tTBName)
        txtDatabaseName = tDBName
        txtTableName = tTBName
    End If
    
End Sub

Private Sub txtInRange_Click()
    txtInRange = ""
    
End Sub

Private Sub txtOutRange_Click()
    i = 1
End Sub

Private Sub txtInRange_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim aRange
    
    On Error Resume Next
    Set aRange = Nothing
    Set aRange = Application.InputBox("Input Range", "InRange", "", Type:=8)
    If aRange Is Nothing Then
        txtInRange = ""
        Exit Sub
    End If
    txtInText = ""
    txtInRange = aRange.Address
End Sub

Private Sub txtInText_Change()
    txtInRange = ""
End Sub

Private Sub txtOutRange_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim aRange
    
    On Error Resume Next
    Set aRange = Nothing
    Set aRange = Application.InputBox("Out Range", "OutRange", "", Type:=8)
    If aRange Is Nothing Then
        Exit Sub
    End If
    txtOutRange = aRange.Address
End Sub

Private Sub txtTableName_Change()
Dim tDBName As String, tTBName As String
    
    k = InStr(1, txtTableName, ".")
    If k > 0 Then
        Call SplitFullTableName(txtTableName, tDBName, tTBName)
        txtDatabaseName = tDBName
        txtTableName = tTBName
    End If

End Sub

Private Sub UserForm_Initialize()

    lblWhere.Font.Bold = True
    lblSelect.Font.Bold = True
    
    Me.top = Application.top + Application.Height / 2 - Me.Height / 2
    Me.left = Application.left + Application.width / 2 - Me.width / 2
End Sub
