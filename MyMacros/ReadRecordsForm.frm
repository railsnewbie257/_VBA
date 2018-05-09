VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReadRecordsForm 
   Caption         =   "Records To Download"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4755
   OleObjectBlob   =   "ReadRecordsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReadRecordsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkHeader_Click()
    chkHeader.Font.Bold = chkHeader
End Sub

Private Sub chkIncludeTableName_Click()
   chkIncludeTableName.Font.Bold = chkIncludeTableName
End Sub

Private Sub cmdCancel_Click()
    formCancel = True
    DBGlbDownloadHeader = False
    Unload Me
End Sub

Private Sub cmdSubmit_Click()
    DBGlbDownloadHeader = chkHeader
    GLBNewWorkbook = optNewWorkbook
    GLBDownloadShowTableName = chkIncludeTableName
    
    If Not chkColumnNameSheet Then
        GLBColumnNameWB = ""
        GLBColumnNameSH = ""
    End If
    
    GLBManualPlacement = optManualPlacement
    GLBDownloadByColumn = optByCol
    GLBSameSheet = optCurrentSheet
    DBGlbRecordsToRead = CDbl(txtRecordsToRead)
    Unload Me
End Sub

Private Sub optFirst10_Click()
    txtRecordsToRead = 10
End Sub

Private Sub optLimit2000_Click()
    txtRecordsToRead = Application.WorksheetFunction.Min(2000, DBGlbRecordsFound)
    optLimit2000 = True
End Sub

Private Sub optManualPlacement_Click()
Dim aRange As Range

    
    On Error Resume Next
    Set aRange = Nothing
    ReadRecordsForm.Hide
    Set aRange = Application.InputBox("Select Download Location", Title:="DBQuery", Default:=Selection.Address, Type:=8)
    ReadRecordsForm.Show
    If aRange Is Nothing Then
        optManualPlacement = False
        Exit Sub
    End If
    
    GLBPlacementRow = aRange.Row
    GLBPlacementColumn = aRange.Column
    GLBDownloadSH = aRange.Parent.Name
    GLBDownloadWB = aRange.Parent.Parent.Name
    GLBManualPlacement = True
End Sub

Private Sub optNewWorkbook_Click()
    chkColumnNameSheet = False
End Sub

Private Sub optReadAll_Click()
    txtRecordsToRead = DBGlbRecordsFound
    optReadAll = True
End Sub

Private Sub optTenRows_Click()
    txtRecordsToRead = 10
End Sub

Private Sub txtRecordsToRead_Change()
    txtRecordsToRead = format(txtRecordsToRead, "##,###,##0")
    optFirst10 = False
    optLimit2000 = False
    optReadAll = False
    If val(txtRecordsToRead) = 2000 Then optLimit2000 = True
    If val(txtRecordsToRead) = 10 Then optFirst10 = True
    If val(txtRecordsToRead) = DBGlbRecordsFound Then optReadAll = True
    
End Sub

Private Sub UserForm_Initialize()
    formCancel = False
    If DBGlbRecordsFound = 0 Then
        lblRecordsFound.caption = "No records found."
    Else
        lblRecordsFound.caption = "Found " & format(DBGlbRecordsFound, "#,##0") & " records."
    End If
    lblRecordsFound.Font.Bold = True
    '
    ' Vertically aligned top message
    '
    Label1.width = 180
    Label1.Height = 25
    With lblRecordsFound
        .AutoSize = False
        .Height = 12
        .width = 175
        .AutoSize = False
        .top = Label1.top + ((Label1.Height - .Height) / 2)
        .left = Label1.left + ((Label1.width - .width) / 2)
    End With
        '
        '
    chkHeader.Font.Bold = True
    chkHeader = True
    '
    ' Handle Column Names
    '
    If GLBQueryName = "ColumnNames" Then
        If GLBColumnNamesWB <> "" Then chkColumnNameSheet = True
        chkIncludeTableName = True
    Else
        chkIncludeTableName = GLBDownloadShowTableName
    End If
    '
    ' Defaults
    '
    optReadAll = True
    optByCol = GLBDownloadByColumn
    If Not GLBDownloadByColumn Then optByRow = True
    optNewWorkbook = True
    GLBManualPlacement = False
    '
    txtRecordsToRead = DBGlbRecordsFound
    If GLBQueryName = "ColumnNames" Then optCurrentSheet = True
    
    Me.top = Application.top + Application.Height / 2 - Me.Height / 2
    Me.left = Application.left + Application.width / 2 - Me.width / 2
End Sub
