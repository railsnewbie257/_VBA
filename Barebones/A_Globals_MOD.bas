Attribute VB_Name = "A_Globals_MOD"
'
' These are Global Variables for states across Sub calls
'
'
' For Database Access
'
Public userName As String
Public Username_save As String
Public Password As String
Public Password_save As String
Public loginCancel As Boolean
'
' Query stuff --------------------------------------------------------------------
'
Public userQuery As String
Public DBGlbTable As String
'
Public GLBQueryBaseWB As String, GLBQueryBaseSH As String
'
' Database Connection ------------------------------------------------------------
'
Public DBGlbAdodbError As Boolean

Public DBGlbConnection As ADODB.Connection
Public DBGlbRecordset As ADODB.Recordset
Public DBGlbRecordsToRead As Long  ' the number of records to read
Public DBGlbRecordsFound As Long
Public DBGlbDownloadHeader As Boolean
Public DBGlbTargetColumn As String ' use this for worksheet name
Public DBGlbUseCallback As String

Public GLBDownloadByRow As Boolean ' direction to read the Resultset, True=rows, False=columns,
Public GLBDownloadByColumn As Boolean ' direction to read the Resultset, True=column, False=rows
Public GLBNewWorkbook As Boolean ' where to put the Resultset, if false, then defaults to new Sheet
Public GLBSameSheet As Boolean ' download to the same sheet
Public GLBDownloadWB As String
Public GLBDownloadSH As String
Public GLBColumnNameWB As String
Public GLBColumnNameSH As String
Public GLBPlacementRow As Long
Public GLBPlacementColumn As Long
Public GLBDownloadShowTableName As Boolean
Public GLBManualPlacement As Boolean

Public DatabaseName As String
Public viewName As String
'
' Userform globals
'
Public GLBTableName As String
Public GLBTableNameList() As String

Public GLBDatabaseName As String
Public GLBDatabaseNameList() As String
'
'
Public formCancel As Boolean  ' returned from Userforms
Public GlbStatusBarTxt As String
Public GlbUseDate As String
Public GLBQueryName As String ' determines the base tabe name for data eg LastGasp, ZeroKWH,
Public GLBUserQuery As String  ' the query string to use
Public GLBColumnNamesWB As String
Public GLBColumnNamesSH As String
'
Public GLBTableNameWorkbook As String
Public GLBTableNameSheet As String



Public GLBOpenReadOnly As Boolean
'
' Progress Bar
'
Public GLBProgressNumerator As Long
Public GLBProgressDenominator As Long
'
' Format of Query sheets
'
Public Const QUERYDATACOL = 2
'
' Point to the TD tables
'
Public Const TD_LASTGASP = "LastGasp" ' "Last_Gasp_2" '
Public Const TD_USAGEDROP = "UsageDrop"
Public Const TD_CTSNOOOP = "CTSnoop"
Public Const TD_ZEROKWH = "ZeroKWH"
Public Const TD_RECEIVEDENERGY = "ReceivedEnergy"
Public Const TD_UNDERVOLTAGE = "UnderVoltage"
'
' File names and paths for Fraud Squad
'
Public Const LASTGASPPATH = "H:\MKT_CS\REV_PRO\Last Gasp\2017\"
Public Const USAGEDROPPATH = "H:\MKT_CS\REV_PRO\Usage Drop\2017\"
Public Const KV2CUNDERVOLTAGEPATH = "H:\MKT_CS\REV_PRO\KV2C Undervoltage\2017\"
Public Const ZEROKWHPATH = "H:\MKT_CS\REV_PRO\Zero KWH\2017\"
Public Const SSNPATH = "H:\MKT_CS\REV_PRO\SSN\2017\"
Public GLBSaveFilename As String
Public GLBFilePath As String       ' used for saving files in correct directories
'-----------------------------------------------
'
' Standard Colors
'
Public currentColor As Long
'
'
Public Const RUST = 192
Public Const RED = 255
Public Const HILITERED = 393372
Public Const ORANGE = 49407
Public Const YELLOW = 65535
Public Const LIGHTGREEN = 5296274
Public Const GREEN = 5287936
Public Const LIGHTBLUE = 15773696
Public Const BLUE = 12611584
Public Const BLACK = 10
Public Const DARKBLUE = 6299648
Public Const PURPLE = 10498160
Public Const PINK = 13395711

Public Const NOCOLOR = 16777215
Public Const LIGHTPINK = 13421823 ' 0.599993896298105
Public Const HILITEPINK = 13551615

Public Const GREY = 9868950
Public Const LIGHTGREY = 14540253
Public Const GREYSPECKLE = 3
Public Const APTCOLLAPSE = 4
'
' Size for Proximity
'
Public Const ROWSPAN = 10
'
' The shape of the data body
'
Public Const DATASTARTROW = 2
Public Const DATAFIRSTCOL = 1
'
Public lg_proximity_row As Long

Public Const MACROWORKBOOK = "My_Macros.xlsm"
'Public Const filePath = "C:\oge\"
'Public Const filePrefix = "SSN-"
'
' DB Stuff
'
Public recordCount As Long

Public startTime As Date
Public finishTime As Date
'
'
Public Const WBMacros = "My_Macros.xlsm"

Sub testdate()

    startTime = Timer
    Debug_Print "> " & format(Now(), "HH:nn:ss") & "." & Strings.Right(Strings.format(Timer, "#0.00"), 2)
        For i = 1 To 100000000
        Next i
    finishTime = Timer
    
    Debug_Print startTime
    Debug_Print finishTime
    Debug_Print format(Now(), "HH:nn:ss") & "." & Strings.Right(Strings.format(Timer, "#0.00"), 2)
    Debug_Print format(Now(), "HH:nn:ss") & "." & Strings.Right(Strings.format(Timer, "#0.00"), 2)
    
End Sub

