Attribute VB_Name = "A__History"
' Sept 13, 2017
' - Load / Save workbook to H:\MKT_CS\REV_PRO\Last Gasp\2017
' - Add Calendar to QueryForm for StartDate
'
' Sept 14, 2017
' - Handle if not on a Last Gasp sheet without RunDate available
' - Added ZeroKwH to Query Form
'
' Sept 17, 2017
' - Save / Load files to H: drive
'
' Sept 19, 2017
' - Load files read-only from H: drive
' - collapse/expand apts in ZeroKWH by BP and address
'
' Sept 24, 2017
' - Turn off Automatic Calucation in Query Download
' - UsageDrop broken into Rate tabs
'
' Oct 2, 2017
' - MsgBox message "non SSN file found"
' - FastLoad working
'
' Oct 4th 2017
' - SSNDirectory call fixed in SSNMerge
' - Better positioning of Userforms (centered)

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const PI = 3.14159265358979


