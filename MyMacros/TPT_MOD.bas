Attribute VB_Name = "TPT_MOD"
Dim filePath As String
Dim tableName As String

Sub Do_TPT()


    On Error Resume Next
    tableName = InputBox("Table Name?", Title:="TPT", Default:=GLBTableName)
    If IsEmpty(tableName) Then Exit Sub
    
    filePath = "C:\oge\tpt\" & tableName & ".tpt"
    
    Kill filePath
    
    Call FastLoadWrite(filePath, "DEFINE JOB " & tableName & "_LOAD (")
    
    Call TPT_File_Schema
    Call TPT_Producer_Operator
    Call TPT_DDL_Operator
    Call TPT_Load_Operator
    Call TPT_Setup_Tables
    Call TPT_Load_File
    
    Call FastLoadWrite(filePath, ");")
    Call FastLoadWrite(filePath, ");")
    
    MsgBox "Finished"
End Sub

'
' Define File Schema and Producer Operatos to read files
'
Sub TPT_File_Schema()
    Call FastLoadWrite(filePath, "DEFINE SCHEMA SCHEMA_" & tableName)
    Call FastLoadWrite(filePath, "(")
    Call FastLoadWrite(filePath, "EMP_NAME VARCHAR(50),")
    Call FastLoadWrite(filePath, "AGE VARCHAR(2)")
    Call FastLoadWrite(filePath, ");")
End Sub

Sub TPT_DDL_Operator()
Call FastLoadWrite(filePath, "DEFINE OPERATOR od_" & tableName)
Call FastLoadWrite(filePath, "Type DDL")
Call FastLoadWrite(filePath, "Attributes")
Call FastLoadWrite(filePath, "(")
Call FastLoadWrite(filePath, "VARCHAR LogonMech = 'LDAP',")
Call FastLoadWrite(filePath, "VARCHAR TdpId = 'td1',")
Call FastLoadWrite(filePath, "VARCHAR UserName = 'pihpj',")
Call FastLoadWrite(filePath, "VARCHAR UserPassword = 'Okcoge2103b',")
Call FastLoadWrite(filePath, "VARCHAR ErrorList = '3807'")
Call FastLoadWrite(filePath, ");")
End Sub

Sub TPT_Producer_Operator()

Call FastLoadWrite(filePath, "DEFINE OPERATOR op_" & tableName)
Call FastLoadWrite(filePath, "TYPE DATACONNECTOR PRODUCER")
Call FastLoadWrite(filePath, "SCHEMA SCHEMA_" & tableName)
Call FastLoadWrite(filePath, "Attributes")
Call FastLoadWrite(filePath, "(")
Call FastLoadWrite(filePath, "VARCHAR FileName = 'C:\OGE\TPT\" & tableName & ".csv',")
Call FastLoadWrite(filePath, "VARCHAR Format = 'Delimited',")
Call FastLoadWrite(filePath, "VARCHAR OpenMode = 'Read',")
Call FastLoadWrite(filePath, "VARCHAR TextDelimiter = '|'")
Call FastLoadWrite(filePath, ");")
End Sub

Sub TPT_Load_Operator()
Call FastLoadWrite(filePath, "DEFINE OPERATOR ol_" & tableName)
Call FastLoadWrite(filePath, "Type LOAD")
Call FastLoadWrite(filePath, "SCHEMA *")
Call FastLoadWrite(filePath, "Attributes")
Call FastLoadWrite(filePath, "(")
Call FastLoadWrite(filePath, "VARCHAR LogonMech = 'LDAP',")
Call FastLoadWrite(filePath, "VARCHAR PrivateLogName = 'load_log',")
Call FastLoadWrite(filePath, "VARCHAR TdpId = 'td1',")
Call FastLoadWrite(filePath, "VARCHAR UserName = 'pihpj',")
Call FastLoadWrite(filePath, "VARCHAR UserPassword = 'Okcoge2103b',")
Call FastLoadWrite(filePath, "VARCHAR LogTable = 'DL_OGE_Analytics." & tableName & "_LOG',")
Call FastLoadWrite(filePath, "VARCHAR ErrorTable1 = 'DL_OGE_Analytics." & tableName & "_E1',")
Call FastLoadWrite(filePath, "VARCHAR ErrorTable2 = 'DL_OGE_Analytics." & tableName & "_E2',")
Call FastLoadWrite(filePath, "VARCHAR TargetTable = 'DL_OGE_Analytics." & tableName & "'")
Call FastLoadWrite(filePath, ");")
End Sub
'
' DROP/CREATE Error Tables and Target Table
'
Sub TPT_Setup_Tables()

    Call FastLoadWrite(filePath, "STEP Setup_Tables")
    Call FastLoadWrite(filePath, "(")
    Call FastLoadWrite(filePath, "Apply")
    Call FastLoadWrite(filePath, "('DROP TABLE DL_OGE_Analytics." & tableName & "_LOG;'),")
    Call FastLoadWrite(filePath, "('DROP TABLE DL_OGE_Analytics." & tableName & "_E1;'),")
    Call FastLoadWrite(filePath, "('DROP TABLE DL_OGE_Analytics." & tableName & "_E2;'),")
    Call FastLoadWrite(filePath, "('DROP TABLE DL_OGE_Analytics." & tableName & ";'),")
    Call FastLoadWrite(filePath, "('CREATE TABLE DL_OGE_Analytics." & tableName & "(ID INTEGER, CREATE_AUDIT_KEY INTEGER, DNIS VARCHAR(25));')")
    Call FastLoadWrite(filePath, "TO OPERATOR (od_" & tableName & ");")
    Call FastLoadWrite(filePath, ");")

End Sub
'
' Define LOAD operator to load target table
'
Sub TPT_Load_File()

Call FastLoadWrite(filePath, "STEP LOAD_FILE")
Call FastLoadWrite(filePath, "(")
Call FastLoadWrite(filePath, "Apply")
Call FastLoadWrite(filePath, "('INSERT INTO DL_OGE_Analytics." & tableName & "(EMP_NAME, AGE)")
Call FastLoadWrite(filePath, "Values")
Call FastLoadWrite(filePath, "(:EMP_NAME,:AGE);")
Call FastLoadWrite(filePath, "')")
Call FastLoadWrite(filePath, "TO OPERATOR (ol_" & tableName & ")")
Call FastLoadWrite(filePath, "SELECT * FROM OPERATOR(op_" & tableName & ");")

End Sub


