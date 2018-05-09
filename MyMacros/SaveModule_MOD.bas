Attribute VB_Name = "SaveModule_MOD"


Sub line_number(Optional strModuleName) ' As String)

Dim vbProj As VBProject
Dim vbComp As VBComponent
Dim cmCode As CodeModule
Dim intLine As Integer

Set vbProj = Application.VBE.ActiveVBProject
Set vbComp = vbProj.VBComponents(strModuleName)
Set cmCode = vbComp.CodeModule

For intLine = 2 To cmCode.CountOfLines - 1
   cmCode.InsertLines intLine, intLine - 1 & cmCode.Lines(intLine, 1)
   cmCode.DeleteLines intLine + 1, 1
Next intLine

End Sub
