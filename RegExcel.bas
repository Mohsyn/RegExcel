Attribute VB_Name = "RegExcel"
Public findonly As Boolean
Public pattern As String
Public rng As Range
Public cell As Range
Public regex As Object
Public replacePattern As String
Public matchCount As Integer
Sub doregex()
Attribute doregex.VB_ProcData.VB_Invoke_Func = "q\n14"
regexForm.RegexFind
End Sub
