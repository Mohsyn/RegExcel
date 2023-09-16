VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} regexForm 
   Caption         =   "Regex Find and Replace by Mohsin"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7080
   OleObjectBlob   =   "regexForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "regexForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'### v1.1
'Author M Mohsyn


Sub RegexFind()

    If Selection.count > 1 Then
        Set rng = Selection
    Else
        Set rng = ActiveSheet.UsedRange
    End If
    Set regex = CreateObject("VBScript.RegExp")
    regexForm.Show

End Sub


Private Sub btnFind_Click()
findonly = True
setvariables

End Sub

Private Sub btnReplace_Click()
findonly = False
setvariables
End Sub


Private Sub setvariables()
regex.Global = chkGlobal.Value
regex.ignorecase = chkIgnoreCase.Value
replacePattern = txtReplace.text
pattern = txtPattern.text

If txtPattern <> "" Then
    regex.pattern = pattern
    regexForm.Hide
    regexProcess
Else
    MsgBox ("Pattern not provided")
End If

End Sub

Private Sub regexProcess()

    If Not findonly Then matchCount = 0
    
    For Each cell In rng
          If cell.Row < ActiveCell.Row Or (cell.Row = ActiveCell.Row And cell.Column <= ActiveCell.Column) Then
            GoTo skipiteration
        End If
        If regex.Test(cell.Value) Then
            If findonly Then
                cell.Activate
                matchCount = matchCount + 1
                Exit Sub
            Else
                Debug.Print (regex.Execute(cell.Value).count)
                matchCount = matchCount + regex.Execute(cell.Value).count
                 cell.Value = regex.Replace(cell.Value, replacePattern)
            End If
        End If
skipiteration:
    Next
    
    If findonly Then
        MsgBox ("Finished Search and Replace" & vbCrLf & matchCount & " occurences found")
    
    Else
        MsgBox ("Finished Search and Replace" & vbCrLf & matchCount & " occurences replaced")
    End If
     matchCount = 0
    
End Sub


Private Sub UserForm_Activate()
txtPattern.text = pattern
txtReplace.text = replacePattern
lblRange.Caption = rng.Address
End Sub

