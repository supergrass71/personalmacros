Attribute VB_Name = "Validation"
Option Explicit

Sub AddCondtionalFormatting()
Attribute AddCondtionalFormatting.VB_ProcData.VB_Invoke_Func = " \n14"
'
' AddCondtionalFormatting Macro
'

'
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="ML1,ML2,ML3"
        '.IgnoreBlank = True
        '.InCellDropdown = True
        '.InputTitle = ""
        '.ErrorTitle = ""
        '.InputMessage = ""
        '.ErrorMessage = ""
        '.ShowInput = True
        '.ShowError = True
    End With
End Sub


