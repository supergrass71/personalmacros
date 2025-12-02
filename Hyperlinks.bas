Attribute VB_Name = "Hyperlinks"
Option Explicit

Sub ConvertToHyperlink(cell As Range)
Attribute ConvertToHyperlink.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ConvertToHyperlink Macro
'new orbit url
Dim replacementUrl As String

replacementUrl = Replace(cell.Value, "orbit", "orbit.hub.airservicesaustralia.com")
replacementUrl = Replace(replacementUrl, "pdf", "docx")

    ActiveSheet.Hyperlinks.Add Anchor:=cell, Address:= _
        replacementUrl, TextToDisplay:=replacementUrl
End Sub

Sub covertHyper()
Dim cell As Range

For Each cell In Selection
    On Error Resume Next
    Call ConvertToHyperlink(cell)
    On Error GoTo 0
Next cell
End Sub
