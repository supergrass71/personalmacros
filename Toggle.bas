Attribute VB_Name = "Toggle"
Option Explicit

Sub ToggleR1C1()
Attribute ToggleR1C1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ToggleR1C1 Macro
'
With Application
'
    If .ReferenceStyle = xlR1C1 Then
        .ReferenceStyle = xlA1
    Else
        .ReferenceStyle = xlR1C1
    End If
End With
End Sub
Sub YellowShading()
Attribute YellowShading.VB_ProcData.VB_Invoke_Func = " \n14"
'
' YellowShading Macro
'

'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
