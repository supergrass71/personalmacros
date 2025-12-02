Attribute VB_Name = "Autofit"
Option Explicit

Sub Autofit_Selected()
Attribute Autofit_Selected.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Autofit_Selected Macro
'
Dim cell As Range

For Each cell In Selection
    If Selection.Columns.Count = 1 Then 'single column
        cell.EntireRow.Autofit
    End If
    If Selection.Rows.Count = 1 Then
        cell.EntireColumn.Autofit
    End If
    If Selection.Rows.Count > 1 And Selection.Columns.Count > 1 Then
        cell.EntireRow.Autofit
        cell.EntireColumn.Autofit
    End If
Next cell
End Sub

