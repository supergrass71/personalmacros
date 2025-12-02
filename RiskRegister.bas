Attribute VB_Name = "RiskRegister"
Option Explicit

Sub cleanUPRiskRegister()
Dim riskColumns As Range, cell As Range

For Each cell In Range("A1:A1000")

    With ActiveSheet
        If cell.MergeCells Then
            'Range(Cells(1,1),Cells(5,5))
            Set riskColumns = .Range(.Cells(cell.row, 1), .Cells(cell.row, 29))
            riskColumns.UnMerge
        End If
    End With
Next cell
End Sub

Sub RemoveBlankColumns()
Dim cell As Range

For Each cell In Range("A9:A286")
    If IsEmpty(cell.Value) Then cell.EntireRow.Delete
Next cell
End Sub

Sub addSystemID()
Dim cell As Range, othercells As Range

For Each cell In Range("RiskRegisterNumbers")
    For Each othercells In Range("SysInfo")
        If cell.Value = othercells.Value Then
            cell.Offset(0, 4).Value = othercells.Offset(0, 4).Value
            cell.Offset(0, 5).Value = othercells.Offset(0, 5).Value
            GoTo Skip
        End If
    Next othercells
Skip:
Next cell

End Sub
