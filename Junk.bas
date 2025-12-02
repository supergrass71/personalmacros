Attribute VB_Name = "Junk"
Sub applicable()
For Each cell In Selection
    Result = Int((2 - 1 + 1) * Rnd + 1)
        If Result = 1 Then
            cell.Value = "Yes"
        Else
            cell.Value = "No"
        End If
Next cell

End Sub

Sub randomtext()
Dim wd As Object
Set wd = GetObject(, "Word.Application")
If wd Is Nothing Then
    Set wd = CreateObject("Word.Application")
End If

ActiveCell.Value = wd.rand(3, 4)
Set wd = Nothing
End Sub
