Attribute VB_Name = "Comments"
Sub SRAComments()
With ActiveSheet
    For Each cell In .Range("Q2:Q165")
    'For Each cell In Selection
        If cell.Offset(0, -13).Style = "Neutral" And cell.Offset(0, -13).Value = "Malicious" Then
            cell.Value = cell.Value & vbCrLf & "Swap " & Chr(34) & "Deliberate" & Chr(34) _
                & "for" & Chr(34) & "Malicious" & Chr(34)
        End If
    Next cell
End With
End Sub


