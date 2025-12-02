Attribute VB_Name = "Comments"
Option Explicit
Dim commentText As String

Sub addcomments()
Attribute addcomments.VB_ProcData.VB_Invoke_Func = "t\n14"
Dim cell As Range
Dim commentText As Variant

commentText = Application.InputBox(Prompt:="Set comment text:", Title:="Add comments to selected cells", Type:=2)
If commentText = False Then Exit Sub
    For Each cell In Selection
        If Not cell.CommentThreaded Is Nothing Then cell.ClearComments
        cell.AddCommentThreaded (commentText)
    Next cell
End Sub
Sub DeleteComments()
Dim cell As Range

For Each cell In Selection
    If Not cell.CommentThreaded Is Nothing Then cell.ClearComments
Next cell
End Sub


Sub addcomments2(cell As Range, comment As String)

If Not cell.CommentThreaded Is Nothing Then cell.ClearComments
cell.AddCommentThreaded (comment)
End Sub

Sub testADDComms()
    Call addcomments2(ActiveCell, "blah,blah")
End Sub

Sub AddControlDescAsComment()
Dim cell As Range, cellA As Range

For Each cell In Range("ISMCtrlList")
    commentText = cell.Offset(0, 11).Value
    For Each cellA In Range("ISM_Review_Controls")
        If cellA.Value = cell.Value Then
            'cell.Offset(0, 8).Value = cellA.Offset(0, 8).Value 'implementation
            Call addcomments2(cellA, commentText)
            commentText = ""
            GoTo Skip
        End If
    Next cellA
Skip:
Next cell
End Sub
