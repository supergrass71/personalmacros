Attribute VB_Name = "MoreComments"
Option Explicit

Sub AddComment()
Attribute AddComment.VB_ProcData.VB_Invoke_Func = "t\n14"
'
' AddComment just shows the cells I have changed
' updated to offer optional common comment text

Dim cmt As Comment
Dim commentCells As Range, cell As Range
Dim commentTime As String, userName As String
Dim commentText As String

userName = GetUserFullName
commentTime = Format(Now(), "dd/mm/yy hh:mm AM/PM")

Set commentCells = Selection

commentText = Application.InputBox(Prompt:="Add text to all comment cells?", Title:="Add Comments to Selected Cells", Type:=2)

If commentText = "False" Then commentText = "" 'result of clicking Cancel on InputBox

For Each cell In commentCells
    Set cmt = cell.Comment
    If cmt Is Nothing Then
        cell.AddComment text:=userName & vbLf & commentTime & vbLf & commentText
        Set cmt = cell.Comment
        Call reset_Comment_size(cmt) 'see below, to make comment fit better
        Set cmt = Nothing
    End If

Next cell

End Sub

Sub reset_box_size()
'https://stackoverflow.com/questions/45515769/resize-excel-comments-to-fit-text-with-specific-width
Dim pComment As Comment
Dim lArea As Double
For Each pComment In Application.ActiveSheet.Comments
    With pComment.Shape

        .TextFrame.AutoSize = True

        lArea = .Width * .Height
        
        'only resize the autosize if width is above 300
        If .Width > 300 Then .Height = (lArea / .Width)       ' used .width so that it is less work to change final width

        
        .TextFrame.AutoMargins = False
        .TextFrame.MarginBottom = 0      ' margins need to be tweaked
        .TextFrame.MarginTop = 0
        .TextFrame.MarginLeft = 0
        .TextFrame.MarginRight = 0
        End With
Next

End Sub

Sub reset_Comment_size(pComment As Comment)
Dim lArea As Double
With pComment.Shape

    .TextFrame.AutoSize = True

    lArea = .Width * .Height
    
    'only resize the autosize if width is above 300
    If .Width > 300 Then .Height = (lArea / .Width)       ' used .width so that it is less work to change final width
    .TextFrame.AutoMargins = False
    .TextFrame.MarginBottom = 0      ' margins need to be tweaked
    .TextFrame.MarginTop = 0
    .TextFrame.MarginLeft = 0
    .TextFrame.MarginRight = 0
End With

End Sub
