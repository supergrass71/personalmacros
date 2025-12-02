Attribute VB_Name = "Filter"
Option Explicit

Sub ModifiedAutoFilter()

Dim rngFilter As Range, rngTopLeft As Range

Dim lastRow As Long, LastCol As Long


Set rngTopLeft = ActiveCell

With ActiveSheet
    If .AutoFilterMode Then .AutoFilterMode = False
    lastRow = .Cells(.Rows.Count, rngTopLeft.Column).End(xlUp).Row
    LastCol = .Cells(rngTopLeft.Row, .Columns.Count).End(xlToLeft).Column
    Set rngFilter = .Range(.Cells(rngTopLeft.Row, rngTopLeft.Column), .Cells(lastRow, LastCol))
    'MsgBox rngFilter.Address
    rngFilter.AutoFilter
End With


End Sub
Function fnModifiedAutoFilter() As Range

Dim rngFilter As Range, rngTopLeft As Range

Dim lastRow As Long, LastCol As Long

If Len(UserForm2.TBLeftCornerCell.Value) = 0 Then
    Set rngTopLeft = ActiveCell
Else
    Set rngTopLeft = Range(UserForm2.TBLeftCornerCell.Value)
End If

With ActiveSheet
    If .AutoFilterMode Then .AutoFilterMode = False
    lastRow = .Cells(.Rows.Count, rngTopLeft.Column).End(xlUp).Row
    LastCol = .Cells(rngTopLeft.Row, .Columns.Count).End(xlToLeft).Column
    Set rngFilter = .Range(.Cells(rngTopLeft.Row, rngTopLeft.Column), .Cells(lastRow, LastCol))
    'MsgBox rngFilter.Address
    Set fnModifiedAutoFilter = rngFilter
End With

End Function
