Attribute VB_Name = "Utilities"
Option Explicit

Sub ClearAllFormatting()
Attribute ClearAllFormatting.VB_ProcData.VB_Invoke_Func = " \n14"
'
Dim Answer As Integer
Application.ScreenUpdating = False
'single cell means delete everything, otherwise delete only selected cells
If (Selection.Rows.Count = 1) And (Selection.Columns.Count = 1) Then
    Answer = MsgBox(Prompt:="Delete everything on sheet?", Title:="Clear All Cells!", Buttons:=vbYesNo)
    If Answer = vbNo Then Exit Sub
    Cells.Select
End If

With Selection
    .Borders(xlDiagonalDown).LineStyle = xlNone
    .Borders(xlDiagonalUp).LineStyle = xlNone
    .Borders(xlEdgeLeft).LineStyle = xlNone
    .Borders(xlEdgeTop).LineStyle = xlNone
    .Borders(xlEdgeBottom).LineStyle = xlNone
    .Borders(xlEdgeRight).LineStyle = xlNone
    .Borders(xlInsideVertical).LineStyle = xlNone
    .Borders(xlInsideHorizontal).LineStyle = xlNone
End With
With Selection.Interior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With
    Selection.ClearContents

    Application.ScreenUpdating = True
End Sub

Sub EnumerateSheets()
Dim ws As Worksheet
Dim I As Integer
Dim startCell As Range

Set startCell = ActiveCell
I = 0
For Each ws In Worksheets
    startCell.Offset(I, 0).Value = ws.Name
    I = I + 1
Next ws
End Sub

Function getLastRow() As Long
Dim activeCol As Integer
activeCol = ActiveCell.Column
With ActiveSheet
    getLastRow = .Cells(.Rows.Count, activeCol).End(xlUp).Row
End With
End Function

Function getRange() As Range

Dim startCell As Range

With ActiveSheet
    Set startCell = Application.InputBox(Prompt:="Select start of Range", Title:="Range Selector", _
                            Type:=8)
    If startCell Is Nothing Then Set getRange = ActiveCell
    
    Set getRange = .Range(.Cells(startCell.Row, startCell.Column), .Cells(getLastRow, startCell.Column))
    
    End With
End Function

Sub CleanUP()
Dim cell As Range, pos As Long
Dim rangeToClean As Range

Set rangeToClean = Range("A2:A124")

For Each cell In rangeToClean

    'pos = InStr(1, cell.Value, "par", vbTextCompare)
    '____________________________________________
    'trim contents in place
    'If pos > 0 Then
    '    cell.Value = Left(cell.Value, pos - 2)
    'End If
    '_____________________________________________
    
    'delete blank rows
    If IsEmpty(cell.Value) = 0 Then cell.EntireRow.Delete
    '_____________________________________________
    

Next cell

End Sub

Sub AdjustRowsOrColumns()
'
' AdjustRowsOrColumns Macro
Dim cell As Range
'
For Each cell In Selection.Columns(1)
'
    cell.EntireRow.AutoFit
Next cell
End Sub

