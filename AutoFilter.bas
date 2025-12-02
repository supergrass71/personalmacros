Attribute VB_Name = "AutoFilter"
Option Explicit

Sub ISMAF()
Attribute ISMAF.VB_ProcData.VB_Invoke_Func = "j\n14"
Load UserForm1
UserForm1.Show
End Sub
Sub AutofilterHere()
Dim topLHScell As Range, topLHSrow As Integer, topLHScolumn As Integer
Dim autofilterRange As Range
Dim lastRow As Long
Dim lastColumn As Long

Set topLHScell = ActiveCell
topLHSrow = topLHScell.row
topLHScolumn = topLHScell.column

lastRow = getLastRow2(topLHScolumn)
lastColumn = LastColumnInOneRow(topLHSrow)

If lastRow > 10000 Or lastColumn > 10000 Then
    MsgBox "Please check your anchor cell!"
    Exit Sub
End If

With ActiveSheet
    .Range("A1").AutoFilter 'reset any active filters
    Set autofilterRange = .Range(.Cells(topLHSrow, topLHScolumn), .Cells(lastRow, lastColumn))
    autofilterRange.AutoFilter
End With
End Sub

Sub FilterByISMControl()
Attribute FilterByISMControl.VB_ProcData.VB_Invoke_Func = "g\n14"
'
' FilterByISMControl Macro
'
Dim controlNumber As String
Dim rng As Range, cell As Range
Dim visibleCount As Long, lastVisibleRow As Long
Dim firstrow As Integer, lastRow As Long
Dim prefix As String 'modify according to ISM control scheme
'
'change first row to avoid hiding heading row.
firstrow = 2
lastRow = getLastRow

'is there a prefix to the ism control?
prefix = "ISM-"

Selection.AutoFilter
    
controlNumber = Application.InputBox(Prompt:="Enter ISM Control Number", Title:="Find ISM Control", Type:=2)
If controlNumber = "" Then
    Selection.AutoFilter
    Exit Sub
End If
    ActiveSheet.Range("$A$" & firstrow & ":$AA$" & lastRow).AutoFilter Field:=4, Criteria1:=prefix & controlNumber
'/* fix later with https://jkp-ads.com/articles/apideclarations.aspx
GoTo SkipCode
Set rng = ActiveSheet.Range("A1").CurrentRegion
' Loop from bottom to top to find the last visible row
For Each cell In rng.Columns(1).Cells
    If Not cell.EntireRow.Hidden Then
        lastVisibleRow = cell.row
    End If
Next cell

'add filtered result to clipboard
If lastVisibleRow > 1 Then
    SetClipboardText (rng.Cells(lastVisibleRow, 11).Value)
Else
    MsgBox "Control not found!"
End If
'*/
SkipCode:
End Sub
Sub SetClipboardText(ByVal text As String)
    Dim clipboard As Object

    ' Create a new DataObject
    Set clipboard = CreateObject("MSForms.DataObject")

    ' Set the text to the clipboard
    clipboard.SetText text
    clipboard.PutInClipboard
End Sub

Sub FindLastVisibleRow()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastVisibleRow As Long

    Set ws = ThisWorkbook.Sheets("Sheet1")
    Set rng = ws.Range("A1").CurrentRegion ' Adjust to your data range

    If ws.AutoFilterMode Then
        ' Loop from bottom to top to find the last visible row
        For Each cell In rng.Columns(1).Cells
            If Not cell.EntireRow.Hidden Then
                lastVisibleRow = cell.row
            End If
        Next cell

        If lastVisibleRow > 0 Then
            MsgBox "Last visible row is: " & lastVisibleRow
        Else
            MsgBox "No visible rows found."
        End If
    Else
        MsgBox "AutoFilter is not applied.", vbExclamation
    End If
End Sub

