Attribute VB_Name = "ISM_Stuff"
Dim existingISMControls As String
Dim newISMControls As String
Option Explicit

Sub addZeros()
Dim cell As Range

For Each cell In Selection
    If Len(cell.Value) = 1 Then cell.Value = "000" & cell.Value
    If Len(cell.Value) = 2 Then cell.Value = "00" & cell.Value
    If Len(cell.Value) = 3 Then cell.Value = "0" & cell.Value
Next cell

End Sub

Sub checkUpdatedControls()
Dim rngNewControls As Range, rngExistingControls As Range, newCell As Range, existCell As Range
Dim resultsWs As Worksheet
Dim I As Integer
Dim existingISMControls As String, newISMControls As String

'set worksheet names to compare (overwritten below! :)
'__________________________________________________________________________________________

existingISMControls = "01 December 2022"
newISMControls = "Maturity Model"

'__________________________________________________________________________________________

I = 0
'MsgBox "select first cell of new controls" & vbLf & _
'        "followed by first cell of old controls"
        
'Set newControls = getRange 'refer Utilities page
'Set existingControls = getRange

Set rngNewControls = ActiveWorkbook.Sheets("Delta Assessment March 2023").Range("C2:C80")
Set rngExistingControls = ActiveWorkbook.Sheets("01 December 2022").Range("D3:D852")
'If newControls.Rows.Count = 1 Then Exit Sub

'Set resultsWs = Sheets.Add
'resultsWs.Name = "comparison"

For Each newCell In rngNewControls
'MsgBox "newcell address entering loop!" & newCell.Address
'If newCell.Offset(0, 1).Value = 0 Then ' brand new contol
'    resultsWs.Range("A1").Offset(i, 0).Value = newCell.Value 'control number
'    resultsWs.Range("A1").Offset(i, 1).Value = newCell.Offset(0, 8).Value 'updated control desc
'    resultsWs.Range("A1").Offset(i, 2).Value = "-" 'new control!
'    resultsWs.Range("A1").Offset(i, 3).Value = "this is a new March 2023 control" 'comments
'    i = i + 1
'    GoTo Skip ' new control
'End If
    'control already exists, match to current comments
    For Each existCell In rngExistingControls
        If existCell.Value = newCell.Value Then
           ' MsgBox "existcell:= " & existCell.Address & vbLf & _
           '         "newcell:= " & newCell.Address
            
            'resultsWs.Range("A1").Offset(i, 0).Value = newCell.Value 'control number
            'resultsWs.Range("A1").Offset(i, 1).Value = newCell.Offset(0, 8).Value 'updated control desc
            'resultsWs.Range("A1").Offset(i, 2).Value = existCell.Offset(0, 8).Value 'exist control desc
            'resultsWs.Range("A1").Offset(i, 3).Value = existCell.Offset(0, 14).Value 'comments
            'i = i + 1
            newCell.Offset(0, 9).Value = existCell.Offset(0, 11).Value 'transfer applicability status
            GoTo Skip
        End If
    Next existCell

Skip:
Next newCell


End Sub
Sub checkUpdatedControls2()
Dim rngNewControls As Range, rngExistingControls As Range, newCell As Range, existCell As Range
Dim resultsWs As Worksheet
Dim I As Integer, j As Integer
Dim oldImplementation As String, newImplementation As String

With Application
    .ScreenUpdating = False
    .ErrorCheckingOptions.BackgroundChecking = False
End With
'set worksheet names to compare
'__________________________________________________________________________________________

existingISMControls = "Delta Assessment March 2023"
newISMControls = "Delta Asssessment June 2023"

'__________________________________________________________________________________________

I = 1
        
Set resultsWs = Sheets.Add
resultsWs.Name = "Changed ISM controls"

With resultsWs
    .Range("A1").Value = "Control"
    .Range("B1").Value = "Implementation Status"
    .Range("C1").Value = Format(Now(), "dd/MM/yy")
    .Columns("A:A").NumberFormat = "@" 'text format for control values
End With


For j = 4 To 853
    
    newImplementation = ActiveWorkbook.Sheets(newISMControls).Range("N" & j).Value
    If newImplementation = "In Implementation" Then GoTo Skip 'don't want to know
    
    oldImplementation = ActiveWorkbook.Sheets(existingISMControls).Range("N" & j).Value
    
    If Not newImplementation = oldImplementation Then
        resultsWs.Range("A1").Offset(I, 0).Value = ActiveWorkbook.Sheets(newISMControls).Range("D" & j).Value
        resultsWs.Range("A1").Offset(I, 1).Value = newImplementation
            I = I + 1
    End If

Skip:
Next j

resultsWs.Columns("B:B").EntireColumn.AutoFit
resultsWs.Columns("C:C").EntireColumn.AutoFit

With Application
    .ScreenUpdating = True
    .ErrorCheckingOptions.BackgroundChecking = True
End With

End Sub

Sub UpdatedNewControls()
Dim rngNewControls As Range, rngExistingControls As Range, newCell As Range, existCell As Range
Dim resultsWs As Worksheet
Dim I As Integer, j As Integer
Dim MarRevision As String, MarDescription As String, controlNumber As String

With Application
    .ScreenUpdating = False
End With
'set worksheet names to compare
'__________________________________________________________________________________________

existingISMControls = "December 2022"
newISMControls = "Delta Assessment March 2023"

'__________________________________________________________________________________________

Set rngNewControls = ActiveWorkbook.Sheets(newISMControls).Range("D2:D80")

I = 1
'
'set pdate based on revision and date
For j = 4 To 853
    Set existCell = ActiveWorkbook.Sheets(existingISMControls).Range("D" & j)
    controlNumber = ActiveWorkbook.Sheets(existingISMControls).Range("D" & j).Value
    For Each newCell In rngNewControls
        If controlNumber = newCell.Value Then
            MarRevision = newCell.Offset(0, 1).Value
            MarDescription = newCell.Offset(0, 8).Value
            If Not (MarRevision = "0") Then ' this is not a new control, update revision and description
                  existCell.Offset(0, 1).Value = MarRevision
                  existCell.Offset(0, 2).Value = "Mar-23" 'updated
                  existCell.Offset(0, 8).Value = MarDescription 'updated description
            End If
        End If
    Next newCell
Next j


With Application
    .ScreenUpdating = True
End With

End Sub

Sub cleanupImplementation()

Dim cell As Range, rngImplementation As Range

Set rngImplementation = Range("P3:P852")

For Each cell In rngImplementation

    If InStr(1, cell.Value, Chr(32), vbTextCompare) > 0 Then
        cell.Value = "TBC"
    End If
Next cell
End Sub

Sub makeRowNA()
Attribute makeRowNA.VB_ProcData.VB_Invoke_Func = "o\n14"
Dim ISMRow As Range

With ActiveSheet()
Set ISMRow = .Range(.Cells(ActiveCell.Row, ActiveCell.Column), .Cells(ActiveCell.Row, 20))
.Range("O" & ActiveCell.Row).Value = "MN"
End With
ISMRow.Style = "Not Applicable"

End Sub
Sub makeRowNACell(cell As Range)

Dim ISMRow As Range

With ActiveSheet()
Set ISMRow = .Range(.Cells(cell.Row, cell.Column), .Cells(cell.Row, 21))
'.Range("O" & cell.Row).Value = "MN"
End With
ISMRow.Style = "Not Applicable"

End Sub

Sub ApplyNA()
Dim cell As Range

'check A column is selected

If Not Selection.Column = 1 Then
    MsgBox "you must select Column A!"
    Exit Sub
End If

For Each cell In Selection
    If cell.Offset(0, 12).Value = "Not Applicable" Then
        Call makeRowNACell(cell)
    End If

Next cell

End Sub

Sub ISMDeleteBlankRows()
Dim rowRange As Range, cell As Range

Set rowRange = Range("A2:A880")

For Each cell In rowRange
    If IsEmpty(cell.Value) Then
        cell.EntireRow.Delete
    End If
Next cell
End Sub

Sub removeZero()
Dim cell As Range

For Each cell In Selection
    If InStr(1, cell.Value, "0") = 1 Then
        cell.Value = Replace(cell.Value, "0", "", 1, 1)
    End If
    cell.Value = Int(cell.Value)
Next cell
End Sub
Sub fixLeadingQuote()
If Range("D79").Value = Range("D81").Value Then
MsgBox "true"
Else
MsgBox "false"
End If
End Sub
Sub checkUpdatedControls21()
Dim rngNewControls As Range, rngExistingControls As Range, newCell As Range, existCell As Range
Dim resultsWs As Worksheet
Dim I As Integer, j As Integer
Dim oldImplementation As String, newImplementation As String

With Application
    .ScreenUpdating = False
    .ErrorCheckingOptions.BackgroundChecking = False
End With
'set worksheet names to compare
'__________________________________________________________________________________________

existingISMControls = "Delta Assessment March 2023"
newISMControls = "Delta Asssessment June 2023"

'__________________________________________________________________________________________

I = 1
        
Set resultsWs = Sheets.Add
resultsWs.Name = "Changed ISM controls"

With resultsWs
    .Range("A1").Value = "Control"
    .Range("B1").Value = "Implementation Status"
    .Range("C1").Value = Format(Now(), "dd/MM/yy")
    .Columns("A:A").NumberFormat = "@" 'text format for control values
End With


For j = 4 To 853
    
    newImplementation = ActiveWorkbook.Sheets(newISMControls).Range("N" & j).Value
    If newImplementation = "In Implementation" Then GoTo Skip 'don't want to know
    
    oldImplementation = ActiveWorkbook.Sheets(existingISMControls).Range("N" & j).Value
    
    If Not newImplementation = oldImplementation Then
        resultsWs.Range("A1").Offset(I, 0).Value = ActiveWorkbook.Sheets(newISMControls).Range("D" & j).Value
        resultsWs.Range("A1").Offset(I, 1).Value = newImplementation
            I = I + 1
    End If

Skip:
Next j

resultsWs.Columns("B:B").EntireColumn.AutoFit
resultsWs.Columns("C:C").EntireColumn.AutoFit

With Application
    .ScreenUpdating = True
    .ErrorCheckingOptions.BackgroundChecking = True
End With

End Sub

Sub UpdatedNewControls3()
Dim rngNewControls As Range, rngExistingControls As Range, newCell As Range, existCell As Range
Dim resultsWs As Worksheet
Dim I As Integer, j As Integer
Dim MarRevision As String, MarDescription As String, controlNumber As String

With Application
    .ScreenUpdating = False
End With
'set worksheet names to compare
'__________________________________________________________________________________________

existingISMControls = "checklist"
newISMControls = "Delta Asssessment June 2023"
'added loop break (goto Skip)

'__________________________________________________________________________________________

Set rngNewControls = ActiveWorkbook.Sheets(newISMControls).Range("D2:D45")
Set rngExistingControls = ActiveWorkbook.Sheets(existingISMControls).Range("D4:D853")

For Each newCell In rngNewControls
If Not IsEmpty(newCell.Offset(0, 10)) Then GoTo Skip
    For Each existCell In rngExistingControls
        If newCell.Value = existCell.Value Then
            newCell.Offset(0, 9).Value = existCell.Offset(0, 11).Value 'applicability
            newCell.Offset(0, 10).Value = existCell.Offset(0, 12).Value 'implementation status
            newCell.Offset(0, 11).Value = "MN"
            newCell.Offset(0, 14).Value = existCell.Offset(0, 15).Value 'implementation comments
        End If
    Next existCell
Skip:
Next newCell

With Application
    .ScreenUpdating = True
End With

End Sub
Sub compareNewToOld()

Dim rngNewControl As Range, rngExistingControls As Range, cell As Range
Dim lookupControl As String, newControlDescription As String, compareDescription As String
'Dim Answer As VbMsgBoxResult

'__________________________________________________________________________________________

existingISMControls = "01 December 2022"
newISMControls = "Delta Asssessment June 2023"

'__________________________________________________________________________________________

Set rngNewControl = ActiveWorkbook.Sheets(newISMControls).Range("D" & ActiveCell.Row)
Set rngExistingControls = ActiveWorkbook.Sheets(existingISMControls).Range("D3:D852")

newControlDescription = rngNewControl.Offset(0, 8).Value

If Len(rngNewControl.Value) = 3 Then
    lookupControl = Str("0" & rngNewControl.Value)
    MsgBox VarType(lookupControl)
Else
    lookupControl = Str(rngNewControl.Value)
End If

For Each cell In rngExistingControls

    If lookupControl = cell.Value Then
        compareDescription = "NEW:" & vbLf & newControlDescription & vbLf & vbLf & _
        "DEC22:" & vbLf & cell.Offset(0, 10).Value
        GoTo Skip
    Else
        compareDescription = "NEW CONTROL"
        SetClipboard (compareDescription)
    End If
    
Next

Skip:
MsgBox2 compareDescription
End Sub

Sub updateChangesColumn()
Attribute updateChangesColumn.VB_ProcData.VB_Invoke_Func = "f\n14"
Dim changes As String, changeCell As Range

Set changeCell = Cells(ActiveCell.Row, 21)

changes = Application.InputBox(Prompt:="Explain changes to control", Title:="Control changes", Default:="NEW CONTROL")

changeCell.Value = changes
End Sub
