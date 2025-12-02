Attribute VB_Name = "SSPCleanUP"
Sub fourdigitControls()
Attribute fourdigitControls.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'
For Each cell In Selection
    If Len(cell.Value) = 3 Then
        cell.Value = "0" & cell.Value
    End If
    MsgBox cell.Value
    If Len(cell.Value) = 2 Then
        cell.Value = "00" & cell.Value
    End If
Next
End Sub

Sub checkControlinTable()

Dim newISMControls As Range, oldISMtable As Range, oldISMTableVertical As Range, vertDestination As Range
Dim obsoleteControlsList As Range, newControlsList As Range, cell As Range
Dim I As Long, lastRow As Long, lastColumn As Long

'define ranges
On Error GoTo Finish 'user cancels with no range
Set newISMControls = Application.InputBox(Prompt:="Choose the top cell in new ISM controls", Title:="New ISM Controls", Type:=8)
With ActiveSheet
    lastRow = .Cells(.Rows.Count, newISMControls.Column).End(xlUp).Row
    Set newISMControls = .Range(.Cells(newISMControls.Row, newISMControls.Column), .Cells(lastRow, newISMControls.Column))
    'MsgBox "newISMcontrols = " & newISMcontrols.Address
End With

'define control table range based on pasting leftmost corner to cell A1, also format the first columns as text
With ActiveSheet
    'handle special case of single row table :)
    If IsEmpty(Range("A2")) Then
        lastRow = 1
    Else
        lastRow = .Cells(1, 1).End(xlDown).Row
    End If
    lastColumn = .Cells(1, 1).End(xlToRight).Column
    Set oldISMtable = .Range(.Cells(1, 1), .Cells(lastRow, lastColumn))
    'format first columns as text
    For I = 1 To lastColumn
    .Columns(I).EntireColumn.NumberFormat = "@"
    Next I
End With

'create/define vertical old ism table column, ensure formatted as text
With ActiveSheet
    Set vertDestination = .Cells(1, newISMControls.Column + 2)
    .Columns(vertDestination.Column).EntireColumn.NumberFormat = "@"
End With

Call verticalise2(oldISMtable, vertDestination)

With ActiveSheet
    lastRow = .Cells(.Rows.Count, vertDestination.Column).End(xlUp).Row
    Set oldISMTableVertical = .Range(.Cells(1, vertDestination.Column), .Cells(lastRow, vertDestination.Column))
End With

'define and label the output starter rows
With ActiveSheet
    Set obsoleteControlsList = .Cells(oldISMtable.Rows.Count + 2, 1)
    Set newControlsList = obsoleteControlsList.Offset(0, 3)
    obsoleteControlsList.Value = "Deleted"
    newControlsList.Value = "New"
End With

'check for obsolete controls (in word table but not in latest ISM)
I = 1
For Each cell In oldISMTableVertical
    If Not IsError(Application.Match(cell.Value, newISMControls, 0)) Then
        GoTo Continue
    Else
        obsoleteControlsList.Offset(I, 0).Value = cell.Value
        I = I + 1
    End If
Continue:
Next cell
obsoleteControlsList.Offset(0, 1).Value = I - 1

'check for new ISM controls that are not in the Word table
I = 1
For Each cell In newISMControls
    If Not IsError(Application.Match(cell.Value, oldISMTableVertical, 0)) Then
        GoTo Continue2
    Else
        newControlsList.Offset(I, 0).Value = cell.Value
        I = I + 1
    End If
Continue2:
Next cell
newControlsList.Offset(0, 1).Value = I - 1
Finish:
On Error GoTo 0
End Sub

Sub verticalise()
I = 1
For Each cell In Selection
    Range("SSP").Offset(I, 0).Value = cell.Value
    I = I + 1
Next
End Sub
Sub verticalise2(sourceRng As Range, destination As Range)
Dim I As Long, cell As Range
Dim controlID As String
I = 0
For Each cell In sourceRng
    'If IsNumeric(cell.Value) Then
    '    controlID = Str(cell.Value)
    'Else
    '    controlID = cell.Value
    'End If
    destination.Offset(I, 0).Value = controlID
    destination.Offset(I, 0).Value = cell.Value

    I = I + 1
Next
End Sub

Sub Textforsure()
For Each cell In Selection
    If Not (Left(cell.Value, 1)) = "'" Then
        cell.Value = "'" & cell.Value
    End If
Next
End Sub
Sub TrimSingleQuote()
For Each cell In Selection
    If Left(cell.Value, 1) = "'" Then
        cell.Value = Mid(cell.Value, 2, Len(cell.Value - 1))
    End If
Next
End Sub
Sub checkControlinTable2()

I = 1
For Each cell In Range("Control_Table")
    If Not IsError(Application.Match(cell.Value, Range("Cryptocontrols"), 0)) Then
        GoTo Continue
    Else
        Range("Deleted").Offset(I, 0).Value = cell.Value
        I = I + 1
    End If
Continue:
Next cell
Range("Deleted").Offset(0, 1).Value = I - 2
I = 1
For Each cell In Range("Cryptocontrols")
    If Not IsError(Application.Match(cell.Value, Range("CryptoTableVertical"), 0)) Then
        GoTo Continue2
    Else
        Range("New_Controls").Offset(I, 0).Value = cell.Value
        I = I + 1
    End If
Continue2:
Next cell
Range("New_Controls").Offset(0, 1).Value = I - 2
End Sub

