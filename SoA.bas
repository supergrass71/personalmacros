Attribute VB_Name = "SoA"
Option Explicit
Sub AddConditionalFormatting()
Dim endRow As Long, startCell As Range, conditionalRange As Range
endRow = LastRowInOneColumn() 'see bits and pieces

startCell = Application.InputBox(Prompt:="Select first cell of Implementation Status Column", Title:="Cell conditional formatter", Type:=8)
If startCell Is Nothing Then Exit Sub

With ActiveSheet

    Set conditionalRange = .Range(.Cells(startCell.row, startCell.column), .Cells(endRow, startCell.column))
    

End With


End Sub
Sub ConvertStatus()
Dim cell As Range
For Each cell In Range("VSS_Status")
    Select Case cell.Value
        Case Is = "Compliant"
            cell.Value = "Effective"
        Case Is = "Non-Compliant"
            cell.Value = "Not Effective"
        Case Is = "Not Compliant"
            cell.Value = "Not Effective"
        Case Is = "Not Implmented"
            cell.Value = "Not Implemented"
        Case Is = "?"
            cell.Value = "No Visibility"
        Case Is = "TBD"
            cell.Value = "No Visibility"
        Case Is = Chr(63) & " Baseline " & Chr(63)
            cell.Value = "Inherited"
        Case Is = "Partially Compliant"
            cell.Value = "Partially Effective"
    End Select

Next cell


End Sub
Sub CopyFromAutoFilterResults()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cellToCopy As Range
    Dim destCell As Range

    Set ws = ThisWorkbook.Sheets("Sheet1")
    Set rng = ws.Range("A1").CurrentRegion ' Adjust as needed
    Set destCell = ws.Range("B1") ' Destination cell for copied value

    ' Check if AutoFilter is applied
    If ws.AutoFilterMode Then
        ' Loop through visible cells in column A, skipping the header
        For Each cellToCopy In rng.Columns(1).Cells
            If Not cellToCopy.EntireRow.Hidden And cellToCopy.row > rng.Cells(1, 1).row Then
                destCell.Value = cellToCopy.Value
                Exit For ' Only copy the first visible cell
            End If
        Next cellToCopy
    Else
        MsgBox "No AutoFilter applied.", vbExclamation
    End If
End Sub

Sub mar25toASABL()
Dim cell As Range, cellA As Range

For Each cell In Range("RVCOTControls")
    For Each cellA In Range("ISM_Review_Controls")
        If cellA.Value = cell.Value Then
            cell.Offset(0, 8).Value = cellA.Offset(0, 8).Value 'implementation
            GoTo Skip
        End If
    Next cellA
Skip:
Next cell
End Sub
Sub MarkRemainingControls()
Dim cell As Range, cellA As Range

For Each cell In Range("ISM2021_Controls")
    For Each cellA In Range("Addresed_Controls")
        If cellA.Value = cell.Value Then
            GoTo Skip
        End If
    Next cellA
    'if not found, this is one we need to check!
    'see whether the responsible entity is Shared or Frequentis
    If cell.Offset(0, 8).Value = "Shared" Or cell.Offset(0, 8).Value = "Frequentis" Then
        With cell.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
        End With
    End If
Skip:
Next cell
End Sub
Sub prependISM()
Dim cell As Range
For Each cell In Range("Old_Controls")
    cell.Value = "ISM-" & cell.Value
Next cell
End Sub
Sub addVerification()
Dim cell As Range

For Each cell In Range("Verification")
    If cell.Offset(-1, 0).Value = "Implemented" And cell.Offset(-2, 0).Value = "ASA" Then
        cell.Value = "Effective"
    End If
Next cell
End Sub
Sub compareSAAB_to_ML()
Dim cell As Range, cell2 As Range, saab As Range, ASD As Range
With ActiveSheet
  Set saab = .Range("SAAB_ML")
  Set ASD = Range("ISM_ML")
  
  For Each cell In saab
    For Each cell2 In ASD
        If cell2.Value = cell.Value Then
            GoTo Skip
        End If
    Next cell2
    cell.Offset(0, 2).Value = "Yes"
Skip:
    Next cell
End With
End Sub
Sub appendSingleQuote()
Dim cell As Range
For Each cell In Selection
    If Not (Left(cell.Value, 1) = "'") Then
        cell.Value = "'" & cell.Value
    End If
Next
End Sub
Sub testUserForm1()
With UserForm1
    Load UserForm1
    .Show
End With
End Sub
Sub hideSoAColumns()
'hides/unhides Columns E to N with all the revision, classification, ML 1-3
Dim i As Integer
Dim reveal As Boolean
i = 5
If Columns(5).Hidden = True Then
    reveal = False
Else
    reveal = True
End If
While i < 15
    Columns(i).Hidden = reveal
    i = i + 1
Wend

End Sub
Sub testMLstats()
Dim cell As Range, applicable As Range, implemented As Range
With ActiveSheet

Set applicable = .Range("E2:E153")
Set implemented = .Range("G2:G153")


    For Each cell In applicable
        If oneOrTwo = 1 Then
            cell.Value = "Yes"
        Else
            cell.Value = "No"
        End If
    Next cell
    For Each cell In implemented
        If oneOrTwo = 1 Then
            cell.Value = "Yes"
        Else
            cell.Value = "No"
        End If
    Next cell
End With
End Sub
Sub MLDropdown(maturitycombo As String, cell As Range)
Select Case maturitycombo
    Case Is = "YesYesYes"
        With cell.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="ML1,ML2,ML3"
        End With
    Case Is = "NoYesYes"
        With cell.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="ML2,ML3"
        End With
    Case Is = "NoNoYes"
        With cell.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="ML3"
        End With
    Case Else
        'don't add validation
End Select
End Sub

Sub testMLDropdown()
Dim MLcombo As String, cell As Range

For Each cell In Range("ISM_DEC24")
    With cell
        MLcombo = .Offset(0, 8).Value & .Offset(0, 9).Value & .Offset(0, 10).Value
    End With
    Call MLDropdown(MLcombo, cell.Offset(0, 23))
Next cell

End Sub

Sub fixZeros()
Dim asaBaselinecontrols As Range, cell As Range

Set asaBaselinecontrols = Range("D2:D911")
With ActiveWorkbook
    For Each cell In asaBaselinecontrols
    
    If cell.Value = "0" Then cell.ClearContents
    If cell.Offset(0, -4).Value = "No" Then
        If IsEmpty(cell.Offset(0, -1).Value) Then
            cell.Offset(0, 1).Value = "20241029 MN please add a reason for N/A"
            cell.Offset(0, 1).Interior.Color = 65535
        End If
    End If
    Next cell
End With
End Sub

Sub appendZeros()
Dim cell As Range, identifiers As Range

'Set identifiers = ActiveSheet.Range("D2:D" & LastRowInOneColumn(4))
Set identifiers = ActiveSheet.Range("Addresed_Controls")

For Each cell In identifiers
    Select Case Len(cell.Value)
        Case Is = 2
            cell.Value = "00" & cell.Value
        Case Is = 3
            cell.Value = "0" & cell.Value
        Case Is = 1
            cell.Value = "000" & cell.Value
        Case Else
            'no change
    End Select
Next cell


End Sub

Function LastRowInOneColumn(columnNumber As Integer) As Long
'Find the last used row in a Column: column A in this example
    Dim lastRow As Long
    With ActiveSheet
        LastRowInOneColumn = .Cells(.Rows.Count, columnNumber).End(xlUp).row
    End With
End Function


Sub transferBaselineComments()

Dim sourceWB As Workbook, destWB As Workbook
Dim sourceCell As Range, destCell As Range
Dim sourceRange As Range, destRange As Range

'workbooks must be open

Set sourceWB = Workbooks("ASA VSS SSP Working copy.xlsx")
Set destWB = Workbooks("OT SoA - E8 and Baseline controls3.xlsx")


With sourceWB
    .Activate
    Sheet1.Activate
    Set sourceRange = sourceWB.Sheets("ISM Compliance BL").Range("B2:B823") ' fortis spreadsheet
    'MsgBox sourceWB.Name
    'MsgBox "source range is:" & vbLf & sourceRange.Address
End With

With destWB
    .Activate
    Sheet1.Activate
    Set destRange = destWB.Sheets("VSS").Range("C2:C127") 'make sure correct sheet is chosen!!
    'MsgBox destWB.Name
    'MsgBox "dest range is:" & vbLf & destRange.Address
End With

'Exit Sub

'transfer source baseline comments to target spreadsheet
For Each destCell In destRange
    For Each sourceCell In sourceRange
        If sourceCell.Value = destCell.Value Then
            destCell.Offset(0, 8).Value = sourceCell.Offset(0, 15).Value 'implementation details
            'destCell.Offset(0, 15).Value = sourcecell.Offset(0, 15).Value 'implementation comments
            GoTo Skip
        End If
    Next sourceCell
Skip:
Next destCell
Set sourceRange = Nothing
Set destRange = Nothing
End Sub
Sub transferBaselineComments2()

Dim sourceWB As Workbook, destWB As Workbook
Dim sourceCell As Range, destCell As Range
Dim sourceRange As Range, destRange As Range

'workbooks must be open

Set sourceWB = Workbooks("NAIPS SoA - Final 1.0.xlsx") '2017 version of SoA
'Set destWB = Workbooks("NAIPS Draft SOA - 11_2024_NAIPS_SOA_v0.1.xlsx")
Set destWB = Workbooks("NAIPS Draft SOA - 11_2024_NAIPS_SOA_v0.1.xlsx")


With sourceWB
    .Activate
    Sheet1.Activate
    Set sourceRange = sourceWB.Sheets("ISM 2017").Range("e2:e946") ' column e
    'MsgBox sourceWB.Name
    'MsgBox "source range is:" & vbLf & sourceRange.Address
End With

With destWB
    .Activate
    Sheet1.Activate
    Set destRange = destWB.Sheets("NAIPS2024SOA").Range("D2:D954") 'make sure correct sheet is chosen!!
    'MsgBox destWB.Name
    'MsgBox "dest range is:" & vbLf & destRange.Address
End With

'Exit Sub

'transfer source baseline comments to target spreadsheet
For Each destCell In destRange
    For Each sourceCell In sourceRange
        If sourceCell.Value = destCell.Value Then
            If IsEmpty(destCell.Offset(0, 15).Value) Then 'only populate blank cells - don't want to overwrite newer baseline
                destCell.Offset(0, 14).Value = sourceCell.Offset(0, 11).Value 'implementation status
                destCell.Offset(0, 15).Value = sourceCell.Offset(0, 12).Value 'implementation comments
                destCell.Offset(0, 16).Value = "NAIPS 2017 SOA" 'sourced from old SOA - to be checked (reference column)
                
                GoTo Skip
            End If
        End If
    Next sourceCell
Skip:
Next destCell
Set sourceRange = Nothing
Set destRange = Nothing
End Sub
Sub transferBaselineComments3()

'using named ranges on worksheets instead of open separate workbooks

Dim sourceCell As Range, destCell As Range
Dim sourceRange As Range, destRange As Range

Set destRange = Range("ISM_DEC24")
Set sourceRange = Range("IATS")

'transfer source baseline comments to target spreadsheet
For Each destCell In destRange
    For Each sourceCell In sourceRange
        If sourceCell.Value = destCell.Value Then
            'If IsEmpty(destcell.Offset(0, 28).Value) Then 'only populate blank cells - don't want to overwrite newer baseline
                'destcell.Offset(0, 14).Value = sourcecell.Offset(0, 11).Value 'implementation status
                If sourceCell.Offset(0, 14).Value = "Not Applicable" Then
                    destCell.Offset(0, 28).Value = "N/A"
                Else
                    destCell.Offset(0, 28).Value = sourceCell.Offset(0, 14).Value 'implementation comments
                'destcell.Offset(0, 16).Value = "NAIPS 2017 SOA" 'sourced from old SOA - to be checked (reference column)
                
                GoTo Skip
                End If
            'End If
        End If
    Next sourceCell
Skip:
Next destCell
Set sourceRange = Nothing
Set destRange = Nothing
End Sub

Sub baselineImplemented()
Dim cell As Range, implementationStatus As Range

With ActiveSheet
    Set implementationStatus = .Range("R2:R954")
    For Each cell In implementationStatus
        If IsEmpty(cell.Offset(0, 1).Value) Then cell.Value = ""
    Next cell
End With
End Sub
Sub swapCategories()
Dim list As Range, cell As Range

Set list = Range("Q2:Q954")
For Each cell In Range
    If cell.Value = "ASA" Then
        cell.Value = "Yes"
    Else
        cell.Value = "No"
    End If
Next cell

End Sub
Sub PasteConditionalFormats(sourceCell As Range, destCell As Range)
'
' PasteConditionalFormats Macro
'
Application.CutCopyMode = False ' clear clipboard
sourceCell.Copy
destCell.PasteSpecial Paste:=xlPasteAllMergingConditionalFormats
'destcell.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
'SkipBlanks:=False, Transpose:=False
End Sub

Sub testpasteConditionals()
Dim refcell As Range, cell As Range

Set refcell = ActiveSheet.Range("R2")

For Each cell In Selection
    If cell.Offset(0, -2).Value = "Not in Scope" Then
        Call PasteConditionalFormats(refcell, cell)
        cell.Value = "Not Applicable"
    End If
Next cell
End Sub
Sub prependISMLabel()
Dim cell As Range, ISMCOntrols As Range
With ActiveSheet

Set ISMCOntrols = .Range("RVC_Controls")

For Each cell In ISMCOntrols
    cell.Value = "ISM-" & cell.Value
Next

End With
End Sub

