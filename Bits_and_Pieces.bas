Attribute VB_Name = "Bits_and_Pieces"
Private Type SystemTime
Year As Integer
Month As Integer
DayOfWeek As Integer
Day As Integer
Hour As Integer
Minute As Integer
Second As Integer
Milliseconds As Integer
End Type
'check 32 or 64 bit excel
#If VBA7 Then
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As SystemTime)
#Else
Private Declare Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As SystemTime)
#End If

Function GetTodayMilliseconds() As Long
Dim CurrentTime As SystemTime
GetSystemTime CurrentTime
GetTodayMilliseconds = Hour(Now) * 3600000 + Minute(Now) * 60000 + _
Second(Now) * 1000 + CurrentTime.Milliseconds
End Function
Function GetMilliseconds() As Long
Dim CurrentTime As SystemTime
GetSystemTime CurrentTime
GetMilliseconds = CurrentTime.Milliseconds
End Function
Sub testgetMilliseconds()
MsgBox GetMilliseconds
End Sub
Sub testMilliseconds()
MsgBox GetTodayMilliseconds
If Int(Right(GetMilliseconds, 1)) Mod 2 = 0 Then
    MsgBox "2"
Else
    MsgBox "1"
End If
End Sub
Function oneOrTwo() As Integer
Select Case Int(Right(GetMilliseconds, 1)) Mod 2
    Case Is = 0
        oneOrTwo = 2
    Case Else
        oneOrTwo = 1
End Select
End Function
Sub ShadeYellow(cell As Range)
'
    With cell.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub Resize_Rows_Columns()
Attribute Resize_Rows_Columns.VB_ProcData.VB_Invoke_Func = "r\n14"
'
' Resize_Rows_Columns Macro
'
' Keyboard Shortcut: Ctrl+r
Dim RowsToFix As Range
'With Application
   ' .ScreenUpdating = False
'    On Error Resume Next
'    Set RowsToFix = .InputBox(Prompt:="Select Rows to resize", Title:="Resize Rows", Type:=8)
'    If Not RowsToFix Is Nothing Then
'        RowsToFix.EntireRow.AutoFit
       ' .ScreenUpdating = True
'        On Error GoTo 0
'    Else
'        On Error GoTo 0
'        Exit Sub
'    End If
'End With
Selection.EntireRow.Autofit
End Sub

Sub addDate()
Attribute addDate.VB_ProcData.VB_Invoke_Func = "d\n14"
Dim cell As Range
Dim dateStamp As String
Dim ans As Integer, y1 As Integer, y2 As Integer

dateStamp = Format(Now(), "YYYYmmdd")
dateStamp = dateStamp & " " & "MN:"

For Each cell In Selection

y1 = InStr(1, cell.Value, "2024", vbTextCompare) '2024 datestamp in cell
y2 = InStr(1, cell.Value, "2025", vbTextCompare) '2025 datestamp in cell

    
    If (y1 > 0 And y2 > 0) And Len(cell.Value) > 0 Then
        cell.Value = dateStamp & cell.Value
    Else
        If IsEmpty(cell.Value) Then
            cell.Value = dateStamp & " checked"
        Else
            cell.Value = cell.Value & vbLf & dateStamp & " checked"
        End If
    End If
    If Selection.Rows.Count = 1 Then 'single cell
        ans = MsgBox(Prompt:="Change Shading/Font?", Title:="Update Cell", Buttons:=vbYesNo)
        If ans = vbYes Then Call Change_Shading(cell)
    End If
    
Next cell
End Sub
Sub Change_Shading(cell As Range)

If cell.Interior.Pattern = xlNone Then
    With cell.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
Else
    With cell.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End If

'With cell.Font
'    .ColorIndex = xlAutomatic
    '.TintAndShade = 0
'End With
End Sub
Function randomiseBinary() As Integer
Dim millisecond As Integer


End Function
Sub LastRowInOneColumn()
'Find the last used row in a Column: column A in this example
    Dim lastRow As Long
    With ActiveSheet
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
    End With
    MsgBox lastRow
End Sub
Function getLastRow() As Long
'Find the last used row in a Column: column A in this example
    Dim lastRow As Long
    With ActiveSheet
        getLastRow = .Cells(.Rows.Count, "A").End(xlUp).row
    End With
    'MsgBox lastrow
End Function
Function LastColumnInOneRow(row As Integer) As Long
'Find the last used column in a Row: row 1 in this example
    With ActiveSheet
        LastColumnInOneRow = .Cells(row, .Columns.Count).End(xlToLeft).column
    End With
End Function
Function getLastRow2(column As Integer) As Long
'Find the last used row in a Column: column A in this example
    Dim lastRow As Long
    With ActiveSheet
        getLastRow2 = .Cells(.Rows.Count, column).End(xlUp).row
    End With
    'MsgBox lastrow
End Function
Sub ReplaceLineBreaksWithCommas()
    Dim cell As Range, newCell As Range
    Dim originalText As String
    Dim modifiedText As String

    ' Change this to the specific cell you want to modify
    Set cell = ActiveCell
    Set newCell = ActiveCell.Offset(0, 1)

    originalText = cell.Value
    ' Replace line breaks (vbLf or vbCrLf) with commas
    modifiedText = Replace(originalText, vbLf, ", ")
    modifiedText = Replace(modifiedText, vbCrLf, ", ")

    ' Update the cell with the modified text
    newCell.Value = modifiedText
End Sub

Sub ClearFillColor()
Attribute ClearFillColor.VB_ProcData.VB_Invoke_Func = "q\n14"
    Dim cell As Range

    For Each cell In Selection
        cell.Interior.ColorIndex = xlNone
    Next cell
End Sub

Sub InsertDigitalSignatureInfo()
    Dim shell As Object
    Dim signatureInfo As String
    Dim psCommand As String

    ' PowerShell command to get signature info
    psCommand = "powershell -command ""(Get-AuthenticodeSignature 'C:\Path\To\Your\File.xlsx').SignerCertificate.Subject"""

    ' Create shell object
    Set shell = CreateObject("WScript.Shell")

    ' Run PowerShell and get output
    signatureInfo = shell.Exec(psCommand).StdOut.ReadAll

    ' Insert into cell A1
   ActiveCell.Value = signatureInfo
End Sub

