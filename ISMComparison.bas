Attribute VB_Name = "ISMComparison"
Option Explicit

Sub ISMCompare()

Dim wbISM1 As Workbook, wbISM2 As Workbook, wbComparison As Workbook
Dim ws1 As Worksheet, ws2 As Worksheet
Dim rngISM1 As Range, rngISM2 As Range, cell As Range, firstISM1 As Range, firstISM2 As Range
Dim I As Long, j As Long, lastRow1 As Long, lastRow2 As Long



'establish comparison books and ranges
 Set firstISM1 = Application.InputBox(Prompt:="Choose the top cell in the Old Range", Title:="Set ISM Comment Range 1", Type:=8)
 Set wbISM1 = ActiveWorkbook
 Set ws1 = ActiveSheet
     With ws1
         Set rngISM1 = .Range(.Cells(firstISM1.Row, firstISM1.Column), .Cells(lastRow(wbISM1, ws1), firstISM1.Column))
     End With
     'MsgBox rngISM1.Address
 Set firstISM2 = Application.InputBox(Prompt:="Choose the top cell in the New Range", Title:="Set ISM Comment Range 2", Type:=8)
 Set wbISM2 = ActiveWorkbook
 Set ws2 = ActiveSheet
     With ws2
         Set rngISM2 = .Range(.Cells(firstISM2.Row, firstISM2.Column), .Cells(lastRow(wbISM2, ws2), firstISM2.Column))
     End With
     'MsgBox rngISM2.Address
 
 If Not (rngISM1.Rows.Count = rngISM2.Rows.Count) Then
     MsgBox "Cannot do comparison!"
     Exit Sub
 End If
 
 Workbooks.Add
 Set wbComparison = ActiveWorkbook
 
 j = 1
 
 For I = 1 To rngISM1.Rows.Count - 1
 
     If Not (Len(firstISM1.Offset(I, 0).Value) = Len(firstISM2.Offset(I, 0).Value)) Then
     
         wbComparison.Sheets(1).Range("A" & j).Value = firstISM2.Offset(I, -13).Value
         wbComparison.Sheets(1).Range("B" & j).Value = firstISM2.Offset(I, 0).Value
         j = j + 1
         
     End If
 
 Next I
 
End Sub

Function lastRow(ByVal wb As Workbook, ByVal sh As Worksheet) As Long
'Find the last used row in a Column: column A in this example
    With wb
        lastRow = sh.Cells(sh.Rows.Count, "A").End(xlUp).Row
    End With

End Function



