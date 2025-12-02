VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
Unload UserForm1
End Sub


Private Sub CommandButton2_Click()
    With ActiveSheet
        .Range("A1").AutoFilter
    End With
End Sub

Private Sub TextBox1_Change()
Dim controlNumber

'If Left(ActiveWorkbook.Name, 3) = "APP" Then Exit Sub

If Me.TextBox1.TextLength = 4 Then
    controlNumber = Me.TextBox1.Value
    With ActiveSheet
        .Range("A1").AutoFilter
        .Range("$A$1:$S$981").AutoFilter Field:=5, Criteria1:=controlNumber
    End With
End If
If Me.TextBox1.TextLength = 0 Then
    controlNumber = Me.TextBox1.Value
    With ActiveSheet
        .Range("A1").AutoFilter
    End With
End If
End Sub
