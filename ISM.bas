Attribute VB_Name = "ISM"
Option Explicit
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Function ScrapeDisplayedContent(webpage As String) As String
'based on code from Co-Pilot
    Dim ie As Object
    Dim html As Object
    Dim url As String
    Dim content As String

    ' Set the URL of the page you want to scrape
    'url = "https://example.com"

    ' Create Internet Explorer object
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = False ' Set to True if you want to see the browser

    ' Navigate to the URL
    ie.Navigate webpage

    ' Wait for the page to fully load
    Do While ie.Busy Or ie.ReadyState <> 4
        DoEvents
    Loop
    ' Get the rendered HTML document
    Set html = ie.Document

    ' Extract the visible text content
    ScrapeDisplayedContent = html.body.innerText

    ' Output to Immediate Window (Ctrl+G in VBA editor)
    'Debug.Print content
    ' Clean up
    ie.Quit
    Set ie = Nothing
    Set html = Nothing
End Function

Sub lookupISMControl()
Dim ismControl As String, webpage As String
'perform input sanitisation based on ISM control number
Select Case Len(ActiveCell.Value)
    Case Is = 8
        ismControl = Right(ActiveCell.Value, 4) 'contains "ISM-" as prefix
    Case Is = 4
        ismControl = ActiveCell.Value
    Case Else
        MsgBox "ISM control not found!", vbInformation
        Exit Sub
End Select
webpage = "https://ismcontrol.xyz/" & "ismControl" & ".html "

Run addcommentText(ScrapeDisplayedContent(webpage))

End Sub
Function addcommentText(text As String)

If Not ActiveCell.CommentThreaded Is Nothing Then ActiveCell.ClearComments
ActiveCell.AddCommentThreaded (text)

End Function

Function IsSeleniumInstalled() As Boolean
    On Error GoTo ErrHandler
    ' Try to create a Selenium WebDriver object
    Dim bot As Object
    Set bot = CreateObject("Selenium.WebDriver")

    ' If successful, Selenium is installed
    IsSeleniumInstalled = True
    Exit Function

ErrHandler:
    ' If error occurs, Selenium is not installed
    IsSeleniumInstalled = False
End Function

Sub testforSelenium()
If IsSeleniumInstalled Then
    MsgBox "selenium found"
Else
    MsgBox "selenium not found"
End If
End Sub
Sub ISMControlXYZComment()
    Dim shell As Object
    Dim psCommand As String
    Dim scriptPath As String
    
    
    scriptPath = "C:\Users\neish_mj\Documents\ISM_Control.ps1" 'have to change this to suit!

    ' PowerShell command to get signature info
    psCommand = "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & scriptPath & """"

    ' Create shell object
    Set shell = CreateObject("WScript.Shell")

    ' Run PowerShell and get output
    shell.Exec (psCommand)
'    While ActiveCell.CommentThreaded Is Nothing
'       Application.Cursor = xlWait
 '       Sleep 1000
  '  Wend
   ' Application.Cursor = xlDefault
End Sub

Sub DeleteRow()
Dim cell As Range
For Each cell In Selection
    If Not IsEmpty(cell.Value) Then cell.EntireRow.Delete
Next cell
End Sub

Sub ChangeToYesNo()
Dim cell As Range
For Each cell In Selection
    If cell.Value = "Not Applicable" Then
        cell.Value = "No"
    Else
        cell.Value = "Yes"
    End If
Next cell
End Sub
Sub Applicability()
Dim cell As Range
For Each cell In Selection
    If Left(cell.Value, 1) = "(" Or Left(cell.Value, 1) = "O" Then
        cell.Value = "Yes"
    Else
        cell.Value = "No"
    End If
Next cell
End Sub

Sub ChangeResponsible()
Dim cell As Range
Dim contents As String
For Each cell In Selection
    contents = cell.Value
    Select Case contents
        Case Is = "ASA"
            cell.Value = "Airservices"
        Case Is = "Both"
            cell.Value = "Shared"
        Case Is = "N/A"
        Case Is = ""
        Case Else
            cell.Value = "VSS"
    End Select
Next cell
End Sub

