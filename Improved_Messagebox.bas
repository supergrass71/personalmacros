Attribute VB_Name = "Improved_Messagebox"
Option Explicit
 
Function MsgBox2(Prompt As String, _
    Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
    Optional Title As String = "Microsoft Excel", _
    Optional HelpFile As String, _
    Optional Context As Long) As VbMsgBoxResult
     '
     '****************************************************************************************
     '       Title       MsgBox2
     '       Target Application: any
     '       Function:   substitute for standard MsgBox; displays more text than the ~1024 character
     '                   limit of MsgBox.  Displays blocks of approx 900 characters (properly split
     '                   at blanks or line feeds or "returns" and adds some "special text" to suggest
     '                   that more text is coming for each block except the last.  Special text is
     '                   easily changed.
     '
     '                   An EndOfBlack separator is also supported.  If found, MsgBox2 will only
     '                   display the characters through the EndOfBlock separator.  This provides
     '                   complete control over how text is displayed.  The current separator is
     '                   "||".
     '       Limitations:  the optional values for MsgBox display, i.e., Buttons, Title, HelpFile,
     '                     and Context  are the same for each block of text displayed.
     '       Passed Values:  same arguement list and type as standard MsgBox
     '
     '****************************************************************************************
     '
     '
    Dim CurLocn         As Long
    Dim EndOfBlock      As String
    Dim EOBIndex        As Integer
    Dim EOBLen          As Integer
    Dim Index           As Integer
    Dim MaxLen          As Integer
    Dim OldIndex        As Integer
    Dim strMoreToCome   As String
    Dim strTemp         As String
    Dim ThisChar        As String
    Dim TotLen          As Integer
     
     '
     '           set procedure variable that control how/what procedure does:
     '
     '       EndOfBlock is the string variable containing the character or characters
     '           that denote the end of a block of text.  These characters are not displayed.
     '           Do not use a character or characters that might be used in normal text.
     '       MaxLen is the maximum number of characters to be displayed at one time.  The
     '           limit for MsgBox is approx 1024, but that depends on the particular chars
     '           in the prompt string.  900 is a safe number as long as the len(strMoreToCome)
     '           is reasonable.
     '       strMoreToCome is text displayed at the bottom of each block indicating that more
     '           text/data is coming.
     '
    EndOfBlock = "||"
    MaxLen = 900
    strMoreToCome = " ... press any button except CANCEL to see next block of text ... "
     
    EOBLen = Len(EndOfBlock)
    CurLocn = 0
    OldIndex = 1
    TotLen = 0
     
NextBlock:
     '
     '           test for special break and, if found, that it is not the last chars in Prompt
     '
    EOBIndex = InStr(1, Mid(Prompt, OldIndex, MaxLen), EndOfBlock)
    If EOBIndex > 0 And CurLocn < Len(Prompt) - 1 Then
        CurLocn = EOBIndex + OldIndex - 1
        strTemp = Mid(Prompt, OldIndex, CurLocn - OldIndex)
        TotLen = TotLen + Len(strTemp) + EOBLen
        OldIndex = CurLocn + EOBLen
        GoTo MidDisplay
    End If
     '
     '           no special break, handle as normal block
     '
    Index = OldIndex + MaxLen
     '
     '           test for last block
     '
    If Index > Len(Prompt) Then
        strTemp = Mid(Prompt, OldIndex, Len(Prompt) - OldIndex + 1)
LastDisplay:
        MsgBox2 = MsgBox(strTemp, Buttons, Title, HelpFile, Context)
        Exit Function
    End If
     '
     '           not last display; process block
     '
    CurLocn = Index
NextIndex:
    ThisChar = Mid(Prompt, CurLocn, 1)
    If ThisChar = " " Or _
    ThisChar = Chr(10) Or _
    ThisChar = Chr(13) Then
         '
         '           block break found
         '
        strTemp = Mid(Prompt, OldIndex, CurLocn - OldIndex + 1)
        TotLen = TotLen + Len(strTemp)
        OldIndex = CurLocn + 1
MidDisplay:
         '
         '           display current block of text appending string indicating that
         '           more text is to come.  Then test if user hit Cancel button or
         '           equivalent; if so, exit MsgBox2 without further processing
         '
        MsgBox2 = MsgBox(strTemp & vbCrLf & strMoreToCome, _
        Buttons, Title, HelpFile, Context)
        If MsgBox2 = vbCancel Then Exit Function
        GoTo NextBlock
    End If
    CurLocn = CurLocn - 1
    If CurLocn > OldIndex Then GoTo NextIndex
     '
     '           no blanks, CR's, LF's or special breaks found in previous block
     '           display these characters and move on
     '
    strTemp = Mid(Prompt, OldIndex, MaxLen)
    CurLocn = OldIndex + MaxLen
    TotLen = TotLen + Len(strTemp)
    OldIndex = CurLocn + 1
    GoTo MidDisplay
     
End Function
 
Sub MsgBox2_Test(TestNum)
     '
     '****************************************************************************************
     '       Title       MsgBox2_Test
     '       Target Application: any
     '       Function;   demos use of MsgBox2
     '       Limitations:    none
     '       Passed Values:  none
     '****************************************************************************************
     '
     '
    Dim I           As Long
    Dim Answer      As VbMsgBoxResult
    Dim strPrompt   As String
     
    Select Case TestNum
    Case Is = 1
        strPrompt = "Initial stuff ..." & vbCrLf & vbCrLf
        For I = 48 To 122
            strPrompt = strPrompt & String(25, Chr(I)) & vbCrLf
        Next I
        strPrompt = strPrompt & vbCrLf & "... final stuff"
        Answer = MsgBox2(strPrompt, vbYesNoCancel, "1st Demo of MsgBox2")
    Case Is = 2
        strPrompt = "Initial stuff ..." & vbCrLf & vbCrLf
        For I = 48 To 122
            strPrompt = strPrompt & String(25, Chr(I))
        Next I
        strPrompt = strPrompt & vbCrLf & "... final stuff"
        Answer = MsgBox2(strPrompt, vbYesNoCancel, "2nd Demo of MsgBox2")
    Case Is = 3, 4
        strPrompt = "MsgBox is one of the most useful VB/VBA functions and it would be unlikely " & _
        "to find a VB/VBA application that did not use MsgBox at least once.  Unfortunately " & _
        "MsgBox has several not-easily-solved limitations, e.g., text size, text font, " & _
        "colors, and amount of text.  The former are irritating, but probably not fatal.  " & _
        "The latter, i.e., the amount of text that can be easily displayed via the Prompt " & _
        "string, is non-trivial.  MsgBox limits the number of characters to ~ 1024 (the " & _
        "exact number depends on the actual characters displayed).  If the length of Prompt " & _
        "is greater, the remaining characters are not displayed.  This can be particularly " & _
        "annoying (and possibly disastrous) if the last few words clarify an important " & _
        "result or what options are available or what is expected of the user." & vbCrLf & vbCrLf & "||"
        strPrompt = strPrompt & _
        "An alternative to MsgBox is a custom UserForm.  This is a good solution if one " & _
        "wants to improve several of MsgBox's limitations, but may be overkill if just " & _
        "displaying more text is desired." & vbCrLf & vbCrLf & "||" & _
        "MsgBox2 eliminates this limit by breaking the Prompt string into displayed blocks " & _
        "of approx 900 characters each.  For each block except the last, MsgBox2 displays " & _
        "the block and adds a line feed and special text suggesting that 'more data' is " & _
        "coming.  The special text is defined by the appl developer.  The current text is " & _
        vbCrLf & "       ... press any button except CANCEL to see next block of text ..." & vbCrLf & _
        "Text blocks are broken at " & _
        "logical separators: blanks; line feeds; or 'returns'.  Thus a Prompt string of, " & _
        "say, 2000 characters would be displayed in 3 blocks, the first two of approximately " & _
        "900 characters (ending with CrLf and '? more ?') and a final block with " & _
        "approximately 200 characters.  Each display is tested for 'Cancel' and, if " & _
        "encountered, MsgBox2 exits with a functional value equal to vbCancel or 2 (the " & _
        "numerical value for vbCancel)" & vbCrLf & vbCrLf & "||" & _
        "MsgBox2 also supports an 'end-of-block' option.  If the end-of-block character " & _
        "sequence is encountered (see code for current setting), MsgBox2 will automatically " & _
        "display the current buffer regardless of length." & vbCrLf & vbCrLf & "||" & _
        "Although simple is concept and execution, MsgBox2 is a very handy and" & vbCrLf & _
        "useful function.   MsgBox2 can be used in any VBA application." & vbCrLf & _
        "The demo is Excel based."
        If TestNum = 3 Then Answer = MsgBox2(strPrompt, vbYesNoCancel, "3rd Demo of MsgBox2")
        If TestNum = 4 Then MsgBox2 strPrompt, vbYesNoCancel, "4th Demo of MsgBox2"
         
    Case Else
        MsgBox "Invalid case fo MsgBox2_Test", vbCritical
    End Select
    If TestNum < 4 Then MsgBox "MsgBox2 return = " & MsgBoxResult(Answer)
     
End Sub
 
Function MsgBoxResult(Result As VbMsgBoxResult) As String
     '
     '****************************************************************************************
     '       Title       MsgBoxResult
     '       Target Application: any
     '       Function:   returns (as a string) the "vb constant" associated with a MsgBox result
     '       Limitations:    none
     '       Passed Values:
     '           Result  [input, type=vbMsgBoxResult] result or from call to MsgBox
     '****************************************************************************************
     '
     '
    Select Case Result
    Case Is = 1
        MsgBoxResult = "vbOK"
    Case Is = 2
        MsgBoxResult = "vbCancel"
    Case Is = 3
        MsgBoxResult = "vbAbort"
    Case Is = 4
        MsgBoxResult = "vbRetry"
    Case Is = 5
        MsgBoxResult = "vbAbort"
    Case Is = 6
        MsgBoxResult = "vbYes"
    Case Is = 7
        MsgBoxResult = "vbNo"
    Case Else
        MsgBoxResult = "UNKNOWN"
    End Select
     
End Function

