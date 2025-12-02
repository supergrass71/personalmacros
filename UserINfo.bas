Attribute VB_Name = "UserINfo"
  Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                      (ByVal IpBuffer As String, nSize As Long) As Long
  Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
                      (ByVal lpBuffer As String, nSize As Long) As Long



Sub AAA()
Dim UName As String * 255
Dim L As Long: L = 255
Dim Res As Long
Res = GetUserName(UName, L)
UName = Left$(UName, L - 1)
MsgBox UName
MsgBox GetUserFullName
End Sub

Function GetUserFullName() As String
'this one will return the user alias seen at the Windows button
    Dim WSHnet As Object, objuser As Object
    Dim userName As String, UserDomain As String
    Set WSHnet = CreateObject("WScript.Network")
    userName = WSHnet.userName
    UserDomain = WSHnet.UserDomain
    Set objuser = GetObject("WinNT://" & UserDomain & "/" & userName & ",user")
    GetUserFullName = objuser.FullName
    Set objuser = Nothing
End Function
