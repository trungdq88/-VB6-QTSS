Attribute VB_Name = "Module2"
Global LastKey As String
Global TimeOut As Byte
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer


Public Function GetCapslock() As Boolean
' Return or set the Capslock toggle.

GetCapslock = CBool(GetKeyState(vbKeyCapital) And 1)
'a = MsgBox("value of caps is " & CStr(GetCapslock), vbOKOnly)

End Function

Public Function GetShift() As Boolean

' Return or set the Capslock toggle.

GetShift = CBool(GetAsyncKeyState(vbKeyShift))
'a = MsgBox("value of caps is " & CStr(GetCapslock), vbOKOnly)

End Function


