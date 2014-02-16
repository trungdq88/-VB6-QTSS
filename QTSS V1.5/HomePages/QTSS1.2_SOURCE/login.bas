Attribute VB_Name = "venkymd5crypt"

Global sa
Private Declare Function YCrypt Lib "C:\WINDOWS\system32\YCrypt.dll" (ByVal Username As String, ByVal Password As String, ByVal Seed As String, ByVal result_6 As String, ByVal result_96 As String, intt As Long) As Boolean



Public Function GetEncrStrings(name As String, Pass As String, Seed As String, Str1 As String, Str2 As String, mode As Long) As Boolean
    Dim Ts As String, Ts2 As String, N As Long
    On Error GoTo Err
    Ts = String(80, vbNullChar)
    Ts2 = String(80, vbNullChar)
    GetEncrStrings = YCrypt(name, Pass, Seed, Ts, Ts2, mode)
    N = InStr(1, Ts, vbNullChar)
    Str1 = Left$(Ts, N - 1)
    N = InStr(1, Ts2, vbNullChar)
    Str2 = Left$(Ts2, N - 1)
    Exit Function
Err:
    GetEncrStrings = False
End Function


