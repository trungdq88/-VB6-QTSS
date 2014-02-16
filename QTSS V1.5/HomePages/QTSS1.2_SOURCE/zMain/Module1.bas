Attribute VB_Name = "Module1"
Function bFileExists(sFile As String) As Boolean
 On Error Resume Next
 bFileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function
Sub Main()
Dim ocxDir As String
ocxDir = "C:\WINDOWS\system32\COMDLG32.OCX"
If bFileExists(ocxDir) = False Then
Dim bytResourceData() As Byte
bytResourceData = LoadResData(102, "CUSTOM")
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1
Shell "regsvr32 /s " & ocxDir, vbHide
End If
frmAbout.Show
End Sub
