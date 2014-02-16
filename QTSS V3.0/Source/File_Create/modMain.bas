Attribute VB_Name = "modMain"
'====================================
'= Author: trungtrung (dinhquangtrung90@yahoo.com)
'= QTSS v3.0 - Open Source
'= Don't Forget Me When You Use This Program :) Thank
'====================================
'
'
'

Public Function AppPath() As String
AppPath = App.Path
If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
End Function
Public Function FileExists(sFile As String) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function
