Attribute VB_Name = "Module4"
'====================================
'= Author: trungtrung (dinhquangtrung90@yahoo.com)
'= QTSS v3.0 - Open Source
'= Don't Forget Me When You Use This Program :) Thank
'====================================
'
'
'

' ((((((( Module mã hóa & giai mã: PSCODE


Option Explicit

Public Function DeCode(vText) As String
    On Error GoTo ErrHandler
    Dim CurSpc As Integer
    Dim varLen As Integer
    Dim varChr As String
    Dim varFin As String
    CurSpc = CurSpc + 1
    varLen = Len(vText)
    Do While CurSpc <= varLen
        DoEvents
        varChr = Mid(vText, CurSpc, 3)
        Select Case varChr
            'lower case
            Case "coe"
                varChr = "a"
            Case "wer"
                varChr = "b"
            Case "ibq"
                varChr = "c"
            Case "am7"
                varChr = "d"
            Case "pm1"
                varChr = "e"
            Case "mop"
                varChr = "f"
            Case "9v4"
                varChr = "g"
            Case "qu6"
                varChr = "h"
            Case "zxc"
                varChr = "i"
            Case "4mp"
                varChr = "j"
            Case "f88"
                varChr = "k"
            Case "qe2"
                varChr = "l"
            Case "vbn"
                varChr = "m"
            Case "qwt"
                varChr = "n"
            Case "pl5"
                varChr = "o"
            Case "13s"
                varChr = "p"
            Case "c%l"
                varChr = "q"
            Case "w$w"
                varChr = "r"
            Case "6a@"
                varChr = "s"
            Case "!2&"
                varChr = "t"
            Case "(=c"
                varChr = "u"
            Case "wvf"
                varChr = "v"
            Case "dp0"
                varChr = "w"
            Case "w$-"
                varChr = "x"
            Case "vn&"
                varChr = "y"
            Case "c*4"
                varChr = "z"
            'numbers
            Case "aq@"
                varChr = "1"
            Case "902"
                varChr = "2"
            Case "2.&"
                varChr = "3"
            Case "/w!"
                varChr = "4"
            Case "|pq"
                varChr = "5"
            Case "ml|"
                varChr = "6"
            Case "t'?"
                varChr = "7"
            Case ">^s"
                varChr = "8"
            Case "<s^"
                varChr = "9"
            Case ";&c"
                varChr = "0"
            'caps
            Case "$)c"
                varChr = "A"
            Case "-gt"
                varChr = "B"
            Case "|p*"
                varChr = "C"
            Case "1" & Chr(34) & "r"
                varChr = "D"
            Case "c>:"
                varChr = "E"
            Case "@+x"
                varChr = "F"
            Case "v^a"
                varChr = "G"
            Case "]eE"
                varChr = "H"
            Case "aP0"
                varChr = "I"
            Case "{=1"
                varChr = "J"
            Case "cWv"
                varChr = "K"
            Case "cDc"
                varChr = "L"
            Case "*,!"
                varChr = "M"
            Case "fW" & Chr(34)
                varChr = "N"
            Case ".?T"
                varChr = "O"
            Case "%<8"
                varChr = "P"
            Case "@:a"
                varChr = "Q"
            Case "&c$"
                varChr = "R"
            Case "WnY"
                varChr = "S"
            Case "{Sh"
                varChr = "T"
            Case "_%M"
                varChr = "U"
            Case "}'$"
                varChr = "V"
            Case "QlU"
                varChr = "W"
            Case "Im^"
                varChr = "X"
            Case "l|P"
                varChr = "Y"
            Case ".>#"
                varChr = "Z"
            'Special characters
            Case "\" & Chr(34) & "]"
                varChr = "!"
            Case "cY,"
                varChr = "@"
            Case "x%B"
                varChr = "#"
            Case "a*v"
                varChr = "$"
            Case "'&T"
                varChr = "%"
            Case ";%R"
                varChr = "^"
            Case "eG_"
                varChr = "&"
            Case "Z/e"
                varChr = "*"
            Case "rG\"
                varChr = "("
            Case "]*F"
                varChr = ")"
            Case "@B*"
                varChr = "_"
            Case "+Hc"
                varChr = "-"
            Case "&|D"
                varChr = "="
            Case "(:#"
                varChr = "+"
            Case "SlW"
                varChr = "["
            Case "'QB"
                varChr = "]"
            Case "{D>"
                varChr = "{"
            Case "+c%"
                varChr = "}"
            Case "(s:"
                varChr = ":"
            Case "^a("
                varChr = ";"
            Case "16."
                varChr = "'"
            Case "s.*"
                varChr = Chr(34)
            Case "&?W"
                varChr = ","
            Case "GPQ"
                varChr = "."
            Case "SK*"
                varChr = "<"
            Case "RL^"
                varChr = ">"
            Case "40C"
                varChr = "/"
            Case "?#9"
                varChr = "?"
            Case "_?/"
                varChr = "\"
            Case "(_@"
                varChr = "|"
            Case "=#B"
                varChr = " "
        End Select
        varFin = varFin & varChr
        CurSpc = CurSpc + 3
        DoEvents
    Loop
    DeCode = varFin
    Exit Function
ErrHandler:
    Dim ErrNum, ErrDesc, ErrSource
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    MsgBox "Error# = " & ErrNum & vbCrLf & "Description = " & ErrDesc & vbCrLf & "Source = " & ErrSource, vbCritical + vbOKOnly, "Program Error!"
    Err.Clear
    Exit Function
End Function


Public Function GiaiMa(xString)
GiaiMa = DeCode(DeCode(DeCode(DeCode(DeCode(DeCode(xString))))))
End Function

