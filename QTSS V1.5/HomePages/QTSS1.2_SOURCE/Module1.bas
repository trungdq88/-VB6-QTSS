Attribute VB_Name = "Pub"
Option Explicit
Public Declare Function IsCharUpper Lib "user32" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long


Global NumOfOpenIMs As Byte
Global OpenedImsCaptions(1 To 100) As String



Global Separator As String

Global SessionID As String



Public Const YAHOO_SERVICE_AUTHRESP = &H54
Public Const YAHOO_SERVICE_AUTH = &H57
Public Const YAHOO_SERVICE_MESSAGE = &H6
Public Const YAHOO_STATUS_AVAILABLE = &H0


Public Type Ymsg9Packet

    Service(1 To 2) As Byte
    Status(1 To 4) As Byte
    SessionID(1 To 4) As Byte
    Data As String
    IsInline As Boolean
End Type


Global SentPacket As Ymsg9Packet
Global WinsockIsInSendingProgress As Boolean
Global TotalPacketsBufferd As Byte
Global PreviousPacketFinished As Boolean
Global ReceivedPacket(1 To 30) As Ymsg9Packet
Global InlinePacketLen(1 To 255) As Byte
Global LastPacketReceived As Ymsg9Packet
Global ReceivedInlineFragmentedData(1 To 30, 1 To 255) As String
Global TotalInlinesBufferd As Byte
Global LastWSState As Byte
Global Last97AltValue As String * 1
Global FunctionMadeError As String
Global ErrorFileNum As Byte

Function DataSection(ymp() As Byte) As String
Dim strT As String
Dim i As Long
For i = LBound(ymp) + 20 To UBound(ymp)
    If ymp(i) <> 0 Then
        strT = strT + Chr(ymp(i))
    Else
        strT = strT + "<0x0>"
    End If
Next i
DataSection = strT
End Function




Function AnySection2HexString(ymp() As Byte) As String
Dim i As Byte
Dim strT As String
For i = LBound(ymp) To UBound(ymp)
    strT = strT + Hex(ymp(i))
Next i
AnySection2HexString = strT
End Function





Public Sub ConvertYmsgType2ByteArray(YmsgP As Ymsg9Packet, ByteArray() As Byte)
    Dim DataLen As Long
    Dim i As Long
    ReDim ByteArray(1 To 20 + Len(YmsgP.Data))
    ByteArray(1) = Asc("Y")
    ByteArray(2) = Asc("M")
    ByteArray(3) = Asc("S")
    ByteArray(4) = Asc("G")
    
    ByteArray(5) = 0
    ByteArray(6) = 12
    ByteArray(7) = 0
    ByteArray(8) = 0
    
    DataLen = Len(YmsgP.Data)
    ByteArray(9) = DataLen \ 256
    ByteArray(10) = DataLen Mod 256

    ByteArray(11) = YmsgP.Service(1)
    ByteArray(12) = YmsgP.Service(2)

    ByteArray(13) = YmsgP.Status(1)
    ByteArray(14) = YmsgP.Status(2)
    ByteArray(15) = YmsgP.Status(3)
    ByteArray(16) = YmsgP.Status(4)
    
    ByteArray(17) = YmsgP.SessionID(1)
    ByteArray(18) = YmsgP.SessionID(2)
    ByteArray(19) = YmsgP.SessionID(3)
    ByteArray(20) = YmsgP.SessionID(4)
    
    For i = 1 To DataLen
        ByteArray(20 + i) = Asc(Mid(YmsgP.Data, i, 1))
    Next i

End Sub
