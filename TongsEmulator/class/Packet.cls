VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Packet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private PacketData() As Byte
Public PacketIndex As Long

Private Sub Class_Initialize()
    ReDim PacketData(0)
    PacketIndex = 0
End Sub

Public Sub SetPacketData(data() As Byte)
    PacketData = data
End Sub

Public Function GetPacketData() As Byte()
    GetPacketData = PacketData
End Function

' 보내기 전용 패킷일때 맨 처음 초기화하기 위함
' XX XX YY YY DD DD DD DD ... : XX = 옵코드 / YY = 길이 / UU = 아직모르는데이터
Public Sub InitSendPacket(OpCode As Long)
    ReDim PacketData(7)
    PacketIndex = 7
    
    CopyMemory PacketData(0), OpCode, 2
End Sub

Private Sub FixPacketLen()
    Dim PacketLen As Integer
    PacketLen = PacketIndex + 1
    CopyMemory PacketData(2), PacketLen, 2
End Sub

Public Sub WriteByte(value As Long)
    ReDim Preserve PacketData(PacketIndex + 1)
    
    CopyMemory PacketData(PacketIndex + 1), value, 1
    
    PacketIndex = PacketIndex + 1
    
    FixPacketLen
End Sub

Public Sub WriteShort(value As Long)
    ReDim Preserve PacketData(PacketIndex + 2)
    
    CopyMemory PacketData(PacketIndex + 1), value, 2
    
    PacketIndex = PacketIndex + 2
    
    FixPacketLen
End Sub

Public Sub WriteInt(value As Long)
    ReDim Preserve PacketData(PacketIndex + 4)
    
    CopyMemory PacketData(PacketIndex + 1), value, 4
    
    PacketIndex = PacketIndex + 4
    
    FixPacketLen
End Sub

Public Sub WriteString(value As String)
    Dim str() As Byte, strlen As Long
    str = StrConv(value, vbFromUnicode)
    strlen = UBound(str) + 1

    ReDim Preserve PacketData(PacketIndex + strlen)
    
    CopyMemory PacketData(PacketIndex + 1), value, strlen
    
    PacketIndex = PacketIndex + strlen
    
    FixPacketLen
End Sub

Public Sub WriteUnicodeString(value As String)
    Dim str() As Byte, strlen As Long
    str = StrConv(StrConv(value, vbUnicode), vbFromUnicode)
    strlen = UBound(str) + 1

    ReDim Preserve PacketData(PacketIndex + strlen)
    
    CopyMemory PacketData(PacketIndex + 1), value, strlen
    
    PacketIndex = PacketIndex + strlen
    
    FixPacketLen
End Sub

Public Sub WriteBytes(value() As Byte)
    Dim bLen As Long
    bLen = UBound(value) + 1
    
    ReDim Preserve PacketData(PacketIndex + bLen)
    
    CopyMemory PacketData(PacketIndex + 1), value(0), bLen
    
    PacketIndex = PacketIndex + bLen
    
    FixPacketLen
End Sub

Public Function ReadByte() As Long
    If UBound(PacketData) + 1 < PacketIndex + 1 Then
        Log "Packet::ReadByte() 범위 초과", 0
        Exit Function
    End If

    Dim result As Long
    CopyMemory result, PacketData(PacketIndex), 1
    PacketIndex = PacketIndex + 1
    
    ReadByte = result
End Function

Public Function ReadShort() As Long
    If UBound(PacketData) + 1 < PacketIndex + 2 Then
        Log "Packet::ReadShort() 범위 초과", 0
        Exit Function
    End If

    Dim result As Long
    CopyMemory result, PacketData(PacketIndex), 2
    PacketIndex = PacketIndex + 2
    
    ReadShort = result
End Function

Public Function ReadInt() As Long
    If UBound(PacketData) + 1 < PacketIndex + 4 Then
        Log "Packet::ReadInt() 범위 초과", 0
        Exit Function
    End If

    Dim result As Long
    CopyMemory result, PacketData(PacketIndex), 4
    PacketIndex = PacketIndex + 4
    
    ReadInt = result
End Function

Public Function ReadString(slen As Long) As String
    If UBound(PacketData) + 1 < PacketIndex + slen Then
        Log "Packet::ReadString() 범위 초과", 0
        Exit Function
    End If

    Dim result As String
    Dim temp() As Byte
    ReDim temp(slen)
    CopyMemory temp(0), PacketData(PacketIndex), slen
    result = StrConv(temp, vbUnicode)
    PacketIndex = PacketIndex + slen
    
    ReadString = result
End Function

