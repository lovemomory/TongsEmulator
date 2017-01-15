VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents socket As Winsock
Attribute socket.VB_VarHelpID = -1

Public Key As String

Public ID As String
Public PW As String
Public IsLogin As Boolean
Public Nick As String

Public Character As Long

Public UserState As Long
Public UserLevel As Long
Public UserBlockDate As String

Private Sub Class_Initialize()
    Set socket = New Winsock
End Sub

Private Sub socket_Close()
    Log "<< 접속 해제 / IP: " & socket.RemoteHostIP & " >>"
End Sub

Private Sub socket_DataArrival(ByVal bytesTotal As Long)
    Dim p As Packet
    Set p = New Packet

    Dim data() As Byte
    socket.GetData data, vbByte, bytesTotal
    p.SetPacketData data
    
    Dim OpCode As Long, DataLen As Long, Unknown As Long
    OpCode = p.ReadShort()
    DataLen = p.ReadShort()
    Unknown = p.ReadInt()
    
    Select Case OpCode
    Case 20001: ' 로그인 요청
        OnLoginReq p
    Case 20021: ' 캐릭터 선택 완료
        OnSelectCharacter p
    Case 30002: ' 채팅
        OnChat p
    Case 50007: ' 방 만들기
        OnCreateRoom p
    Case Else:
        Log "<< 미정의 패킷 수신 / OpCode: " & OpCode & " >>"
        Open AP() & "PacketDump\" & CStr(OpCode) & ".txt" For Append As #1
        Print #1, HexToStr(data)
        Print #1, StrConv(data, vbUnicode)
        Print #1, ""
        Close #1
    End Select
End Sub

Private Sub socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Log "<< 소켓 오류 / IP: " & socket.RemoteHostIP & " >>"
End Sub

Private Sub OnLoginReq(p As Packet)
    Dim TempID As String, TempPW As String
    TempID = p.ReadString(24)
    TempPW = p.ReadString(24)
    
    ' 널문자 제거
    TempID = RemoveNullStr(TempID)
    TempPW = RemoveNullStr(TempPW)
    
    Dim UD1 As Long
    UD1 = p.ReadInt()
    
    Dim TempSex As Long
    TempSex = p.ReadInt()
    
    Log "<< 로그인 요청 >>" & vbCrLf & _
        "----------" & vbCrLf & _
        "요청 클라이언트: " & Key & vbCrLf & _
        "아이디: " & TempID & vbCrLf & _
        "비밀번호: " & TempPW & vbCrLf & _
        "UnknownData1: " & UD1 & vbCrLf & _
        "성별: " & TempSex & vbCrLf & _
        "----------"

    ' 로그인 확인하기
    Dim UserPW As String
    UserPW = ReadINI("info", "PW", AP & "Users\" & TempID & ".txt", "")

    UserState = ReadINI("info", "State", AP & "Users\" & TempID & ".txt", "")
    UserLevel = Val(ReadINI("info", "Level", AP & "Users\" & TempID & ".txt", "0"))
    UserBlockDate = ReadINI("info", "BlockDate", AP & "Users\" & TempID & ".txt", "")
    Nick = ReadINI("info", "Nick", AP & "Users\" & TempID & ".txt", "")
    
    ' 로그인 결과
    ' PTCL 20002 : 1 = 성공
    ' 2 = 통스통스 서버 로그인 실패
    ' 3 = 이미 로그인된 아이디
    ' 4 = 통스통스 아이디 비밀번호 일치하지 않음
    ' 5 = 오픈베타 서비스 기간 아님
    ' 6 = 등록된 베타테스터가 아님
    ' 7 = 클라이언트 버전이 일치하지 않음
    ' 8 = 욕설 및 음담패설로 1차경고
    ' 9 = 욕설 및 음담패설로 XX까지 계정정지 (메시지를 보냄)
    ' 기타 = 로그인 실패
    ' PTCL 20023 : 1 = 성공
    ' 2 = 별명 입력에 실패
    ' 3 이후 = 별명이 중복됨
    
    Dim LoginResult As Long
    LoginResult = 1
    If Val(UserState) = 2 Then LoginResult = 8
    If Val(UserState) = 3 Then LoginResult = 9
    If StrComp(TempPW, UserPW, vbBinaryCompare) <> 0 Then LoginResult = 4
    If StrComp("", UserPW, vbBinaryCompare) = 0 Then LoginResult = 4
    
    ' 로그인에 성공했다면 로그인했다는 기록을 남기자
    If LoginResult = 1 Then
        ID = TempID
        PW = TempPW
        IsLogin = True
    End If
    
    ' 패킷 만들기
    Dim op As Packet
    Set op = New Packet
    
    op.InitSendPacket 20002
    op.WriteInt LoginResult
    op.WriteInt 0
    op.WriteInt 0
    op.WriteInt UserLevel
    op.WriteInt 0
    op.WriteInt 0
    op.WriteString UserBlockDate
    
    socket.SendData op.GetPacketData
    DoEvents
End Sub

Private Sub OnSelectCharacter(p As Packet)
    Dim TempID As String
    TempID = p.ReadString(24)
    TempID = RemoveNullStr(TempID)
    
    Dim TempCharacter As Long
    TempCharacter = p.ReadInt()
    
    Log "<< 캐릭터 선택 >>" & vbCrLf & _
        "----------" & vbCrLf & _
        "요청 클라이언트: " & Key & vbCrLf & _
        "아이디: " & TempID & vbCrLf & _
        "캐릭터: " & TempCharacter & vbCrLf & _
        "----------"

    ' 이 패킷을 보낸사람이 정상인가?
    If TempID <> ID Or IsLogin = False Then
        Log "로그인 되지 않은 클라이언트가 요청함 OnSelectCharacter()"
        Exit Sub
    End If
    
    ' 캐릭터 변경
    Character = TempCharacter
    
    ' 닉네임은 최대 12바이트
    Dim TempNick() As Byte
    TempNick = StrConv(Nick, vbFromUnicode)
    ReDim Preserve TempNick(11)
    
    ' 패킷 만들기
    Dim op As Packet
    Set op = New Packet
    
    op.InitSendPacket 50002
    op.WriteShort 1
    op.WriteShort Character
    op.WriteShort 1
    op.WriteShort UserLevel
    op.WriteShort 0
    op.WriteShort 0
    op.WriteBytes TempNick
    op.WriteByte 71
    op.WriteByte 71
    op.WriteInt 0
    op.WriteInt 0
    op.WriteInt 0
    op.WriteInt 0
    op.WriteInt 0
    op.WriteInt 0
    op.WriteInt 0
    op.WriteInt 0
    op.WriteInt 0
    
    socket.SendData op.GetPacketData
    DoEvents
End Sub

Private Sub OnChat(p As Packet)
    ' 32 75 B0 00 CA 7A 00 00 CC CC CC CC CC CC CC CC 01 00 00 00 01 00 00 00 61 62 63 64 65 66 67 00 CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC 47 4D 5F BC AE 00 CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC

    Dim UD1 As Long, UD2 As Long
    UD1 = p.ReadInt()
    UD2 = p.ReadInt()
    
    Dim UD3 As Long, UD4 As Long
    UD3 = p.ReadInt() ' 1
    UD4 = p.ReadInt() ' 1
    
    Dim TempChat As String, rChat As String
    TempChat = p.ReadString(128)
    TempChat = RemoveNullStr(TempChat)
    
    Dim TempNick As String, rNick As String
    TempNick = p.ReadString(24)
    TempNick = RemoveNullStr(TempNick)
    
    Log "<< 채팅 >>" & vbCrLf & _
        "----------" & vbCrLf & _
        "요청 클라이언트: " & Key & vbCrLf & _
        "UD3, UD4: " & UD3 & ", " & UD4 & vbCrLf & _
        "닉네임: " & TempNick & vbCrLf & _
        "내용: " & TempChat & vbCrLf & _
        "----------"

    ' 이 패킷을 보낸사람이 정상인가?
    If TempNick <> Nick Or IsLogin = False Then
        Log "로그인 되지 않은 클라이언트가 요청함 OnChat()"
        Exit Sub
    End If
    
    ' 패킷 만들기
    Dim op As Packet
    Set op = New Packet
    
    op.InitSendPacket 30002
    op.WriteInt 0
    op.WriteInt 0
    op.WriteInt UD3
    op.WriteInt UD4
    op.WriteString TempChat
    op.WriteString TempNick

    socket.SendData p.GetPacketData
    DoEvents
End Sub

Private Sub OnCreateRoom(p As Packet)
    ' 32 75 B0 00 CA 7A 00 00 CC CC CC CC CC CC CC CC 01 00 00 00 01 00 00 00 61 62 63 64 65 66 67 00 CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC 47 4D 5F BC AE 00 CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC CC

    Dim TempRoomName As String
    TempRoomName = p.ReadString(24)
    TempID = RemoveNullStr(TempRoomName)
    
    Dim TempRoomPass As String
    TempRoomPass = p.ReadString(12)
    TempRoomPass = RemoveNullStr(TempRoomPass)
    
    Dim MaxUser As Long
    MaxUser = p.ReadInt()
    
    Dim UD1 As Long
    UD1 = p.ReadInt()
    
    Dim NoItem As Long
    NoItem = p.ReadInt()
    
    Log "<< 방 만들기 >>" & vbCrLf & _
        "----------" & vbCrLf & _
        "요청 클라이언트: " & Key & vbCrLf & _
        "방 이름: " & TempRoomName & vbCrLf & _
        "방 비밀번호: " & TempRoomPass & vbCrLf & _
        "인원수: " & MaxUser & vbCrLf & _
        "노템전: " & NoItem & vbCrLf & _
        "----------"

    ' 이 패킷을 보낸사람이 정상인가?
    If IsLogin = False Then
        Log "로그인 되지 않은 클라이언트가 요청함 OnChat()"
        Exit Sub
    End If
    
    ' 패킷 만들기
    Dim op As Packet
    Set op = New Packet
    
    op.InitSendPacket 50008
    op.WriteInt 0
    op.WriteInt 0
    op.WriteInt 0
    op.WriteInt 0

    socket.SendData op.GetPacketData
    DoEvents
End Sub


