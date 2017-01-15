VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents socket As Winsock
Attribute socket.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set socket = New Winsock
End Sub

Public Sub Run()
    socket.Close
    socket.LocalPort = 15003
    socket.Listen
End Sub

Private Sub socket_ConnectionRequest(ByVal requestID As Long)
    socket.Close
    socket.Listen
    
    ' ���ο� ���� ������ ���� ����
    Dim client As Session
    Set client = New Session
    client.socket.Accept requestID
    
    ' Ű = IP:PORT
    Dim clientKey As String
    clientKey = client.socket.RemoteHostIP & ":" & client.socket.RemotePort
    client.Key = clientKey
    
    ' ���� ����Ʈ�� �߰�
    ClientSession.Add client, clientKey
    
    Log "<< ���� / IP: " & clientKey & " >>"
End Sub
