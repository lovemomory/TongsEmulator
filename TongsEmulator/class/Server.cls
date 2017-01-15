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
    
    ' 货肺款 技记 积己饶 立加 荐遏
    Dim client As Session
    Set client = New Session
    client.socket.Accept requestID
    
    ' 虐 = IP:PORT
    Dim clientKey As String
    clientKey = client.socket.RemoteHostIP & ":" & client.socket.RemotePort
    client.Key = clientKey
    
    ' 技记 府胶飘俊 眠啊
    ClientSession.Add client, clientKey
    
    Log "<< 立加 / IP: " & clientKey & " >>"
End Sub

