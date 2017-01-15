VERSION 5.00
Begin VB.Form formMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '단일 고정
   Caption         =   "통스통스 예물레이터"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   9135
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox textLog 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   1440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   1
      Text            =   "formMain.frx":0000
      Top             =   120
      Width           =   7515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "시작"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "formMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Command1.Enabled = False
    
    MainServer.Run
    Log "15003 포트로 소켓 리스닝 시작"
End Sub

Private Sub Form_Load()
    textLog.Text = "▶▶ 통스통스 예물레이터 GUI 시작 ◀◀" & vbCrLf
End Sub
