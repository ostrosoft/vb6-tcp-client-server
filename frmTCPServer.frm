VERSION 5.00
Begin VB.Form frmTCPServer 
   Caption         =   "TCP Server"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   6465
   Begin VB.CommandButton cmdStartClients 
      Caption         =   "Start Clients"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtNumberOfClients 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Text            =   "3"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtStatus 
      Height          =   3255
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   60
      Width           =   6375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "# of test clients"
      Height          =   195
      Left            =   1080
      TabIndex        =   3
      Top             =   3600
      Width           =   1080
   End
End
Attribute VB_Name = "frmTCPServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents server As OSWINSCK.Winsock
Attribute server.VB_VarHelpID = -1
Dim tcpListener() As clsTCPServerListener
Dim socketCount As Integer

Private Sub cmdStartClients_Click()
    Dim i As Integer
    Dim clientCount As Integer
    Dim frm As frmTCPClient
    clientCount = Val(txtNumberOfClients.Text)
    For i = 0 To clientCount - 1
        Set frm = New frmTCPClient
        frm.Left = i * frm.Width + 60
        frm.Top = Me.Height + 60
        frm.Show
    Next
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    
    Set server = New OSWINSCK.Winsock
    server.LocalPort = 22222
    txtStatus.Text = "Server is listening on port " & server.LocalPort & vbCrLf
    server.Listen
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    For i = 0 To socketCount - 1
        tcpListener(i).CloseWinsock
        Set tcpListener(i) = Nothing
    Next
    server.CloseWinsock
    Set server = Nothing
    
    End
End Sub

Private Sub server_OnConnectionRequest(ByVal requestID As Long)
    Dim socketIndex As Integer
    socketIndex = -1
    
    Dim i As Integer
    For i = 0 To socketCount - 1
        If tcpListener(i).State = sckClosed Then
            socketIndex = i
            Exit For
        End If
    Next
    If socketIndex = -1 Then
        socketIndex = socketCount
        socketCount = socketCount + 1
        ReDim Preserve tcpListener(socketCount)
        Set tcpListener(socketIndex) = New clsTCPServerListener
        tcpListener(socketIndex).SetCallback socketIndex, Me
    End If
    
    tcpListener(socketIndex).LocalPort = server.LocalPort
    tcpListener(socketIndex).Accept requestID
    
    tcpListener(socketIndex).SendData "connection request accepted by listener #" & socketIndex & vbCrLf
End Sub

Public Sub OnClose(ByVal Index As Long)
    tcpListener(Index).CloseWinsock
End Sub

Public Sub OnDataArrival(ByVal Index As Long, ByVal bytesTotal As Long)
    Dim s As String
    server.GetData s
    
    tcpListener(Index).GetData s
    txtStatus.Text = txtStatus.Text & "Listener #" & Index & " got data: " & s & vbCrLf
    tcpListener(Index).SendData "you sent " & s & vbCrLf
End Sub

Public Sub OnError(ByVal Index As Long, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub

Public Sub OnStatusChanged(ByVal Index As Long, ByVal Status As String)
    txtStatus.Text = txtStatus.Text & "Listener #" & Index & ": " & Status & vbCrLf
End Sub

Private Sub server_OnStatusChanged(ByVal Status As String)
    txtStatus.Text = txtStatus.Text & "Server: " & Status & vbCrLf
End Sub

Private Sub txtStatus_Change()
    txtStatus.SelStart = Len(txtStatus.Text)
End Sub
