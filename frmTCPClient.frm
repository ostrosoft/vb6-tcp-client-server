VERSION 5.00
Begin VB.Form frmTCPClient 
   Caption         =   "TCP Client"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   6465
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   315
      Left            =   5280
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox txtPort 
      Height          =   315
      Left            =   4320
      TabIndex        =   4
      Text            =   "22222"
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox txtHost 
      Height          =   315
      Left            =   600
      TabIndex        =   3
      Text            =   "localhost"
      Top             =   0
      Width           =   3255
   End
   Begin VB.TextBox txtSend 
      Enabled         =   0   'False
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   3000
      Width           =   5415
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5520
      TabIndex        =   1
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtRecv 
      Height          =   2535
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Port"
      Height          =   195
      Left            =   3960
      TabIndex        =   7
      Top             =   60
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Host"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   330
   End
End
Attribute VB_Name = "frmTCPClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents client As OSWINSCK.Winsock
Attribute client.VB_VarHelpID = -1

Private Sub cmdConnect_Click()
  If cmdConnect.Caption = "Connect" Then
    txtRecv.Text = ""
    client.Connect txtHost.Text, txtPort.Text
  Else
    If client.State <> sckClosed Then
      client.CloseWinsock
    End If
  End If
End Sub

Private Sub cmdSend_Click()
  client.SendData txtSend.Text & vbCrLf
End Sub

Private Sub Form_Load()
  Set client = New OSWINSCK.Winsock
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set client = Nothing
End Sub

Private Sub client_OnClose()
  client.CloseWinsock
  cmdConnect.Caption = "Connect"
  SetControls False
End Sub

Private Sub client_OnConnect()
  cmdConnect.Caption = "Disconnect"
  SetControls True
End Sub

Private Sub SetControls(ByVal bConnected As Boolean)
  txtHost.Enabled = Not bConnected
  txtPort.Enabled = Not bConnected
  txtSend.Enabled = bConnected
  cmdSend.Enabled = bConnected
End Sub

Private Sub client_OnDataArrival(ByVal bytesTotal As Long)
  Dim s As String
  client.GetData s
  txtRecv.Text = txtRecv.Text & s
End Sub

Private Sub client_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  MsgBox "Error " & Number & ": " & Description
End Sub
