VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTCPServerListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Callback As frmTCPServer
Private WithEvents listener As OSWINSCK.Winsock
Attribute listener.VB_VarHelpID = -1
Private m_Index As Long

'===== PROPERTIES =====
Public Property Get BytesReceived() As Long
  BytesReceived = listener.BytesReceived
End Property

Public Property Get LocalHostName() As String
  LocalHostName = listener.LocalHostName
End Property

Public Property Get LocalIP() As String
  LocalIP = listener.LocalIP
End Property

Public Property Get LocalPort() As Long
  LocalPort = listener.LocalPort
End Property
Public Property Let LocalPort(NewValue As Long)
  listener.LocalPort = NewValue
End Property

Public Property Get Protocol() As ProtocolConstants
  Protocol = listener.Protocol
End Property
Public Property Let Protocol(NewValue As ProtocolConstants)
  listener.Protocol = NewValue
End Property

Public Property Get RemoteHost() As String
  RemoteHost = listener.RemoteHost
End Property
Public Property Let RemoteHost(NewValue As String)
  listener.RemoteHost = NewValue
End Property

Public Property Get RemoteHostIP() As String
  RemoteHostIP = listener.RemoteHostIP
End Property

Public Property Get RemotePort() As Long
  RemotePort = listener.RemotePort
End Property
Public Property Let RemotePort(NewValue As Long)
  listener.RemotePort = NewValue
End Property

Public Property Get SocketHandle() As Long
  SocketHandle = listener.SocketHandle
End Property

Public Property Get State() As StateConstants
  State = listener.State
End Property

Public Property Get Status() As String
  Status = listener.Status
End Property
Public Property Let Status(ByVal strTemp As String)
  listener.Status = NewValue
End Property

Public Property Get Tag() As String
  Tag = listener.Tag
End Property
Public Property Let Tag(ByVal vNewValue As String)
  listener.Tag = vNewValue
End Property

'===== METHODS =====
Public Sub Accept(requestID As Long)
    listener.Accept (requestID)
End Sub

Public Sub CloseWinsock()
    listener.CloseWinsock
End Sub

Public Sub Connect(Optional RemoteHost As Variant, Optional RemotePort As Variant)
    listener.Connect RemoteHost, RemotePort
End Sub

Public Sub GetData(data As Variant, Optional vtype As Variant, Optional maxLen As Variant)
    listener.GetData data, vtype, maxLen
End Sub

Public Sub SendData(data As Variant)
    listener.SendData data
End Sub

Public Sub SetCallback(ByVal Index As Long, ByRef frm As frmTCPServer)
    m_Index = Index
    Set m_Callback = frm
End Sub

Private Sub Class_Initialize()
    Set listener = New OSWINSCK.Winsock
End Sub

Private Sub Class_Terminate()
    Set listener = Nothing
End Sub

Private Sub listener_OnClose()
    m_Callback.OnClose m_Index
End Sub

Private Sub listener_OnDataArrival(ByVal bytesTotal As Long)
    m_Callback.OnDataArrival m_Index, bytesTotal
End Sub

Private Sub listener_OnStatusChanged(ByVal Status As String)
    m_Callback.OnStatusChanged m_Index, Status
End Sub
