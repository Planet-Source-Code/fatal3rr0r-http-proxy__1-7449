VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HTTP Proxy"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   1980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   30
      Width           =   735
   End
   Begin VB.TextBox txtPort 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Text            =   "80"
      Top             =   30
      Width           =   615
   End
   Begin MSWinsockLib.Winsock wsTCP 
      Index           =   0
      Left            =   960
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsProxy 
      Index           =   0
      Left            =   1320
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stopped"
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   435
      Width           =   1950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Port"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   285
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim s(255) As String
Dim h(255) As String
Dim p(255) As String
Dim i As Integer

Private Sub cmdStart_Click()
  If cmdStart.Caption = "Start" Then
    wsTCP(0).LocalPort = txtPort
    wsTCP(0).Listen
    lblStatus = "Running..."
    cmdStart.Caption = "Stop"
  Else
    cmdStart.Caption = "Start"
    wsTCP(0).Close
    lblStatus = "Stopped"
  End If
End Sub

Private Sub wsProxy_Close(Index As Integer)
  On Error Resume Next
  Unload wsProxy(Index)
  wsTCP(Index).SendData p(Index)
End Sub

Private Sub wsProxy_Connect(Index As Integer)
  wsProxy(Index).SendData s(Index)
End Sub

Private Sub wsProxy_DataArrival(Index As Integer, ByVal bytesTotal As Long)
  wsProxy(Index).GetData h(Index)
  Debug.Print "(" & Index & ") " & h(Index)
  p(Index) = p(Index) & h(Index)
End Sub

Private Sub wsProxy_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  Debug.Print "(" & Index & ") Error " & Number & ": " & Description
  Unload wsProxy(Index)
End Sub

Private Sub wsTCP_Close(Index As Integer)
  Unload wsTCP(Index)
End Sub

Private Sub wsTCP_ConnectionRequest(Index As Integer, ByVal requestID As Long)
  i = i + 1
  Load wsTCP(i)
  Load wsProxy(i)
  wsTCP(i).Accept requestID
End Sub

Private Sub wsTCP_DataArrival(Index As Integer, ByVal bytesTotal As Long)
  wsTCP(Index).GetData s(Index)
  Debug.Print "(" & Index & ") " & s(Index)
  Dim strHost As String, iPort As Integer
  iPort = 80
  If InStr(UCase(s(Index)), "GET ") > 0 Then
    strHost = Mid(s(Index), InStr(UCase(s(Index)), "GET ") + 4)
  ElseIf InStr(UCase(s(Index)), "PUT ") > 0 Then
    strHost = Mid(s(Index), InStr(UCase(s(Index)), "PUT ") + 4)
  Else
    wsTCP(Index).SendData "Mailformed HTTP request"
    Exit Sub
  End If
  strHost = Left(strHost, InStr(strHost, " ") - 1)
  If InStr(strHost, "://") <> 0 Then strHost = Mid(strHost, InStr(strHost, "://") + 3)
  If InStr(strHost, ":") <> 0 Then
    iPort = Val(Mid(strHost, InStr(strHost, ":") + 1))
    strHost = Left(strHost, InStr(strHost, ":") - 1)
  End If
  If InStr(strHost, "/") > 0 Then strHost = Left(strHost, InStr(strHost, "/") - 1)
  With wsProxy(Index)
    .RemoteHost = strHost
    .RemotePort = iPort
    .Connect
  End With
End Sub

Private Sub wsTCP_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  Debug.Print "(" & Index & ") Error " & Number & ": " & Description
  Unload wsTCP(Index)
End Sub

Private Sub wsTCP_SendComplete(Index As Integer)
  wsTCP(Index).Close
End Sub
