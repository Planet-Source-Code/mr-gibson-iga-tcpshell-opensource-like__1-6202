VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mainfrm 
   BackColor       =   &H00C0C0C0&
   Caption         =   "TCP-Shell 0.5b"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   720
      Top             =   2520
   End
   Begin MSWinsockLib.Winsock sock 
      Index           =   0
      Left            =   2160
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub initialize()
Me.Show
logofrm.Left = Me.Width - logofrm.Width - 450
logofrm.Top = Me.Height - logofrm.Height - 600
Connectbox.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
For x = 1 To 10
sock(x).Close
Unload sock(x)
Next x

End Sub

Private Sub sock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim x As Integer
For x = 1 To 10
If sock(x).State = sckConnected Then
Else
If sniffnode > 0 Then
snif.Text1.Text = ""
End If

inbuffer(x) = ""
outbuffer(x) = ""
sock(x).Close
sock(x).Accept requestID
If scan4bans(sock(x).RemoteHostIP) = True Then
sock_out x, vbCrLf + "+++ Your connection is permban, good bye!" + vbCrLf + String(512, " ") + String(10, Chr(10)) + vbCrLf

Else
sock_out x, "TCPShell 0.5b by Yanick Bourbeau '29/01/1999 [noNVT-ASCII/L-ECHO] (discovery)" + vbCrLf
sock_out x, "Running on " + os_info + " " + cpu_info + " " + ram_info + " upt(" + uptime_info + ")" + vbCrLf
sock_out x, vbCrLf + "Welcome to " + cfg.sysname + ", owned by " + cfg.adminname + vbCrLf + vbCrLf
sock_out x, "Login:"
nodes(x).idle = 0
nodes(x).stage = 0
nodes(x).user.username = "Nobody"
End If
Exit Sub
End If
Next x
End Sub

Private Sub sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim buff As String
sock(Index).GetData buff
inbuffer(Index) = inbuffer(Index) + buff
End Sub

Private Sub Timer1_Timer()
Dim prebuff As String
Dim x As Integer

For x = 1 To 10
If sock(x).State = sckConnected Then

If nodes(x).stage = 0 Then
If nodes(x).idle = 60 Then
sock(x).Close
End If
End If

'outputs action
If Len(outbuffer(x)) > 0 Then
prebuff = Mid(outbuffer(x), 1, 512)
If InStr(prebuff, String(10, Chr(10))) > 0 Then
sock(x).Close
Else
sock(x).SendData prebuff
If sniffnode > 0 Then
snif.Text1.SelStart = Len(snif.Text1.Text)
snif.Text1.SelLength = 0
snif.Text1.SelText = prebuff
End If

End If
outbuffer(x) = Mid(outbuffer(x), 513, Len(outbuffer(x)))
End If
'inputs action
If InStr(inbuffer(x), Chr(13)) > 0 Then
inbuffer(x) = Replace(inbuffer(x), Chr(10), "")
parse x, Mid(inbuffer(x), 1, InStr(inbuffer(x), Chr(13)) - 1)
If sniffnode > 0 Then
snif.Text1.SelStart = Len(snif.Text1.Text)
snif.Text1.SelLength = 0
snif.Text1.SelText = Mid(inbuffer(x), 1, InStr(inbuffer(x), Chr(13)) - 1) + vbCrLf
End If
inbuffer(x) = Mid(inbuffer(x), InStr(inbuffer(x), Chr(13)) + 1, Len(inbuffer(x)))
Else
nodes(x).idle = nodes(x).idle + 1
End If

End If

Next x
End Sub
