VERSION 5.00
Begin VB.Form Connectbox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server active connections"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "Connectbox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5160
      Top             =   720
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Status"
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Sniff connection"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Kill/Ban Host"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Kill (reason)"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fast Kill"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "Connectbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lastselected As Integer

Private Sub Command1_Click()
If List1.ListIndex > -1 Then
sock_out (List1.ListIndex + 1), vbCrLf + "+++ Admin was killed your connection, good bye!" + vbCrLf + String(512, " ") + String(10, Chr(10)) + vbCrLf
End If

End Sub

Private Sub Command2_Click()
If List1.ListIndex > -1 Then
sock_out (List1.ListIndex + 1), vbCrLf + "+++ Admin was killed your connection(reason:" + InputBox("Reason:", "Killing user") + "), good bye!" + vbCrLf + String(512, " ") + String(10, Chr(10)) + vbCrLf
End If


End Sub

Private Sub Command3_Click()
If List1.ListIndex > -1 Then
hostbans(0).count = hostbans(0).count + 1
hostbans(hostbans(0).count).host = mainfrm.sock(List1.ListIndex + 1).RemoteHostIP
hostbans(hostbans(0).count).own = "Admin"
hostbans(hostbans(0).count).reason = "Banned from local machine"
save_ban_r
sock_out (List1.ListIndex + 1), vbCrLf + "+++ Admin was killed your connection(banned), good bye!" + vbCrLf + String(512, " ") + String(10, Chr(10)) + vbCrLf
End If
End Sub

Private Sub Command4_Click()
If List1.ListIndex > -1 Then
sniffnode = List1.ListIndex + 1
snif.Show
End If
End Sub

Private Sub List1_Click()
lastselected = List1.ListIndex

'MsgBox List1.ListIndex
End Sub

Private Sub Timer1_Timer()
List1.Clear

For x = 1 To 10
If mainfrm.sock(x).State = sckConnected Then
If Len(Str(x)) = 2 Then
List1.AddItem "Node " + _
Replace(Str(x), Chr(32), "") + _
" :Used(" + mainfrm.sock(x).RemoteHostIP + ")" + Str(nodes(x).stage) + " >> " + nodes(x).user.username _
+ " (Idling:" + Replace(Str(nodes(x).idle), Chr(32), "") + ")"
Else
List1.AddItem "Node " + _
Replace(Str(x), Chr(32), "") + _
" :Used(" + mainfrm.sock(x).RemoteHostIP + ")" + Str(nodes(x).stage) + " >> " + nodes(x).user.username _
+ " (Idling:" + Replace(Str(nodes(x).idle), Chr(32), "") + ")"
End If

Else
If Len(Str(x)) = 2 Then
List1.AddItem "Node 0" + Replace(Str(x), Chr(32), "") + " :Available"
Else
List1.AddItem "Node " + Replace(Str(x), Chr(32), "") + " :Available"
End If

End If

Next x
List1.Selected(lastselected) = True
End Sub
