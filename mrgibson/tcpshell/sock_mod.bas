Attribute VB_Name = "sock_mod"
Public inbuffer(1 To 20) As String
Public outbuffer(1 To 20) As String

Type nodes_Rec
user As users_rec
idle As Integer
stage As Integer
End Type

Public nodes(1 To 10) As nodes_Rec

Public Sub init_sockets()
For x = 1 To 10
Load mainfrm.sock(x)
inbuffer(x) = ""
outbuffer(x) = ""
nodes(x).stage = 0
Next x
mainfrm.sock(0).LocalPort = 23
mainfrm.sock(0).Listen
cout "-- Server initialized (" + mainfrm.sock(0).LocalIP + ":80)" + vbCrLf
End Sub

Public Sub sock_out(node As Integer, content As String)
outbuffer(node) = outbuffer(node) + content
End Sub
Public Function who_cmd(node As Integer, param As String) As String
who_cmd = ""

If param = "" Then ' If params is empty , show generic who
For x = 1 To 10
If mainfrm.sock(x).State = sckConnected Then
If node = x Then
who_cmd = who_cmd + "[Node" + Str(x) + _
"] (You) u:" + nodes(x).user.username + vbCrLf
Else
who_cmd = who_cmd + "[Node" + Str(x) + _
"] (Idle:" + Replace(Str(nodes(x).idle), Chr(32), "") + _
") u:" + nodes(x).user.username + vbCrLf
End If

End If
Next x
Exit Function
End If

If param = "?" Then
who_cmd = "(?) Use 'who' to see a complete list or 'who #n' to" + vbCrLf + "(?) see just one node specific informations." + vbCrLf
Exit Function
End If

If mainfrm.sock(Val(param)).State = sckConnected Then
who_cmd = "(WHO) Specific informations about node " + param + vbCrLf
who_cmd = who_cmd + "[Node " + param + "] Username  :" + nodes(Val(param)).user.username + vbCrLf
who_cmd = who_cmd + "[Node " + param + "] Real Name :" + nodes(Val(param)).user.realname + vbCrLf
who_cmd = who_cmd + "[Node " + param + "] User Level:" + Replace(Str(nodes(Val(param)).user.level), Chr(32), "") + vbCrLf
who_cmd = who_cmd + "[Node " + param + "] Home Dir. :" + nodes(Val(param)).user.homedir + vbCrLf
who_cmd = who_cmd + "[Node " + param + "] Comment   :" + nodes(Val(param)).user.comment + vbCrLf
Exit Function
Else
who_cmd = "[Node " + param + "] Offline, nobody on this node." + vbCrLf
Exit Function
End If


End Function
