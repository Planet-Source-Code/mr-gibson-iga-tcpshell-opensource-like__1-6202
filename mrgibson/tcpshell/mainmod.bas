Attribute VB_Name = "mainmod"
Public sniffnode As Integer

Sub Main()

createconsole ("TCPShell 0.5b Console")
cout "TCP-Shell 0.5b (Init_console) by Yanick Bourbeau '99" + vbCrLf
cout "----------------------------------------------------" + vbCrLf
cout vbCrLf
cout "Setting system..." + vbCrLf + vbCrLf
init_config ' reading/writing config file
load_users 'loading/creating users database
init_secu ' security initialisation
init_msg
cout vbCrLf + "-- Initializing MDI_Form and sockets array ..." + vbCrLf
init_sockets
FreeConsole
mainfrm.Visible = True
mainfrm.initialize
End Sub
 
'nodes.stages Description
'0 = just connected/not logged
'1 = user entered, pass query
'2 = access granted - normal user

Public Sub parse(node As Integer, command As String)
Dim ret As Integer

If nodes(node).stage = 0 Then
nodes(node).user.username = command
nodes(node).stage = 1
sock_out node, "Password:"
Exit Sub
End If

If nodes(node).stage = 1 Then
nodes(node).user.pass = command
ret = seekuser(nodes(node).user.username, nodes(node).user.pass)
If ret = 0 Then
sock_out node, vbCrLf + "Invalid username or/and password" + vbCrLf
sock_out node, "Login:"
nodes(node).stage = 0
Else
nodes(node).user = users(ret)
sock_out node, vbCrLf + vbCrLf + "Access granted, welcome back " + nodes(node).user.username + vbCrLf + vbCrLf
nodes(node).stage = 2
sock_out node, "[" + Str(Time) + "] #"
End If
Exit Sub
End If

If nodes(node).stage = 2 Then
interpret node, command
sock_out node, "[" + Str(Time) + "] #"
Exit Sub
End If

End Sub


Public Sub interpret(node As Integer, command As String)
Dim x As Integer
nodes(node).idle = 0
For x = 1 To rcmds(0).level
If InStr(command, rcmds(x).name + Chr(32)) = 1 Then
If nodes(node).user.level < rcmds(x).level Then
sock_out node, "(ERROR) You are not allowed to use this cmd" + vbCrLf + vbCrLf
Exit Sub
End If
Else
If command = rcmds(x).name Then
If nodes(node).user.level < rcmds(x).level Then
sock_out node, "(ERROR) You are not allowed to use this cmd" + vbCrLf + vbCrLf
Exit Sub
End If
End If
End If

Next x

If InStr(command, "who") = 1 Then
sock_out node, who_cmd(node, Mid(command, 5, Len(command)))
Exit Sub
End If
If command <> "" Then
sock_out node, "(Error) Cmd not found" + vbCrLf
End If
End Sub
