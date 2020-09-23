Attribute VB_Name = "secu_mod"
Type restrict_cmd_rec
name As String
level As Integer
End Type

Public rcmds(0 To 1000) As restrict_cmd_rec

Type hostban_rec
host As String
own As String
reason As String
count As Integer
End Type
Public hostbans(0 To 1000) As hostban_rec

Public Function load_ban_r() As Boolean
Dim line As String
Dim fs As Integer

On Error GoTo errhandler
load_ban_r = True
hostbans(0).count = 0
fs = FreeFile
Open (cfg.sysroot + "security\hosts.ban") For Input As fs
While Not EOF(fs)
Input #fs, line
hostbans(0).count = hostbans(0).count + 1
hostbans(hostbans(0).count).host = Mid(line, 1, InStr(line, ";") - 1)
line = Mid(line, InStr(line, ";") + 1, Len(line))
hostbans(hostbans(0).count).own = Mid(line, 1, InStr(line, ";") - 1)
line = Mid(line, InStr(line, ";") + 1, Len(line))
hostbans(hostbans(0).count).reason = line
Wend
Close fs
Exit Function
errhandler:
load_ban_r = False
End Function

Public Function save_ban_r() As Boolean
Dim line As String
Dim fs As Integer
Dim x As Integer

On Error GoTo errhandler
save_ban_r = True
fs = FreeFile
Open (cfg.sysroot + "security\hosts.ban") For Output As fs
For x = 1 To hostbans(0).count
Print #fs, hostbans(x).host + ";" + hostbans(x).own + ";" + hostbans(x).reason
Next x

Close fs
Exit Function
errhandler:
save_ban_r = False
End Function
Public Function scan4bans(host As String) As Boolean
Dim x As Integer
scan4bans = False
For x = 1 To hostbans(0).count
If host = hostbans(x).host Then 'banned host found
                           scan4bans = True
                           Exit Function
                        End If

Next x
End Function
Public Function load_cmd_r() As Boolean
Dim line As String
Dim fs As Integer

On Error GoTo errhandler
load_cmd_r = True
rcmds(0).level = 0
fs = FreeFile
Open (cfg.sysroot + "security\restrict.cmd") For Input As fs
While Not EOF(fs)
Input #fs, line
rcmds(0).level = rcmds(0).level + 1
rcmds(rcmds(0).level).name = Mid(line, 1, InStr(line, ";") - 1)
rcmds(rcmds(0).level).level = Val(Mid(line, InStr(line, ";") + 1, Len(line)))
Wend
Close fs
Exit Function
errhandler:
load_cmd_r = False
End Function

Public Function create_cmd_r() As Boolean
Dim folder As String
Dim fs As Integer

create_cmd_r = True
On Error GoTo errhandler
folder = Dir(cfg.sysroot + "security\", vbDirectory)
If folder = "" Then
    MkDir cfg.sysroot + "security\"
End If
fs = FreeFile
Open (cfg.sysroot + "security\restrict.cmd") For Output As fs
Print #fs, "who;0"
Print #fs, "kill;999";
Close fs
Exit Function
errhandler:
create_cmd_r = False
End Function
Public Function create_ban_r() As Boolean
Dim folder As String
Dim fs As Integer

create_ban_r = True
On Error GoTo errhandler
folder = Dir(cfg.sysroot + "security\", vbDirectory)
If folder = "" Then
    MkDir cfg.sysroot + "security\"
End If
fs = FreeFile
Open (cfg.sysroot + "security\hosts.ban") For Output As fs
Print #fs, "you.are.in.my.blacklist;admin;a generated ban(remove them)"
Close fs
Exit Function
errhandler:
create_ban_r = False
End Function

Public Function init_secu()
cout "-- Init security modules" + vbCrLf + vbCrLf
cout "     + Restrict.cmd ..."

If load_cmd_r = True Then
cout "[OK]" + vbCrLf
Else
If create_cmd_r = True Then
cout "[CREATED]" + vbCrLf
Else
cout "[FAIL]" + vbCrLf + vbCrLf
cout "Problem: Unable to create file/directory, ensure your paths(cfg) dont finish by \" + vbCrLf
cout "++ Program Halted (press enter to exit)"
cin
FreeConsole
End
End If
End If

cout "     + hosts.ban ..."
If load_ban_r = True Then
cout "[OK]" + vbCrLf
Else
If create_ban_r = True Then
cout "[CREATED]" + vbCrLf
Else
cout "[FAIL]" + vbCrLf + vbCrLf
cout "Problem: Unable to create file/directory, ensure your paths(cfg) dont finish by \" + vbCrLf
cout "++ Program Halted (press enter to exit)"
cin
FreeConsole
End
End If

End If
End Function
