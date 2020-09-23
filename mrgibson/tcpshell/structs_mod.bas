Attribute VB_Name = "user_mod"
Type users_rec
username As String
pass As String
homedir As String
level As Integer
realname As String
comment As String
End Type

Public users(0 To 10000) As users_rec

Public Sub createuserdb(filename As String)
Dim fs As Integer
On Error GoTo errhandler
fs = FreeFile
users(1).username = "admin"
users(1).comment = "The generated admin account"
users(1).level = 999
users(1).realname = "The Admin"
users(1).pass = "admin"
users(1).homedir = "/home/admin"
Open filename For Output As fs
Print #fs, putuser(1)
Close fs
Exit Sub
errhandler:
MkDir (cfg.sysroot + "system\")
createuserdb (cfg.sysroot + "system\users.db")
End Sub

Public Function putuser(x As Integer) As String
putuser = users(x).username + ";"
putuser = putuser + users(x).pass + ";"
putuser = putuser + Replace(Str(users(x).level), Chr(32), "") + ";"
putuser = putuser + users(x).homedir + ";"
putuser = putuser + users(x).realname + ";"
putuser = putuser + users(x).comment + ";"
End Function
Public Function getuser(line As String, x As Integer) As Boolean
On Error GoTo errhandler
users(x).username = Mid(line, 1, InStr(line, ";") - 1)
line = Mid(line, InStr(line, ";") + 1, Len(line))
users(x).pass = Mid(line, 1, InStr(line, ";") - 1)
line = Mid(line, InStr(line, ";") + 1, Len(line))
users(x).level = Val(Mid(line, 1, InStr(line, ";") - 1))
line = Mid(line, InStr(line, ";") + 1, Len(line))
users(x).homedir = Mid(line, 1, InStr(line, ";") - 1)
line = Mid(line, InStr(line, ";") + 1, Len(line))
users(x).realname = Mid(line, 1, InStr(line, ";") - 1)
line = Mid(line, InStr(line, ";") + 1, Len(line))
users(x).comment = Mid(line, 1, InStr(line, ";") - 1)
getuser = True
Exit Function

errhandler:
getuser = False
End Function

Public Function write_userdb(filename As String) As Boolean
Dim fs As Integer
Dim count As Integer

On Error GoTo errhandler

Open filename For Output As fs
For count = 1 To users(0).level
Print #fs, putuser(count)
Next count

Close fs

write_userdb = True
errhandler:
write_userdb = False
End Function
Public Function read_userdb(filename As String) As Boolean
Dim line As String
Dim count As Integer
Dim fs As Integer

On Error GoTo errhandler
count = 1
users(0).level = 0
read_userdb = True
fs = FreeFile
Open filename For Input As fs

While Not EOF(fs)
Input #fs, line
If getuser(line, count) Then
count = count + 1
users(0).level = users(0).level + 1
End If
Wend
Close fs

Exit Function
errhandler:
read_userdb = False
End Function

Public Function seekuser(user As String, pass As String) As Integer
Dim count As Integer
For count = 1 To users(0).level
If users(count).username = user Then
If users(count).pass = pass Then
seekuser = count
Exit Function
End If
End If

Next count
seekuser = 0
End Function
