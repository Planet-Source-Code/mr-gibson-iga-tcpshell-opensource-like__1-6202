Attribute VB_Name = "cfg_mod"
Type cfg_rec
sysname As String
adminname As String
workdir As String
sysroot As String
End Type

Public cfg As cfg_rec

Public Function loadcfg() As Boolean
Dim fs As Integer
Dim line As String

On Error GoTo errhandler
fs = FreeFile
loadcfg = True
Open "system.cfg" For Input As fs
Input #fs, line
Close fs

cfg.adminname = Mid(line, 1, InStr(line, ";") - 1)
line = Mid(line, InStr(line, ";") + 1, Len(line))
cfg.sysname = Mid(line, 1, InStr(line, ";") - 1)
line = Mid(line, InStr(line, ";") + 1, Len(line))
cfg.workdir = Mid(line, 1, InStr(line, ";") - 1)
line = Mid(line, InStr(line, ";") + 1, Len(line))
cfg.sysroot = Mid(line, 1, InStr(line, ";") - 1)
line = Mid(line, InStr(line, ";") + 1, Len(line))

Exit Function
errhandler:
loadcfg = False
End Function

Public Sub savecfg()
Dim fs As Integer
fs = FreeFile
Open "system.cfg" For Output As fs
Print #fs, cfg.adminname + ";" + cfg.sysname + ";" + cfg.workdir + ";" + cfg.sysroot + ";"
Close fs


End Sub

Public Sub init_config()
cout "-- Reading System Configuration "
If loadcfg = True Then
cout "[Found] OK" + vbCrLf
Else
cout "[Created]"
cfgfrm.Show 1
cout " OK" + vbCrLf
End If

cout vbCrLf + "-- Config loaded {" + vbCrLf + vbCrLf
cout "    - Admin name    :'" + cfg.adminname + "'" + vbCrLf
cout "    - System name   :'" + cfg.sysname + "'" + vbCrLf
cout "    - Working Dir   :'" + cfg.workdir + "'" + vbCrLf
cout "    - System Root   :'" + cfg.sysroot + "'" + vbCrLf + vbCrLf
cout "                 }" + vbCrLf + vbCrLf


End Sub


