Attribute VB_Name = "init_mod"
Public Sub load_users()
Dim buff As String
cout "-- Reading Users database "


If exists(cfg.sysroot + "system\users.db") = False Then
createuserdb (cfg.sysroot + "system\users.db")
cout "[Created]" + vbCrLf + vbCrLf
cout "****IMPORTANT****************************" + vbCrLf
cout "The system not found users database, the " + vbCrLf
cout "system have created it by default. This  " + vbCrLf
cout "data only contain one user(admin) and the" + vbCrLf
cout "default password is 'admin'.Its important" + vbCrLf
cout "to log first and change your password." + vbCrLf
cout "Use the 'passwd' cmd to change it. " + vbCrLf
cout "****IMPORTANT****************************" + vbCrLf + vbCrLf
cout "(If you agreed press enter)"
cin

If read_userdb(cfg.sysroot + "system\users.db") Then
cout "-- Reading Users database [Found] OK"
Else
cout "++ Fatal Error(-1): Unable to read users_database" + vbCrLf
cout "System stopped, press Enter to leave"
cin
End
End If
Else
cout "[Found] "
If read_userdb(cfg.sysroot + "system\users.db") Then
cout "OK" + vbCrLf
Else
cout "++ Fatal Error(-1): Unable to read users_database" + vbCrLf
cout "System stopped, press Enter to leave"
cin
End
End If
End If

End Sub

Public Sub init_msg()
cout "-- System Constants {" + vbCrLf + vbCrLf
cout "    - Max Sockets (in same time) :'10'" + vbCrLf
cout "    - Speed per socket           :'1024bytes/sec'" + vbCrLf
cout "    - Server listening port      :'23'" + vbCrLf
cout "    - Hosts allowed              :'ALL'" + vbCrLf
cout "    - Terminal Mode              :'No-NVT compliant'" + vbCrLf + vbCrLf
cout "                     }" + vbCrLf

End Sub
