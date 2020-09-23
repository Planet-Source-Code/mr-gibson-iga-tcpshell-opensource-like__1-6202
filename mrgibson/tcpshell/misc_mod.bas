Attribute VB_Name = "misc_mod"
Public rx As Integer
Public sx As Integer


Public Function exists(filename As String) As Boolean
Dim fs As Integer
fs = FreeFile
On Error GoTo errhandler
Open filename For Input As fs
Close fs
exists = True
Exit Function
errhandler:
exists = False
End Function

Public Function Round(nNumber) As Double
    On Error Resume Next
    n1 = nNumber
    n2 = Int(nNumber) + 1
    If n2 - n1 < 0.5 Then
        Round = n2
    Else
        Round = n2 - 1
    End If
    
End Function
