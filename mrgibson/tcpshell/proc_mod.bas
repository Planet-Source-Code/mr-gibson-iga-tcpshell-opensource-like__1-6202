Attribute VB_Name = "proc_mod"
Type proclist_rec
owner As Integer
begin As String
'--------
dwSize As String
cntUsage As String
th32ProcessID As String
th32DefaultHeapID As String
th32ModuleID As String
cntThreads As String
th32ParentProcessID As String
pcPriClassBase As String
dwFlags As String
szexeFile As String
exename As String
End Type

Public procs(1 To 9999) As proclist_rec


Const MAX_PATH& = 260


Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long


Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long


Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long


Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long


Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long


Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
    End Type

Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long


Public Function init_process() As Boolean

    Const PROCESS_ALL_ACCESS = 0
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim AppKill As Boolean
    Dim appCount As Integer
    Dim i As Integer
    On Local Error GoTo Finish
    appCount = 0
    
    Const TH32CS_SNAPPROCESS As Long = 2&
    
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
    procs(0).owner = 0
    
    Do While rProcessFound
        i = InStr(1, uProcess.szexeFile, Chr(0))
        
        procs(procs(0).owner).szexeFile = LCase$(Left$(uProcess.szexeFile, i - 1))
        procs(procs(0).owner).begin = Str(time)
        procs(procs(0).owner).cntThreads = Str(uProcess.cntThreads)
        procs(procs(0).owner).cntUsage = Str(uProcess.cntUsage)
        procs(procs(0).owner).dwFlags = Str(uProcess.dwFlags)
        procs(procs(0).owner).dwSize = Str(uProcess.dwSize)
        procs(procs(0).owner).pcPriClassBase = Str(uProcess.pcPriClassBase)
        procs(procs(0).owner).th32DefaultHeapID = Str(uProcess.th32DefaultHeapID)
        procs(procs(0).owner).th32ModuleID = Str(uProcess.th32ModuleID)
        procs(procs(0).owner).th32ParentProcessID = Str(uProcess.th32ParentProcessID)
        procs(procs(0).owner).th32ProcessID = Str(uProcess.th32ProcessID)
        procs(procs(0).owner).owner = 999
        'List1.AddItem Hex(uProcess.th32ProcessID) + (szExename)
        'If Right$(szExename, Len(myName)) = LCase$(myName) Then
         '   KillApp = True
          '  appCount = appCount + 1
           ' myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
           ' AppKill = TerminateProcess(myProcess, exitCode)
           ' Call CloseHandle(myProcess)
        'End If


        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop


    Call CloseHandle(hSnapshot)
Finish:
End Function

