Attribute VB_Name = "sys_mod"
Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
    End Type

Global osinfo As OSVERSIONINFO
Declare Function GetTickCount& Lib "kernel32" ()

Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
    End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    End Type


Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)


Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
    Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
    Private Const PROCESSOR_INTEL_386 = 386
    Private Const PROCESSOR_INTEL_486 = 486
    Private Const PROCESSOR_INTEL_PENTIUM = 586
    Private Const PROCESSOR_LEVEL_80386 As Long = 3
    Private Const PROCESSOR_LEVEL_80486 As Long = 4
    Private Const PROCESSOR_LEVEL_PENTIUM As Long = 5
    Private Const PROCESSOR_LEVEL_PENTIUMII As Long = 6
    Private Const VER_PLATFORM_WIN32_NT = 2
    Private Const VER_PLATFORM_WIN32_WINDOWS = 1


Public Type udtCPU
    lClockSpeed As Variant
    lProcType As Integer
    strProcLevel As String
    strProcRevision As String
    lNumberOfProcessors As Long
    End Type


Public Enum eVersion
    eWindowsNT = 1
    eWindows95_98 = 2
    eUnknown = 3
End Enum
                        
Public Function GetCPUInfo(ptCPUInfo As udtCPU)
    Dim tSYS As SYSTEM_INFO
    Dim intProcType As Integer
    Dim strProcLevel As String
    Dim strProcRevision As String
    Call GetSystemInfo(tSYS)


    Select Case tSYS.dwProcessorType
        Case PROCESSOR_INTEL_386: intProcType = 386
        Case PROCESSOR_INTEL_486: intProcType = 486
        Case PROCESSOR_INTEL_PENTIUM: intProcType = 586
    End Select


Select Case tSYS.wProcessorLevel
    Case PROCESSOR_LEVEL_80386: strProcLevel = "386"
    Case PROCESSOR_LEVEL_80486: strProcLevel = "486"
    Case PROCESSOR_LEVEL_PENTIUM: strProcLevel = "pI"
    Case PROCESSOR_LEVEL_PENTIUMII: strProcLevel = "pII/Pro"
End Select
strProcRevision = "r" & HiByte(tSYS.wProcessorRevision) & "s" & LoByte(tSYS.wProcessorRevision)


With ptCPUInfo
.lClockSpeed = GetCPUSpeed
.lNumberOfProcessors = tSYS.dwNumberOfProcessors
.lProcType = intProcType
.strProcLevel = IIf(strProcLevel = "", "0", strProcLevel)
.strProcRevision = IIf(strProcRevision = "", "0", strProcRevision)
End With
End Function


Private Function GetVersion() As eVersion
    Dim os As OSVERSIONINFO
    os.dwOSVersionInfoSize = Len(os)


    If GetVersionEx(os) Then


        If os.dwPlatformId = VER_PLATFORM_WIN32_NT Then
            GetVersion = eWindowsNT
        Else
            GetVersion = eWindows95_98
        End If
    Else
        GetVersion = eUnknown
    End If
End Function


Public Function HiByte(ByVal wParam As Integer) As Byte
    HiByte = (wParam And &HFF00&) \ (&H100)
End Function


Public Function LoByte(ByVal wParam As Integer) As Byte
    LoByte = wParam And &HFF&
End Function


Private Function GetCPUSpeed() As Variant
    Dim hKey As Long
    Dim lClockSpeed As Long
    Dim strKey As String


    If GetVersion = eWindowsNT Then
        strKey = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
        Call RegOpenKey(HKEY_LOCAL_MACHINE, strKey, hKey)
        Call RegQueryValueEx(hKey, "~MHz", 0, 0, lClockSpeed, 4)
        Call RegCloseKey(hKey)
        GetCPUSpeed = lClockSpeed
    Else
        GetCPUSpeed = ""
    End If
End Function
'-----------

Public Function cpu_info() As String
  Dim tCPU As udtCPU
    Call GetCPUInfo(tCPU)
    
    cpu_info = tCPU.lNumberOfProcessors
    cpu_info = cpu_info + "x" & tCPU.lProcType
    cpu_info = cpu_info + "(" + tCPU.strProcLevel + ")"
    cpu_info = cpu_info + tCPU.strProcRevision
    cpu_info = cpu_info + "[" & tCPU.lClockSpeed + "]"
End Function

Public Function os_info() As String
Dim s As String
Dim verNum As Long, verWord As Integer
     
        osinfo.dwOSVersionInfoSize = 148
        My& = GetVersionEx&(osinfo)
        
           If osinfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
                s$ = "Win9x "
            ElseIf myVer.dwPlatformId = VER_PLATFORM_WIN32_NT Then
                s$ = "WinNT "
            End If
        
        os_info = s$ & osinfo.dwMajorVersion & "." & osinfo.dwMinorVersion & "b" & (osinfo.dwBuildNumber And &HFFFF&)
                
        
End Function

Public Function ram_info() As String
Dim memsts As MEMORYSTATUS
Dim memory As Long

GlobalMemoryStatus memsts
memory = memsts.dwTotalPhys
ram_info = Format$(((memory \ 1024) \ 1024) + 1, "###") + "mb(RAM)"

End Function

Public Function sec2time(nSecs As Long) As String
On Error Resume Next
sec2time = 0
nDays = 0
nRest = 0
nHour = 0
nMinutes = 0
nSeconds = 0
nSeconds = nSecs
Do Until Int(nSeconds / 60) = nSeconds / 60
    nSeconds = nSeconds - 1
Loop
nMinutes = nSeconds / 60
nSeconds = nSecs - nSeconds
nTot = nMinutes
nTotHour = 0
If nMinutes >= 60 Then
    nHour = nMinutes
    Do Until Int(nHour / 60) = nHour / 60
        nHour = nHour - 1
    Loop
    nMinutes = nTot - nHour
    nTotHour = nHour / 60
End If
nTot = nTotHour
If nTotHour > 24 Then
    nDays = nTotHour
    Do Until Int(nDays / 24) = nDays / 24
        nDays = nDays - 1
    Loop
    nTotHour = nTot - nDays
    ntotdays = nDays / 24
End If
If CStr(ntotdays) = "" Then
sec2time = "0/" + Format(CStr(nTotHour) + ":" + CStr(nMinutes) + ":" + CStr(nSeconds), "hh:nn:ss")
Else
sec2time = CStr(ntotdays) + "/" + Format(CStr(nTotHour) + ":" + CStr(nMinutes) + ":" + CStr(nSeconds), "hh:nn:ss")
End If
End Function

Public Function uptime_info() As String
uptime_info = sec2time(Round(GetTickCount& / 1000))
End Function
