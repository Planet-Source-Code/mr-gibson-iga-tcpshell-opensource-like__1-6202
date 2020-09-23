Attribute VB_Name = "conio_mod"

Public Declare Function AllocConsole Lib "kernel32" () As Long

Public Declare Function FreeConsole Lib "kernel32" () As Long

Private Declare Function GetStdHandle Lib "kernel32" _
(ByVal nStdHandle As Long) As Long

Private Declare Function ReadConsole Lib "kernel32" Alias _
"ReadConsoleA" (ByVal hConsoleInput As Long, _
ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, _
lpNumberOfCharsRead As Long, lpReserved As Any) As Long

Private Declare Function SetConsoleMode Lib "kernel32" (ByVal _
hConsoleOutput As Long, dwMode As Long) As Long

Private Declare Function SetConsoleTextAttribute Lib _
"kernel32" (ByVal hConsoleOutput As Long, ByVal _
wAttributes As Long) As Long

Private Declare Function SetConsoleTitle Lib "kernel32" Alias _
"SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long

Private Declare Function WriteConsole Lib "kernel32" Alias _
"WriteConsoleA" (ByVal hConsoleOutput As Long, _
ByVal lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, _
lpNumberOfCharsWritten As Long, lpReserved As Any) As Long

Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&

Private Const FOREGROUND_BLUE = &H1
Private Const FOREGROUND_GREEN = &H2
Private Const FOREGROUND_RED = &H4
Private Const FOREGROUND_INTENSITY = &H8
Private Const BACKGROUND_BLUE = &H10
Private Const BACKGROUND_GREEN = &H20
Private Const BACKGROUND_RED = &H40
Private Const BACKGROUND_INTENSITY = &H80
Private Const ENABLE_LINE_INPUT = &H2
Private Const ENABLE_ECHO_INPUT = &H4
Private Const ENABLE_MOUSE_INPUT = &H10
Private Const ENABLE_PROCESSED_INPUT = &H1
Private Const ENABLE_WINDOW_INPUT = &H8
Private Const ENABLE_PROCESSED_OUTPUT = &H1
Private Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2
Private hConsoleIn As Long
Private hConsoleOut As Long
Private hConsoleErr As Long


Public Sub createconsole(title As String)
AllocConsole

SetConsoleTitle title
hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)
hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)
hConsoleErr = GetStdHandle(STD_ERROR_HANDLE)
End Sub
Public Function cin() As String
Dim sUserInput As String * 256
Call ReadConsole(hConsoleIn, sUserInput, Len(sUserInput), vbNull, vbNull)
cin = Left$(sUserInput, InStr(sUserInput, Chr$(0)) - 3)
End Function

Public Sub cout(szOut As String)
WriteConsole hConsoleOut, szOut, Len(szOut), vbNull, vbNull
End Sub


