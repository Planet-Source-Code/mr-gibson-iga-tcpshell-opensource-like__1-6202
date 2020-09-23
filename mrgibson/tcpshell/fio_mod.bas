Attribute VB_Name = "fio_mod"
Global sysroot As String 'the root
   Declare Function FindFirstFile Lib "kernel32" Alias _
   "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData _
   As WIN32_FIND_DATA) As Long

   Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
   (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

   Declare Function GetFileAttributes Lib "kernel32" Alias _
   "GetFileAttributesA" (ByVal lpFileName As String) As Long

   Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) _
   As Long

   Declare Function FileTimeToLocalFileTime Lib "kernel32" _
   (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
     
   Declare Function FileTimeToSystemTime Lib "kernel32" _
   (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

   Public Const MAX_PATH = 260
   Public Const MAXDWORD = &HFFFF
   Public Const INVALID_HANDLE_VALUE = -1
   Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
   Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
   Public Const FILE_ATTRIBUTE_HIDDEN = &H2
   Public Const FILE_ATTRIBUTE_NORMAL = &H80
   Public Const FILE_ATTRIBUTE_READONLY = &H1
   Public Const FILE_ATTRIBUTE_SYSTEM = &H4
   Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

   Type FILETIME
     dwLowDateTime As Long
     dwHighDateTime As Long
   End Type

   Type WIN32_FIND_DATA
     dwFileAttributes As Long
     ftCreationTime As FILETIME
     ftLastAccessTime As FILETIME
     ftLastWriteTime As FILETIME
     nFileSizeHigh As Long
     nFileSizeLow As Long
     dwReserved0 As Long
     dwReserved1 As Long
     cFileName As String * MAX_PATH
     cAlternate As String * 14
   End Type

   Type SYSTEMTIME
     wYear As Integer
     wMonth As Integer
     wDayOfWeek As Integer
     wDay As Integer
     wHour As Integer
     wMinute As Integer
     wSecond As Integer
     wMilliseconds As Integer
   End Type

   Public Function StripNulls(OriginalStr As String) As String
      If (InStr(OriginalStr, Chr(0)) > 0) Then
         OriginalStr = Left(OriginalStr, _
          InStr(OriginalStr, Chr(0)) - 1)
      End If
      StripNulls = OriginalStr
   End Function

 



   Function FindFilesAPI(path As String, SearchStr As String, _
    FileCount As Integer, DirCount As Integer)
   Dim filename As String
   Dim DirName As String
   Dim dirNames() As String
   Dim nDir As Integer
   Dim i As Integer
   Dim hSearch As Long
   Dim WFD As WIN32_FIND_DATA
   Dim Cont As Integer
   Dim FT As FILETIME
   Dim ST As SYSTEMTIME
   Dim DateCStr As String, DateMStr As String
     nx = 0
   If Right(path, 1) <> "\" Then path = path & "\"
   nDir = 0
   ReDim dirNames(nDir)
   Cont = True
   hSearch = FindFirstFile(path & "*", WFD)
   If hSearch <> INVALID_HANDLE_VALUE Then
      Do While Cont
         DirName = StripNulls(WFD.cFileName)
         If (DirName <> ".") And (DirName <> "..") Then
            If GetFileAttributes(path & DirName) And _
             FILE_ATTRIBUTE_DIRECTORY Then
               dirNames(nDir) = DirName
               DirCount = DirCount + 1
               nDir = nDir + 1
               ReDim Preserve dirNames(nDir)
            End If
         End If
         Cont = FindNextFile(hSearch, WFD)
      Loop
      Cont = FindClose(hSearch)
   End If

   hSearch = FindFirstFile(path & SearchStr, WFD)
   Cont = True
   If hSearch <> INVALID_HANDLE_VALUE Then
      While Cont
         filename = StripNulls(WFD.cFileName)
            If (filename <> ".") And (filename <> "..") And _
              ((GetFileAttributes(path & filename) And _
               FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
            FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * _
             MAXDWORD) + WFD.nFileSizeLow
            FileCount = FileCount + 1
            
            
            
           FileTimeToLocalFileTime WFD.ftCreationTime, FT
           FileTimeToSystemTime FT, ST
           DateCStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
              " " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
           FileTimeToLocalFileTime WFD.ftLastWriteTime, FT
           FileTimeToSystemTime FT, ST
           DateMStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
              " " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
          End If
         Cont = FindNextFile(hSearch, WFD)
      Wend
      Cont = FindClose(hSearch)
   End If

   End Function
