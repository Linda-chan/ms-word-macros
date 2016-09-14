Option Explicit

Dim WithEvents DocEvents As Application

' Этот каталог можно переопределить. Желательно чтобы он был как
' можно короче. Или же в UNICODE можно использовать вариант записи
' "\\?\C:\Backups".
Const BACKUP_PATH = "\\?\D:\Word Backups"

' Всякие объявления и константы...
Private Const FILE_ATTRIBUTE_ARCHIVE& = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED& = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY& = &H10
Private Const FILE_ATTRIBUTE_HIDDEN& = &H2
Private Const FILE_ATTRIBUTE_NORMAL& = &H80
Private Const FILE_ATTRIBUTE_READONLY& = &H1
Private Const FILE_ATTRIBUTE_SYSTEM& = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY& = &H100

Private Const MAX_PATH& = 260

Private Const INVALID_HANDLE_VALUE& = -1

Private Const GENERIC_READ& = &H80000000
Private Const GENERIC_WRITE& = &H40000000

Private Const FILE_SHARE_READ& = &H1
Private Const FILE_SHARE_WRITE& = &H2

Private Const CREATE_NEW& = 1
Private Const CREATE_ALWAYS& = 2
Private Const OPEN_ALWAYS& = 4
Private Const OPEN_EXISTING& = 3
Private Const TRUNCATE_EXISTING& = 5

Private Const VER_PLATFORM_WIN32_NT& = 2
Private Const VER_PLATFORM_WIN32_WINDOWS& = 1
Private Const VER_PLATFORM_WIN32s& = 0

Private Const ERROR_SHARING_VIOLATION& = 32&
Private Const ERROR_ACCESS_DENIED& = 5&
Private Const ERROR_PATH_NOT_FOUND& = 3&
Private Const ERROR_FILE_NOT_FOUND& = 2&

Private Const MB_ABORTRETRYIGNORE& = &H2&
Private Const MB_APPLMODAL& = &H0&
Private Const MB_COMPOSITE& = &H2
Private Const MB_DEFAULT_DESKTOP_ONLY& = &H20000
Private Const MB_DEFBUTTON1& = &H0&
Private Const MB_DEFBUTTON2& = &H100&
Private Const MB_DEFBUTTON3& = &H200&
Private Const MB_DEFMASK& = &HF00&
Private Const MB_ICONASTERISK& = &H40&
Private Const MB_ICONEXCLAMATION& = &H30&
Private Const MB_ICONHAND& = &H10&
Private Const MB_ICONINFORMATION& = MB_ICONASTERISK
Private Const MB_ICONMASK& = &HF0&
Private Const MB_ICONQUESTION& = &H20&
Private Const MB_ICONSTOP& = MB_ICONHAND
Private Const MB_MISCMASK& = &HC000&
Private Const MB_MODEMASK& = &H3000&
Private Const MB_NOFOCUS& = &H8000&
Private Const MB_OK& = &H0&
Private Const MB_OKCANCEL& = &H1&
Private Const MB_PRECOMPOSED& = &H1
Private Const MB_RETRYCANCEL& = &H5&
Private Const MB_SETFOREGROUND& = &H10000
Private Const MB_SYSTEMMODAL& = &H1000&
Private Const MB_TASKMODAL& = &H2000&
Private Const MB_TYPEMASK& = &HF&
Private Const MB_USEGLYPHCHARS& = &H4
Private Const MB_YESNO& = &H4&
Private Const MB_YESNOCANCEL& = &H3&

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA_A
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName(0 To MAX_PATH - 1) As Byte
  cAlternate(0 To 13) As Byte
End Type

Private Type WIN32_FIND_DATA_W
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName(0 To MAX_PATH - 1) As Integer
  cAlternate(0 To 13) As Integer
End Type

Private Type OSVERSIONINFO_A
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion(0 To 127) As Byte
End Type

Private Type OSVERSIONINFO_W
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion(0 To 127) As Integer
End Type

Private Declare Function FindFirstFileA Lib "kernel32" (ByRef lpFileName As Any, ByRef lpFindFileData As WIN32_FIND_DATA_A) As Long
Private Declare Function FindFirstFileW Lib "kernel32" (ByRef lpFileName As Any, ByRef lpFindFileData As WIN32_FIND_DATA_W) As Long

Private Declare Function FindNextFileA Lib "kernel32" (ByVal hFindFile As Long, ByRef lpFindFileData As WIN32_FIND_DATA_A) As Long
Private Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, ByRef lpFindFileData As WIN32_FIND_DATA_W) As Long

Private Declare Function GetFileAttributesA Lib "kernel32" (ByRef lpFileName As Any) As Long
Private Declare Function GetFileAttributesW Lib "kernel32" (ByRef lpFileName As Any) As Long

Private Declare Function CreateFileA Lib "kernel32" (ByRef lpFileName As Any, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByRef lpFileName As Any, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long

Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, ByRef lpFileSizeHigh As Any) As Long

Private Declare Function CreateDirectoryA Lib "kernel32" (ByRef lpPathName As Any, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function CreateDirectoryW Lib "kernel32" (ByRef lpPathName As Any, ByVal lpSecurityAttributes As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function GetVersionExA Lib "kernel32" (ByRef lpVersionInformation As OSVERSIONINFO_A) As Long
Private Declare Function GetVersionExW Lib "kernel32" (ByRef lpVersionInformation As OSVERSIONINFO_W) As Long

Private Declare Function MessageBoxA Lib "user32" (ByVal hWnd As Long, ByRef lpText As Any, ByRef lpCaption As Any, ByVal wType As Long) As Long
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByRef lpText As Any, ByRef lpCaption As Any, ByVal wType As Long) As Long

Private Declare Function GetActiveWindow Lib "user32" () As Long

'====================================================================
Private Sub Document_New()
  Set DocEvents = Application
End Sub

Private Sub Document_Open()
  Set DocEvents = Application
End Sub

'====================================================================
Private Sub DocEvents_DocumentBeforeSave(ByVal Doc As Document, _
                                         SaveAsUI As Boolean, _
                                         Cancel As Boolean)
  Dim BackupPath As String
  Dim Buff() As Byte
  
  On Error Resume Next
  
  ' Хитрость. Если файл только что создан, то нам его бэкапить как
  ' раз и не надо. Его не существует.
  If Not IsFileExist(Doc.FullName) Then Exit Sub
  
  BackupPath = BACKUP_PATH
  
  If Not IsDirExist(BackupPath) Then MakeDir BackupPath
  If Err Then
    ErrorDisplay Err, BackupPath & vbCrLf & vbCrLf & _
                      "Can't create backup folder."
    Exit Sub
  End If
  
  If Right(BackupPath, 1) <> "\" Then BackupPath = BackupPath & "\"
  BackupPath = BackupPath & Replace( _
                                Replace( _
                                    Replace(Doc.FullName, "/", "~"), _
                                "\", "~"), _
                            ":", "~")
  
  If Not IsDirExist(BackupPath) Then MakeDir BackupPath
  If Err Then
    ErrorDisplay Err, BackupPath & vbCrLf & vbCrLf & _
                      "Can't create backup folder."
    Exit Sub
  End If
  
  If Right(BackupPath, 1) <> "\" Then BackupPath = BackupPath & "\"
  BackupPath = BackupPath & Format(Now, "yyyy\-mm\-dd hh\-nn\-ss") & ".BAK"
  
  ' Магия. CopyFile пытается открыть файл с записью (нафига?),
  ' поэтому ничего у него не получается. Поэтому сделаем вот
  ' так: сами прочитаем и сами сохраним...
  Buff = GetFileB(Doc.FullName)
  If Err Then
    ErrorDisplay Err, Doc.FullName & vbCrLf & vbCrLf & _
                      "Can't read source file." & _
                      vbCrLf & vbCrLf & BackupPath
    Exit Sub
  End If
  
  PutFileB BackupPath, Buff
  If Err Then
    ErrorDisplay Err, BackupPath & vbCrLf & vbCrLf & _
                      "Can't write file to backup location." & _
                      vbCrLf & vbCrLf & BackupPath
  End If
End Sub

'====================================================================
Private Function IsWinNT() As Boolean
  Dim OSV As OSVERSIONINFO_W
  
  OSV.dwOSVersionInfoSize = Len(OSV)
  GetVersionExW OSV
  
  IsWinNT = CBool(OSV.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

'====================================================================
Private Function IsDirExist(ByVal Path As String) As Boolean
  Dim hFind As Long
  Dim FD As WIN32_FIND_DATA_W
  Dim FileName() As Byte
  
  Path = Replace(Path, "/", "\")
  If Right(Path, 1) = "\" Then Path = Left(Path, Len(Path) - 1)
  
  FileName = Path & vbNullChar
  hFind = FindFirstFileW(FileName(0), FD)
  If hFind = 0 Or hFind = INVALID_HANDLE_VALUE Then
    'MsgBox hFind & " :: " & Err.LastDllError
    IsDirExist = False
  Else
    CloseHandle hFind
    If (FD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0 Then
      IsDirExist = True
    Else
      IsDirExist = False
    End If
  End If
End Function

'====================================================================
Private Function IsFileExist(ByVal Path As String) As Boolean
  Dim hFind As Long
  Dim FD As WIN32_FIND_DATA_W
  Dim FileName() As Byte
  
  Path = Replace(Path, "/", "\")
  
  FileName = Path & vbNullChar
  hFind = FindFirstFileW(FileName(0), FD)
  If hFind = 0 Or hFind = INVALID_HANDLE_VALUE Then
    'MsgBox hFind & " :: " & Err.LastDllError
    IsFileExist = False
  Else
    CloseHandle hFind
    If (FD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
      IsFileExist = True
    Else
      IsFileExist = False
    End If
  End If
End Function

'====================================================================
Public Function GetFileB(ByVal FileName As String) As Byte()
  Dim hFile As Long
  Dim bFileName() As Byte
  Dim FileSize As Long
  Dim BytesRead As Long
  Dim Buff() As Byte
  
  bFileName = FileName & vbNullChar
  hFile = CreateFileW(bFileName(0), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
  If hFile = 0 Or hFile = INVALID_HANDLE_VALUE Then
    Err.Raise Win32ErrorToVBAError(Err.LastDllError)
    Exit Function
  End If
  
  FileSize = GetFileSize(hFile, ByVal 0&)
  If FileSize < 0 Then
    CloseHandle hFile
    Err.Raise 57 ' I/O Error
  End If
  
  On Error Resume Next
  
  ReDim Buff(0 To FileSize - 1)
  If Err.Number <> 0 Then
    CloseHandle hFile
    
    On Error GoTo 0
    Err.Raise 7 ' Out of memory
  End If
  
  If FileSize = 0 Then
    CloseHandle hFile
    GetFileB = Buff
    Exit Function
  End If
  
  If ReadFile(hFile, Buff(0), FileSize, BytesRead, ByVal 0&) = 0 Then
    CloseHandle hFile
    Err.Raise 57 ' I/O Error
  End If
  
  CloseHandle hFile
  GetFileB = Buff
End Function

'====================================================================
Public Sub PutFileB(ByVal FileName As String, _
                    ByRef Data() As Byte)
  Dim hFile As Long
  Dim bFileName() As Byte
  Dim FileSize As Long
  Dim BytesWritten As Long
  
  bFileName = FileName & vbNullChar
  hFile = CreateFileW(bFileName(0), GENERIC_WRITE, 0, ByVal 0&, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
  If hFile = 0 Or hFile = INVALID_HANDLE_VALUE Then
    Err.Raise Win32ErrorToVBAError(Err.LastDllError)
    Exit Sub
  End If
  
  FileSize = UBound(Data) - LBound(Data) + 1
  
  If FileSize = 0 Then
    CloseHandle hFile
    Exit Sub
  End If
  
  If WriteFile(hFile, Data(0), FileSize, BytesWritten, ByVal 0&) = 0 Then
    CloseHandle hFile
    Err.Raise 57 ' I/O Error
  End If
  
  CloseHandle hFile
End Sub

'====================================================================
Public Sub MakeDir(ByVal Path As String)
  Dim bPath() As Byte
  
  If IsDirExist(Path) Then Exit Sub
  
  bPath = Path & vbNullChar
  If CreateDirectoryW(bPath(0), ByVal 0&) = 0 Then
    Err.Raise Win32ErrorToVBAError(Err.LastDllError)
    Exit Sub
  End If
End Sub

'====================================================================
Private Function Win32ErrorToVBAError(ByVal Number As Long) As Long
  Select Case Number
    Case ERROR_SHARING_VIOLATION
      Win32ErrorToVBAError = 70 ' VBAError_PermissionDenied
    Case ERROR_ACCESS_DENIED
      Win32ErrorToVBAError = 70 ' VBAError_PermissionDenied
    Case ERROR_PATH_NOT_FOUND
      Win32ErrorToVBAError = 76 ' VBAError_PathNotFound
    Case ERROR_FILE_NOT_FOUND
      Win32ErrorToVBAError = 53 ' VBAError_FileNotFound
    Case Else
      Win32ErrorToVBAError = 75 ' VBAError_PathFileAccess
  End Select
End Function

'====================================================================
Public Sub ErrorDisplay(ByVal ErrX As ErrObject, _
                        Optional ByVal Message As String = "Something happens.")
  Dim TXT As String
  
  If Message <> "" Then TXT = Message & vbCrLf & vbCrLf
  TXT = TXT & "Error number: " & FormatErrorNumber(ErrX.Number) & vbCrLf & _
              "Description:  " & ErrX.Description
  
  MyMsgBox TXT, vbCritical
End Sub

Private Function FormatErrorNumber(ByVal Number As Long) As String
  Dim HexNum As String
  
  HexNum = UCase(Hex(Number))
  HexNum = String(8 - Len(HexNum), "0") & HexNum
  
  FormatErrorNumber = CStr(Number) & " (0x" & HexNum & ")"
End Function

'====================================================================
Private Function MyMsgBox(ByVal Prompt As String, _
                          Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                          Optional ByVal Title As String) As VbMsgBoxResult
  Dim bPrompt() As Byte
  Dim bTitle() As Byte
  
  bPrompt = Prompt & vbNullChar
  If Title = "" Then
    bTitle = Application.Name & vbNullChar
  Else
    bTitle = Title & vbNullChar
  End If
  
  MyMsgBox = MessageBoxW(0, bPrompt(0), bTitle(0), Buttons Or MB_TASKMODAL)
End Function
