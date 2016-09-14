Option Explicit

Dim WithEvents DocEvents As Application

' Этот каталог можно переопределить. Желательно чтобы он был как
' можно короче. Или же в UNICODE можно использовать вариант записи
' "\\?\C:\Backups".
Const BACKUP_PATH = "D:\Word Backups"

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
  Buff = GetFileB(Doc.FullName, True)
  If Err Then
    ErrorDisplay Err, Doc.FullName & vbCrLf & vbCrLf & _
                      "Can't read source file." & _
                      vbCrLf & vbCrLf & BackupPath
    Exit Sub
  End If
  
  PutFileB BackupPath, Buff, True
  If Err Then
    ErrorDisplay Err, BackupPath & vbCrLf & vbCrLf & _
                      "Can't write file to backup location." & _
                      vbCrLf & vbCrLf & BackupPath
  End If
End Sub

'====================================================================
Private Function IsDirExist(ByVal Path As String) As Boolean
  Dim TXT As String
  
  On Error Resume Next
  
  TXT = Dir(Path, vbArchive + vbDirectory + vbHidden + _
                  vbNormal + vbReadOnly + vbSystem)
  IsDirExist = CBool(TXT <> "")
End Function

Private Function IsFileExist(ByVal Path As String) As Boolean
  Dim TXT As String
  
  On Error Resume Next
  
  TXT = Dir(Path, vbArchive + vbHidden + _
                  vbNormal + vbReadOnly + vbSystem)
  IsFileExist = CBool(TXT <> "")
End Function

'====================================================================
Public Function GetFileB(ByVal FileName As String, _
                         ByVal RaiseErrors As Boolean) As Byte()
  Dim hFile As Long
  
  On Error GoTo hError
  
  If Not IsFileExist(FileName) Then Err.Raise 53
  
  hFile = FreeFile
  Open FileName For Binary Access Read As #hFile
  
  ReDim GetFileB(0 To LOF(hFile) - 1)
  If LOF(hFile) > 0 Then Get #hFile, , GetFileB
  
  Close #hFile
  Exit Function
  
hError:
  If hFile > 0 Then Close #hFile
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub PutFileB(ByVal FileName As String, _
                    ByRef Data() As Byte, _
                    ByVal RaiseErrors As Boolean)
  Dim hFile As Long
  
  On Error GoTo hError
  
  If IsFileExist(FileName) Then Kill FileName
  
  hFile = FreeFile
  Open FileName For Binary Access Write As #hFile
  
  Put #hFile, , Data
  
  Close #hFile
  Exit Sub
  
hError:
  If hFile > 0 Then Close #hFile
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'====================================================================
Public Sub MakeDir(ByVal Path As String)
  MkDir Path
End Sub

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
  If Title = "" Then Title = Application.Name
  
  MyMsgBox = MsgBox(Prompt, Buttons, Title)
End Function
