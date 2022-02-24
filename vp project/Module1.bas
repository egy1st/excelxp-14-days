Attribute VB_Name = "Module1"
Declare Function ExitWindowsEx Lib "user32.dll" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long


Public Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(128) As Byte
End Type

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type



Public DriveLetter As String
Public PathString As String

Const MAX_TEXT_SIZE = 50
Dim hwndCurrent As Long
Dim sClass As String
Dim sCaption As String
Dim lRetVal As Long
Dim Str As String
Public Function CheckVCD() As Boolean

hwndCurrent = FindWindow(vbNullString, vbNullString)
While hwndCurrent <> 0

sCaption = Space$(MAX_TEXT_SIZE + 1)
lRetVal = GetWindowText(hwndCurrent, sCaption, MAX_TEXT_SIZE)

If InStr(Trim$(sCaption), "VCDPLAYER") > 0 Then
CheckVCD = True
Exit Function
End If
hwndCurrent = GetWindow(hwndCurrent, 2) ' next
Wend

End Function
