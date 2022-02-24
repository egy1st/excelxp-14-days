Attribute VB_Name = "Module1"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Public Function sGetINI(sINIFile As String, sSection As String, sKey As String, sDefault As String) As String
Dim sTemp As String * 256
Dim nLeenth As Integer

sTemp = Space$(256)
nlenth = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 255, sINIFile)
sGetINI = Left$(sTemp, nlenth)
End Function

Public Sub writeINI(sINIFile As String, sSection As String, sKey As String, sValue As String)

Dim n As Integer
Dim sTemp As String

sTemp = sValue
For n = 1 To Len(sValue)
If Mid$(sValue, n, 1) = vbCr Or Mid$(sValue, n, 1) = vbLf Then Mid$(sValue, n) = " "
Next n
n = WritePrivateProfileString(sSection, sKey, sTemp, sINIFile)

End Sub

