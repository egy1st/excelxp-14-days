VERSION 5.00
Begin VB.Form StartForm 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "StartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim WinSys As String
Dim WinDir As String
Dim sINIFile As String
Dim sAviDriver As String




Dim x As Long
WinDir = Space(64)
x = GetWindowsDirectory(WinDir, 64)
WinDir = Trim(WinDir)
WinDir = Left(WinDir, Len(WinDir) - 1) & "\"

WinSys = WinDir & "SYSTEM\"


'sINIFile = WinDir & "system.ini"
'sAviDriver = sGetINI(sINIFile, "drivers32", "VIDC.TSCC", "?")

'If sAviDriver = "?" Then
'writeINI sINIFile, "drivers32", "VIDC.TSCC", "tsccvid.dll"
'End If


If Dir$(WinSys & "oct.73") = "" Then
Shell "Setup.exe"
End If


End Sub
