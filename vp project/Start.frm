VERSION 5.00
Begin VB.Form StartForm 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Start.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Start.frx":000C
   ScaleHeight     =   1995
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00FFC0C0&
      Default         =   -1  'True
      Height          =   615
      Left            =   60
      Picture         =   "Start.frx":1EE6E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   930
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
End
Attribute VB_Name = "StartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim Serial As Long
Dim SerialStr As String

On Error Resume Next



ReadAgain:

If CheckVCD Then
   MsgBox "„‰ ›÷·ﬂ ﬁ„ »≈“«·… »—‰«„Ã Virtual CD ", vbMsgBoxRight + vbOKOnly + vbInformation, "«‰ »Â „‰ ›÷·ﬂ"
   End
End If

For X = 65 To 92
    
    d = GetDriveType(Chr(X) & ":")
        Serial_ = GetVolumeInformation(Chr(X) + ":\", SerialStr, 30, Serial, 30, 0, 30, 30)
      If d = 5 Then
         DriveLetter = Chr(X)
         Exit For
      End If

Next X

PathString = DriveLetter & ":"
'PathString = App.Path
'Exit Sub
Drive1.Drive = DriveLetter & ":"
If UCase(Drive1.Drive) = (DriveLetter & ":") Then
   Exit Sub
End If


For X = 1 To 28
    DriveLetter = Chr(Asc(DriveLetter) + 1)
    Drive1.Drive = DriveLetter
    If UCase(Drive1.Drive) = (DriveLetter & ":") Then
       Exit Sub
       Exit For
    End If
Next X

If MsgBox("„‰ ›÷·ﬂ ﬁ„ »≈œ—«Ã «·√”ÿÊ«‰… ", vbMsgBoxRight + vbOKCancel + vbInformation, "«‰ »Â „‰ ›÷·ﬂ") = vbCancel Then
   End
Else
   GoTo ReadAgain
End If

End Sub

Private Sub OkButton_Click()
On Error Resume Next

DriveLetter = Left(Drive1.Drive, 1)
Unload Me

ShellWindow.Show vbModal

End Sub
