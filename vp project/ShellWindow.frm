VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form ShellWindow 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8970
   ClientLeft      =   15
   ClientTop       =   120
   ClientWidth     =   11970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "ShellWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "ShellWindow.frx":030A
   RightToLeft     =   -1  'True
   ScaleHeight     =   8970
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   705
      Left            =   510
      Picture         =   "ShellWindow.frx":15424C
      RightToLeft     =   -1  'True
      ScaleHeight     =   645
      ScaleWidth      =   10965
      TabIndex        =   21
      Top             =   420
      Width           =   11025
      Begin VB.CommandButton CloseButton 
         BackColor       =   &H0000FFFF&
         Caption         =   "€·ﬁ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10260
         MouseIcon       =   "ShellWindow.frx":16E526
         MousePointer    =   99  'Custom
         Picture         =   "ShellWindow.frx":16E968
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton MAAButton 
         BackColor       =   &H0000FFFF&
         Caption         =   "M.A.A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   0
         MouseIcon       =   "ShellWindow.frx":16EC72
         MousePointer    =   99  'Custom
         Picture         =   "ShellWindow.frx":16F0B4
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   -60
         Width           =   705
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9030
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ShellWindow.frx":16F3BE
      MousePointer    =   99  'Custom
      Picture         =   "ShellWindow.frx":16F6C8
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7710
      Width           =   2415
   End
   Begin VB.CommandButton HelpButton 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6600
      MaskColor       =   &H0080FFFF&
      MouseIcon       =   "ShellWindow.frx":175392
      MousePointer    =   99  'Custom
      Picture         =   "ShellWindow.frx":1757D4
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton sub8 
      Height          =   702
      Left            =   5880
      MouseIcon       =   "ShellWindow.frx":17DFCE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7740
      Width           =   3135
   End
   Begin VB.CommandButton sub7 
      Height          =   702
      Left            =   5880
      MouseIcon       =   "ShellWindow.frx":17E410
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7020
      Width           =   3135
   End
   Begin VB.CommandButton sub6 
      Height          =   702
      Left            =   5880
      MouseIcon       =   "ShellWindow.frx":17E71A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6300
      Width           =   3135
   End
   Begin VB.CommandButton sub5 
      Height          =   702
      Left            =   5880
      MouseIcon       =   "ShellWindow.frx":17EB5C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5580
      Width           =   3135
   End
   Begin VB.CommandButton sub4 
      Height          =   702
      Left            =   5880
      MouseIcon       =   "ShellWindow.frx":17EE66
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4860
      Width           =   3135
   End
   Begin VB.CommandButton sub3 
      Height          =   702
      Left            =   5880
      MouseIcon       =   "ShellWindow.frx":17F2A8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4140
      Width           =   3135
   End
   Begin VB.CommandButton sub2 
      Height          =   702
      Left            =   5880
      MouseIcon       =   "ShellWindow.frx":17F5B2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3420
      Width           =   3135
   End
   Begin VB.CommandButton sub1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   5880
      MouseIcon       =   "ShellWindow.frx":17F9F4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2700
      Width           =   3135
   End
   Begin VB.CommandButton day7 
      Height          =   735
      Left            =   9030
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ShellWindow.frx":17FCFE
      MousePointer    =   99  'Custom
      Picture         =   "ShellWindow.frx":180140
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6990
      Width           =   2415
   End
   Begin VB.CommandButton day6 
      Height          =   735
      Left            =   9030
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ShellWindow.frx":186582
      MousePointer    =   99  'Custom
      Picture         =   "ShellWindow.frx":1869C4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6270
      Width           =   2415
   End
   Begin VB.CommandButton day5 
      Height          =   735
      Left            =   9030
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ShellWindow.frx":18CE06
      MousePointer    =   99  'Custom
      Picture         =   "ShellWindow.frx":18D248
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5580
      Width           =   2415
   End
   Begin VB.CommandButton day4 
      Height          =   735
      Left            =   9030
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ShellWindow.frx":19368A
      MousePointer    =   99  'Custom
      Picture         =   "ShellWindow.frx":193ACC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4860
      Width           =   2415
   End
   Begin VB.CommandButton day3 
      Height          =   735
      Left            =   9030
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ShellWindow.frx":199F0E
      MousePointer    =   99  'Custom
      Picture         =   "ShellWindow.frx":19A350
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4140
      Width           =   2415
   End
   Begin VB.CommandButton day2 
      Height          =   735
      Left            =   9030
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ShellWindow.frx":1A0792
      MousePointer    =   99  'Custom
      Picture         =   "ShellWindow.frx":1A0BD4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3420
      Width           =   2415
   End
   Begin VB.CommandButton day1 
      Height          =   735
      Left            =   9030
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ShellWindow.frx":1A7016
      MousePointer    =   99  'Custom
      Picture         =   "ShellWindow.frx":1A7458
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2700
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Command2"
      Height          =   1395
      Left            =   6510
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1230
      Width           =   1905
   End
   Begin VB.CommandButton about 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9420
      MaskColor       =   &H0080FFFF&
      MouseIcon       =   "ShellWindow.frx":1AD89A
      MousePointer    =   99  'Custom
      Picture         =   "ShellWindow.frx":1ADCDC
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1350
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Command2"
      Height          =   1395
      Left            =   9330
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1260
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   8865
      Left            =   0
      Picture         =   "ShellWindow.frx":1B4CA6
      Stretch         =   -1  'True
      Top             =   90
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   8925
      Left            =   11610
      Picture         =   "ShellWindow.frx":236900
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   0
      Picture         =   "ShellWindow.frx":2B855A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11745
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   240
      Picture         =   "ShellWindow.frx":339FB0
      Stretch         =   -1  'True
      Top             =   8610
      Width           =   11745
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   7215
      Left            =   510
      TabIndex        =   15
      Top             =   1230
      Width           =   5325
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   0   'False
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   0
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "ShellWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Flag As Integer
Private Sub about_Click()
On Error Resume Next

 PlayerFrm.Player1.FileName = PathString & "\avi\0.maa"
 PlayerFrm.Show vbModal

End Sub
Private Sub Command1_Click()
ScriptForm.Show vbModal
End Sub

Private Sub CloseButton_Click()
On Error Resume Next
Unload Me
End
End Sub
Private Sub HelpButton_Click()
On Error Resume Next
HelpFrm.Show vbModal
End Sub
Private Sub day1_Click()
On Error Resume Next

Flag = 1

 day1.Picture = LoadPicture(PathString & "\bmp\day11.bmp")
 day2.Picture = LoadPicture(PathString & "\bmp\day2.bmp")
 day3.Picture = LoadPicture(PathString & "\bmp\day3.bmp")
 day4.Picture = LoadPicture(PathString & "\bmp\day4.bmp")
 day5.Picture = LoadPicture(PathString & "\bmp\day5.bmp")
 day6.Picture = LoadPicture(PathString & "\bmp\day6.bmp")
 day7.Picture = LoadPicture(PathString & "\bmp\day7.bmp")
 
 ShowAllButtons

 sub3.Visible = False
 sub4.Visible = False
 sub5.Visible = False
 sub6.Visible = False
 sub7.Visible = False
 sub8.Visible = False
  
sub1.Picture = LoadPicture(PathString & "\bmp\b1.bmp")
sub2.Picture = LoadPicture(PathString & "\bmp\b2.bmp")

  
End Sub

Private Sub day2_Click()
On Error Resume Next
 
 Flag = 2
 
 day1.Picture = LoadPicture(PathString & "\bmp\day1.bmp")
 day2.Picture = LoadPicture(PathString & "\bmp\day22.bmp")
 day3.Picture = LoadPicture(PathString & "\bmp\day3.bmp")
 day4.Picture = LoadPicture(PathString & "\bmp\day4.bmp")
 day5.Picture = LoadPicture(PathString & "\bmp\day5.bmp")
 day6.Picture = LoadPicture(PathString & "\bmp\day6.bmp")
 day7.Picture = LoadPicture(PathString & "\bmp\day7.bmp")


 ShowAllButtons
 
 sub1.Visible = False
 sub4.Visible = False
 sub5.Visible = False
 sub6.Visible = False
 sub7.Visible = False
 sub8.Visible = False
 
 sub2.Picture = LoadPicture(PathString & "\bmp\b3.bmp")
 sub3.Picture = LoadPicture(PathString & "\bmp\b4.bmp")


End Sub

Private Sub day3_Click()
On Error Resume Next
 
 Flag = 3
 
 day1.Picture = LoadPicture(PathString & "\bmp\day1.bmp")
 day2.Picture = LoadPicture(PathString & "\bmp\day2.bmp")
 day3.Picture = LoadPicture(PathString & "\bmp\day33.bmp")
 day4.Picture = LoadPicture(PathString & "\bmp\day4.bmp")
 day5.Picture = LoadPicture(PathString & "\bmp\day5.bmp")
 day6.Picture = LoadPicture(PathString & "\bmp\day6.bmp")
 day7.Picture = LoadPicture(PathString & "\bmp\day7.bmp")

 ShowAllButtons
 
 sub1.Visible = False
 sub2.Visible = False
 sub5.Visible = False
 sub6.Visible = False
 sub7.Visible = False
 sub8.Visible = False
 
 sub3.Picture = LoadPicture(PathString & "\bmp\b5.bmp")
 sub4.Picture = LoadPicture(PathString & "\bmp\b6.bmp")


End Sub

Private Sub day4_Click()
On Error Resume Next
 
Flag = 4
 
 day1.Picture = LoadPicture(PathString & "\bmp\day1.bmp")
 day2.Picture = LoadPicture(PathString & "\bmp\day2.bmp")
 day3.Picture = LoadPicture(PathString & "\bmp\day3.bmp")
 day4.Picture = LoadPicture(PathString & "\bmp\day44.bmp")
 day5.Picture = LoadPicture(PathString & "\bmp\day5.bmp")
 day6.Picture = LoadPicture(PathString & "\bmp\day6.bmp")
 day7.Picture = LoadPicture(PathString & "\bmp\day7.bmp")

 ShowAllButtons
 
 sub1.Visible = False
 sub2.Visible = False
 sub3.Visible = False
 sub6.Visible = False
 sub7.Visible = False
 sub8.Visible = False
 
 sub4.Picture = LoadPicture(PathString & "\bmp\b7.bmp")
 sub5.Picture = LoadPicture(PathString & "\bmp\b8.bmp")

End Sub

Private Sub day5_Click()
On Error Resume Next
 
 Flag = 5
 
 day1.Picture = LoadPicture(PathString & "\bmp\day1.bmp")
 day2.Picture = LoadPicture(PathString & "\bmp\day2.bmp")
 day3.Picture = LoadPicture(PathString & "\bmp\day3.bmp")
 day4.Picture = LoadPicture(PathString & "\bmp\day4.bmp")
 day5.Picture = LoadPicture(PathString & "\bmp\day55.bmp")
 day6.Picture = LoadPicture(PathString & "\bmp\day6.bmp")
 day7.Picture = LoadPicture(PathString & "\bmp\day7.bmp")

 ShowAllButtons
 
 sub1.Visible = False
 sub2.Visible = False
 sub3.Visible = False
 sub4.Visible = False
 sub7.Visible = False
 sub8.Visible = False
 
 sub5.Picture = LoadPicture(PathString & "\bmp\b9.bmp")
 sub6.Picture = LoadPicture(PathString & "\bmp\b10.bmp")

End Sub

Private Sub day6_Click()
On Error Resume Next
 
 Flag = 6
 
 day1.Picture = LoadPicture(PathString & "\bmp\day1.bmp")
 day2.Picture = LoadPicture(PathString & "\bmp\day2.bmp")
 day3.Picture = LoadPicture(PathString & "\bmp\day3.bmp")
 day4.Picture = LoadPicture(PathString & "\bmp\day4.bmp")
 day5.Picture = LoadPicture(PathString & "\bmp\day5.bmp")
 day6.Picture = LoadPicture(PathString & "\bmp\day66.bmp")
 day7.Picture = LoadPicture(PathString & "\bmp\day7.bmp")

 ShowAllButtons
 
 sub1.Visible = False
 sub2.Visible = False
 sub3.Visible = False
 sub4.Visible = False
 sub5.Visible = False
 sub8.Visible = False
  
 sub6.Picture = LoadPicture(PathString & "\bmp\b111.bmp")
 sub7.Picture = LoadPicture(PathString & "\bmp\b12.bmp")

End Sub

Private Sub day7_Click()
On Error Resume Next
 
 Flag = 7
 
 day1.Picture = LoadPicture(PathString & "\bmp\day1.bmp")
 day2.Picture = LoadPicture(PathString & "\bmp\day2.bmp")
 day3.Picture = LoadPicture(PathString & "\bmp\day3.bmp")
 day4.Picture = LoadPicture(PathString & "\bmp\day4.bmp")
 day5.Picture = LoadPicture(PathString & "\bmp\day5.bmp")
 day6.Picture = LoadPicture(PathString & "\bmp\day6.bmp")
 day7.Picture = LoadPicture(PathString & "\bmp\day77.bmp")

 ShowAllButtons
 
 sub1.Visible = False
 sub2.Visible = False
 sub3.Visible = False
 sub4.Visible = False
 sub5.Visible = False
 sub6.Visible = False
 
 
 sub7.Picture = LoadPicture(PathString & "\bmp\b13.bmp")
 sub8.Picture = LoadPicture(PathString & "\bmp\b14.bmp")
 
End Sub
Public Sub ShowAllButtons()
On Error Resume Next
sub1.Visible = True
sub2.Visible = True
sub3.Visible = True
sub4.Visible = True
sub5.Visible = True
sub6.Visible = True
sub7.Visible = True
sub8.Visible = True

End Sub
Public Sub HideAllButtons()
On Error Resume Next
 sub1.Visible = False
 sub2.Visible = False
 sub3.Visible = False
 sub4.Visible = False
 sub5.Visible = False
 sub6.Visible = False
 sub7.Visible = False
 sub8.Visible = False
 
 
End Sub

Private Sub Form_Load()
Dim Date1 As Date
Dim X As Integer
Dim Filenum As Integer
Dim FileTime_ As Long
Dim FileDate1 As FILETIME
Dim FileDate2 As FILETIME
Dim FileDate3 As FILETIME
Dim of As OFSTRUCT
Dim handle As Long
Dim FileSize As Long

On Error Resume Next

'''''''''''''''''''''''''''''''''''''''''''''


MediaPlayer1.FileName = PathString & "\Demo.avi"

X = GetSystemMetrics(0)
y = GetSystemMetrics(1)


'''''''''''''''''''''''''''''''''''''''''''''''

HideAllButtons

handle = OpenFile(App.Path & "\maa.dll", of, 0)
FileTime_ = GetFileTime(handle, FileDate1, FileDate2, FileDate3)

If FileDate1.dwHighDateTime <> FileDate3.dwHighDateTime Then
x_ = MsgBox("Â–« «·»—‰«„Ã Ì „  ‘€Ì·Â „‰ ‰”Œ… €Ì— √’·Ì…", " Õ–Ì—")
Unload Me
Exit Sub
End If

If App.PrevInstance <> 0 Then
x_ = MsgBox("Â–« «·»—‰«„Ã Ì „ Õ«·Ì«  ‘€Ì· ‰”Œ… √Œ—Ï „‰Â", " Õ–Ì—")
Unload Me
End If

End Sub

Private Sub MAAButton_Click()
On Error Resume Next
HelpFrm.Show vbModal
End Sub

Private Sub sub1_Click()
On Error Resume Next
If Flag = 1 Then
 sub1.Picture = LoadPicture(PathString & "\bmp\b11.bmp")
 sub2.Picture = LoadPicture(PathString & "\bmp\b2.bmp")

 PlayerFrm.Player1.FileName = PathString & "\avi\1.maa"
 PlayerFrm.Show vbModal
 End If
 
End Sub

Private Sub sub10_Click()
On Error Resume Next
If Flag = 6 Then
 sub6.Picture = LoadPicture(PathString & "\bmp\sub66b.bmp")
 sub7.Picture = LoadPicture(PathString & "\bmp\sub67.bmp")
 
 ElseIf Flag = 7 Then
 sub7.Picture = LoadPicture(PathString & "\bmp\sub77b.bmp")
 sub8.Picture = LoadPicture(PathString & "\bmp\sub78.bmp")
 
End If
End Sub

Private Sub sub2_Click()
On Error Resume Next
If Flag = 1 Then
 sub1.Picture = LoadPicture(PathString & "\bmp\b1.bmp")
 sub2.Picture = LoadPicture(PathString & "\bmp\b22.bmp")
 
 PlayerFrm.Player1.FileName = PathString & "\avi\2.maa"
 PlayerFrm.Show vbModal
ElseIf Flag = 2 Then
 sub2.Picture = LoadPicture(PathString & "\bmp\b33.bmp")
 sub3.Picture = LoadPicture(PathString & "\bmp\b4.bmp")
 PlayerFrm.Player1.FileName = PathString & "\avi\3.maa"
 PlayerFrm.Show vbModal

End If


End Sub

Private Sub sub3_Click()
On Error Resume Next

If Flag = 2 Then
 sub2.Picture = LoadPicture(PathString & "\bmp\b3.bmp")
 sub3.Picture = LoadPicture(PathString & "\bmp\b44.bmp")

 PlayerFrm.Player1.FileName = PathString & "\avi\4.maa"
 PlayerFrm.Show vbModal

ElseIf Flag = 3 Then
 sub3.Picture = LoadPicture(PathString & "\bmp\b55.bmp")
 sub4.Picture = LoadPicture(PathString & "\bmp\b6.bmp")

 PlayerFrm.Player1.FileName = PathString & "\avi\5.maa"
 PlayerFrm.Show vbModal

End If


End Sub
Private Sub sub4_Click()
On Error Resume Next
If Flag = 3 Then
 sub3.Picture = LoadPicture(PathString & "\bmp\b5.bmp")
 sub4.Picture = LoadPicture(PathString & "\bmp\b66.bmp")

 PlayerFrm.Player1.FileName = PathString & "\avi\6.maa"
 PlayerFrm.Show vbModal

ElseIf Flag = 4 Then
 sub4.Picture = LoadPicture(PathString & "\bmp\b77.bmp")
 sub5.Picture = LoadPicture(PathString & "\bmp\b8.bmp")

 PlayerFrm.Player1.FileName = PathString & "\avi\7.maa"
 PlayerFrm.Show vbModal

End If

End Sub

Private Sub sub5_Click()
On Error Resume Next
If Flag = 4 Then
 sub4.Picture = LoadPicture(PathString & "\bmp\b7.bmp")
 sub5.Picture = LoadPicture(PathString & "\bmp\b88.bmp")

 PlayerFrm.Player1.FileName = PathString & "\avi\8.maa"
 PlayerFrm.Show vbModal

ElseIf Flag = 5 Then
 sub5.Picture = LoadPicture(PathString & "\bmp\b99.bmp")
 sub6.Picture = LoadPicture(PathString & "\bmp\b10.bmp")

 PlayerFrm.Player1.FileName = PathString & "\avi\9.maa"
 PlayerFrm.Show vbModal

End If
End Sub

Private Sub sub6_Click()
On Error Resume Next
If Flag = 5 Then
 sub5.Picture = LoadPicture(PathString & "\bmp\b9.bmp")
 sub6.Picture = LoadPicture(PathString & "\bmp\b1010.bmp")

 PlayerFrm.Player1.FileName = PathString & "\avi\10.maa"
 PlayerFrm.Show vbModal

ElseIf Flag = 6 Then
 sub6.Picture = LoadPicture(PathString & "\bmp\b1111.bmp")
 sub7.Picture = LoadPicture(PathString & "\bmp\b12.bmp")
 
 PlayerFrm.Player1.FileName = PathString & "\avi\11.maa"
 PlayerFrm.Show vbModal

End If

End Sub
Private Sub sub7_Click()
On Error Resume Next
If Flag = 6 Then
 sub6.Picture = LoadPicture(PathString & "\bmp\b111.bmp")
 sub7.Picture = LoadPicture(PathString & "\bmp\b1212.bmp")
 
 PlayerFrm.Player1.FileName = PathString & "\avi\12.maa"
 PlayerFrm.Show vbModal

ElseIf Flag = 7 Then
 sub7.Picture = LoadPicture(PathString & "\bmp\b1313.bmp")
 sub8.Picture = LoadPicture(PathString & "\bmp\b14.bmp")

 PlayerFrm.Player1.FileName = PathString & "\avi\13.maa"
 PlayerFrm.Show vbModal

End If
End Sub
Private Sub sub8_Click()
On Error Resume Next

 sub7.Picture = LoadPicture(PathString & "\bmp\b13.bmp")
 sub8.Picture = LoadPicture(PathString & "\bmp\b1414.bmp")

 PlayerFrm.Player1.FileName = PathString & "\avi\14.maa"
 PlayerFrm.Show vbModal

End Sub

