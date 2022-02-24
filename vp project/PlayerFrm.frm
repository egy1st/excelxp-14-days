VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form PlayerFrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8970
   ClientLeft      =   -4425
   ClientTop       =   -1995
   ClientWidth     =   11970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   11970
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton ExitButton 
      BackColor       =   &H8000000D&
      Cancel          =   -1  'True
      Height          =   375
      Left            =   240
      MouseIcon       =   "PlayerFrm.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "PlayerFrm.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8760
      UseMaskColor    =   -1  'True
      Width           =   1275
   End
   Begin VB.CommandButton PauseButton 
      DownPicture     =   "PlayerFrm.frx":0664
      Height          =   255
      Left            =   9120
      MouseIcon       =   "PlayerFrm.frx":1A46
      MousePointer    =   99  'Custom
      Picture         =   "PlayerFrm.frx":1D50
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8760
      UseMaskColor    =   -1  'True
      Width           =   1275
   End
   Begin VB.CommandButton PlayButton 
      DownPicture     =   "PlayerFrm.frx":3132
      Height          =   255
      Left            =   10440
      MouseIcon       =   "PlayerFrm.frx":4514
      MousePointer    =   99  'Custom
      Picture         =   "PlayerFrm.frx":481E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8760
      UseMaskColor    =   -1  'True
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Text            =   "                              ··Œ—ÊÃ „‰ Â–Â «·‘«‘… «÷€ÿ «·“— Stop ›Ï «ﬁ’Ï Ì”«— «·‘«‘… √Ê «÷€ÿ Escape"
      Top             =   8760
      Width           =   12045
   End
   Begin MediaPlayerCtl.MediaPlayer Player1 
      Height          =   9345
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   -120
      Width           =   12015
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   0   'False
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
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   -1  'True
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
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
      ShowControls    =   -1  'True
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "PlayerFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ExitButton_Click()
On Error Resume Next
Unload Me
End Sub
Private Sub PauseButton_Click()
On Error Resume Next
Player1.Pause
End Sub
Private Sub PlayButton_Click()
On Error Resume Next
Player1.Play
End Sub
