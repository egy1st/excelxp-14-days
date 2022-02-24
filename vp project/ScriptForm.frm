VERSION 5.00
Begin VB.Form ScriptForm 
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11145
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "ScriptForm.frx":0000
   ScaleHeight     =   8175
   ScaleWidth      =   11145
   StartUpPosition =   1  'CenterOwner
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
      Left            =   10440
      MouseIcon       =   "ScriptForm.frx":153F42
      MousePointer    =   99  'Custom
      Picture         =   "ScriptForm.frx":154384
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   0
      Width           =   705
   End
   Begin VB.CommandButton NextButton 
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
      Left            =   9480
      MaskColor       =   &H0080FFFF&
      MouseIcon       =   "ScriptForm.frx":1547C6
      MousePointer    =   99  'Custom
      Picture         =   "ScriptForm.frx":154C08
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6780
      Width           =   1485
   End
   Begin VB.CommandButton section14 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   7500
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ScriptForm.frx":15AFBA
      MousePointer    =   99  'Custom
      Picture         =   "ScriptForm.frx":15B3FC
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5850
      Width           =   1800
   End
   Begin VB.CommandButton section13 
      Height          =   735
      Left            =   7500
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ScriptForm.frx":15FABE
      MousePointer    =   99  'Custom
      Picture         =   "ScriptForm.frx":15FF00
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5130
      Width           =   1800
   End
   Begin VB.CommandButton section12 
      Height          =   735
      Left            =   7500
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ScriptForm.frx":1645C2
      MousePointer    =   99  'Custom
      Picture         =   "ScriptForm.frx":164A04
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   1800
   End
   Begin VB.CommandButton section11 
      Height          =   735
      Left            =   7500
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ScriptForm.frx":1690C6
      MousePointer    =   99  'Custom
      Picture         =   "ScriptForm.frx":169508
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   1800
   End
   Begin VB.CommandButton section10 
      Height          =   735
      Left            =   7500
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ScriptForm.frx":16DBCA
      MousePointer    =   99  'Custom
      Picture         =   "ScriptForm.frx":16E00C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Width           =   1800
   End
   Begin VB.CommandButton section9 
      Height          =   735
      Left            =   7500
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ScriptForm.frx":1726CE
      MousePointer    =   99  'Custom
      Picture         =   "ScriptForm.frx":172B10
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2310
      Width           =   1800
   End
   Begin VB.CommandButton section8 
      Height          =   735
      Left            =   7500
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ScriptForm.frx":176EE2
      MousePointer    =   99  'Custom
      Picture         =   "ScriptForm.frx":177324
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   1800
   End
   Begin VB.CommandButton section1 
      Height          =   735
      Left            =   9300
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ScriptForm.frx":17B6F6
      MousePointer    =   99  'Custom
      Picture         =   "ScriptForm.frx":17BB38
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   1800
   End
   Begin VB.CommandButton section2 
      Height          =   735
      Left            =   9300
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ScriptForm.frx":17FF0A
      MousePointer    =   99  'Custom
      Picture         =   "ScriptForm.frx":18034C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1800
   End
   Begin VB.CommandButton section3 
      Height          =   735
      Left            =   9300
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ScriptForm.frx":18471E
      MousePointer    =   99  'Custom
      Picture         =   "ScriptForm.frx":184B60
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1800
   End
   Begin VB.CommandButton section4 
      Height          =   735
      Left            =   9300
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ScriptForm.frx":188F32
      MousePointer    =   99  'Custom
      Picture         =   "ScriptForm.frx":189374
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1800
   End
   Begin VB.CommandButton section5 
      Height          =   735
      Left            =   9300
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ScriptForm.frx":18D746
      MousePointer    =   99  'Custom
      Picture         =   "ScriptForm.frx":18DB88
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   1800
   End
   Begin VB.CommandButton section6 
      Height          =   735
      Left            =   9300
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ScriptForm.frx":191F5A
      MousePointer    =   99  'Custom
      Picture         =   "ScriptForm.frx":19239C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5130
      Width           =   1800
   End
   Begin VB.CommandButton section7 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   9300
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "ScriptForm.frx":19676E
      MousePointer    =   99  'Custom
      Picture         =   "ScriptForm.frx":196BB0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5850
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Command2"
      Height          =   1395
      Left            =   9360
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6690
      Width           =   1695
   End
   Begin VB.CommandButton FirstButton 
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
      Left            =   8490
      MaskColor       =   &H0080FFFF&
      MouseIcon       =   "ScriptForm.frx":19AF82
      MousePointer    =   99  'Custom
      Picture         =   "ScriptForm.frx":19B3C4
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   180
      Width           =   1515
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Command2"
      Height          =   1395
      Left            =   8400
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   90
      Width           =   1695
   End
   Begin VB.CommandButton PreviousButton 
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
      Left            =   7680
      MaskColor       =   &H0080FFFF&
      MouseIcon       =   "ScriptForm.frx":1A238E
      MousePointer    =   99  'Custom
      Picture         =   "ScriptForm.frx":1A27D0
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6780
      Width           =   1515
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FFFF&
      Caption         =   "Command2"
      Height          =   1395
      Left            =   7590
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6690
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Â–Â «·‰”Œ…    €Ì—          „—Œ’…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   7620
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   120
      Width           =   735
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   10380
      Left            =   30
      Picture         =   "ScriptForm.frx":1A8B82
      Top             =   -120
      Width           =   7380
   End
End
Attribute VB_Name = "ScriptForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pic As Integer

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub CloseButton_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub FirstButton_Click()

On Error Resume Next

pic = 1
Board.Picture = LoadPicture(PathString & "/bmp/1.bmp")
End Sub

Private Sub Form_Load()
pic = 1
End Sub

Private Sub NextButton_Click()

On Error Resume Next

pic = pic + 1
Board.Picture = LoadPicture(PathString & "/bmp/" & CStr(pic) & ".bmp")
End Sub

Private Sub PreviousButton_Click()

On Error Resume Next

pic = pic - 1
Board.Picture = LoadPicture(PathString & "/bmp/" & CStr(pic) & ".bmp")
End Sub

Private Sub section1_Click()

On Error Resume Next

pic = 2
Board.Picture = LoadPicture(PathString & "/bmp/2.bmp")
End Sub
Private Sub section10_Click()

On Error Resume Next

pic = 35
Board.Picture = LoadPicture(PathString & "/bmp/35.bmp")
End Sub
Private Sub section11_Click()

On Error Resume Next

pic = 38
Board.Picture = LoadPicture(PathString & "/bmp/38.bmp")
End Sub

Private Sub section12_Click()

On Error Resume Next

pic = 41
Board.Picture = LoadPicture(PathString & "/bmp/41.bmp")
End Sub

Private Sub section13_Click()

On Error Resume Next

pic = 45
Board.Picture = LoadPicture(PathString & "/bmp/45.bmp")
End Sub

Private Sub section14_Click()

On Error Resume Next

pic = 48
Board.Picture = LoadPicture(PathString & "/bmp/48.bmp")
End Sub

Private Sub section2_Click()

On Error Resume Next

pic = 6
Board.Picture = LoadPicture(PathString & "/bmp/6.bmp")
End Sub

Private Sub section3_Click()

On Error Resume Next

pic = 10
Board.Picture = LoadPicture(PathString & "/bmp/10.bmp")
End Sub

Private Sub section4_Click()

On Error Resume Next

pic = 14
Board.Picture = LoadPicture(PathString & "/bmp/14.bmp")
End Sub

Private Sub section5_Click()

On Error Resume Next

pic = 17
Board.Picture = LoadPicture(PathString & "/bmp/17.bmp")
End Sub

Private Sub section6_Click()

On Error Resume Next

pic = 20
Board.Picture = LoadPicture(PathString & "/bmp/20.bmp")
End Sub

Private Sub section7_Click()

On Error Resume Next

pic = 24
Board.Picture = LoadPicture(PathString & "/bmp/24.bmp")
End Sub

Private Sub section8_Click()

On Error Resume Next

pic = 28
Board.Picture = LoadPicture(PathString & "/bmp/28.bmp")
End Sub

Private Sub section9_Click()

On Error Resume Next

pic = 32
Board.Picture = LoadPicture(PathString & "/bmp/32.bmp")
End Sub
