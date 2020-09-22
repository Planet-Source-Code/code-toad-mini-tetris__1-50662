VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mini Tetris"
   ClientHeight    =   3435
   ClientLeft      =   1980
   ClientTop       =   525
   ClientWidth     =   3390
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   3390
   Begin VB.PictureBox picHigh 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1630
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   240
      ScaleWidth      =   1035
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   3135
      Width           =   1035
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   3225
      Left            =   20
      TabIndex        =   39
      Top             =   3480
      Visible         =   0   'False
      Width           =   3450
      Begin VB.TextBox txtKeyMap 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   4
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "<F>"
         Top             =   2760
         Width           =   1275
      End
      Begin VB.TextBox txtKeyMap 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   3
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "<D>"
         Top             =   2400
         Width           =   1275
      End
      Begin VB.TextBox txtKeyMap 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "<Down Arrow>"
         Top             =   1980
         Width           =   1275
      End
      Begin VB.TextBox txtKeyMap 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "<Right Arrow>"
         Top             =   1620
         Width           =   1275
      End
      Begin VB.TextBox txtKeyMap 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "<Left Arrow>"
         Top             =   1260
         Width           =   1275
      End
      Begin VB.CheckBox chkMusic 
         Caption         =   "Music"
         Height          =   195
         Left            =   600
         TabIndex        =   5
         Top             =   720
         Width           =   1035
      End
      Begin VB.CheckBox chkSFX 
         Caption         =   "Sound Effects"
         Height          =   195
         Left            =   600
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         ItemData        =   "frmMain.frx":0AEC
         Left            =   600
         List            =   "frmMain.frx":0AF6
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   1095
      End
      Begin VB.ComboBox cboLevel 
         Height          =   315
         ItemData        =   "frmMain.frx":0B0A
         Left            =   2280
         List            =   "frmMain.frx":0B2C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   675
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Counterclockwise Rot.:"
         Height          =   255
         Index           =   10
         Left            =   60
         TabIndex        =   46
         Top             =   2820
         Width           =   1755
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Clockwise Rotation:"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   45
         Top             =   2460
         Width           =   1575
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Drop Faster:"
         Height          =   255
         Index           =   8
         Left            =   900
         TabIndex        =   44
         Top             =   2040
         Width           =   915
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Move Right:"
         Height          =   255
         Index           =   7
         Left            =   900
         TabIndex        =   43
         Top             =   1680
         Width           =   915
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Move Left:"
         Height          =   255
         Index           =   6
         Left            =   1020
         TabIndex        =   42
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Game:"
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   41
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Level:"
         Height          =   255
         Index           =   4
         Left            =   1740
         TabIndex        =   40
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   870
      Picture         =   "frmMain.frx":0B4E
      ScaleHeight     =   240
      ScaleWidth      =   675
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3135
      Width           =   670
   End
   Begin VB.PictureBox picStart 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   60
      Picture         =   "frmMain.frx":112E
      ScaleHeight     =   240
      ScaleWidth      =   705
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3140
      Width           =   700
   End
   Begin VB.PictureBox picGoodJob 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   3105
      Left            =   5460
      Picture         =   "frmMain.frx":164E
      ScaleHeight     =   203
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   38
      Top             =   900
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picGameOver 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   3105
      Left            =   3840
      Picture         =   "frmMain.frx":3DE0
      ScaleHeight     =   203
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   37
      Top             =   900
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picLevel 
      Appearance      =   0  'Flat
      BackColor       =   &H00B4B4B4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   260
      Left            =   2430
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   795
      Width           =   300
   End
   Begin VB.Timer tmrKeyInput 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   2400
      Top             =   1920
   End
   Begin VB.PictureBox picLines 
      Appearance      =   0  'Flat
      BackColor       =   &H00B4B4B4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   260
      Left            =   2430
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1575
      Width           =   450
   End
   Begin VB.PictureBox picScore 
      Appearance      =   0  'Flat
      BackColor       =   &H00B4B4B4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   260
      Left            =   2430
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1200
      Width           =   900
   End
   Begin VB.PictureBox picNum 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   320
      Index           =   10
      Left            =   5640
      Picture         =   "frmMain.frx":64E0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   29
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picNum 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   320
      Index           =   9
      Left            =   5460
      Picture         =   "frmMain.frx":676B
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   28
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picNum 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   320
      Index           =   8
      Left            =   5280
      Picture         =   "frmMain.frx":6B42
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   27
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picNum 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   320
      Index           =   7
      Left            =   5100
      Picture         =   "frmMain.frx":6F22
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   26
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picNum 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   320
      Index           =   6
      Left            =   4920
      Picture         =   "frmMain.frx":72B6
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   25
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picNum 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   320
      Index           =   5
      Left            =   4740
      Picture         =   "frmMain.frx":768C
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   24
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picNum 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   320
      Index           =   4
      Left            =   4560
      Picture         =   "frmMain.frx":7A5E
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   23
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picNum 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   320
      Index           =   3
      Left            =   4380
      Picture         =   "frmMain.frx":7E05
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   22
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picNum 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   320
      Index           =   2
      Left            =   4200
      Picture         =   "frmMain.frx":81A0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   21
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picNum 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   320
      Index           =   1
      Left            =   4020
      Picture         =   "frmMain.frx":8582
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   20
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picNum 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   320
      Index           =   0
      Left            =   3840
      Picture         =   "frmMain.frx":88D6
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   19
      Top             =   540
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picErase 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   210
      Left            =   3840
      Picture         =   "frmMain.frx":8CB3
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   18
      Top             =   300
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.PictureBox picNext 
      BackColor       =   &H00B4B4B4&
      BorderStyle     =   0  'None
      Height          =   603
      Left            =   2430
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   60
      Width           =   610
   End
   Begin VB.PictureBox picBoard 
      Appearance      =   0  'Flat
      BackColor       =   &H00B4B4B4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   60
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   60
      Width           =   1500
   End
   Begin VB.PictureBox picBG 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   210
      Left            =   3840
      Picture         =   "frmMain.frx":9938
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   16
      Top             =   60
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picBlock 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   210
      Index           =   0
      Left            =   4080
      Picture         =   "frmMain.frx":9C0C
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   15
      Top             =   60
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picBlock 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   210
      Index           =   1
      Left            =   4320
      Picture         =   "frmMain.frx":9F67
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   14
      Top             =   60
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picBlock 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   210
      Index           =   2
      Left            =   4560
      Picture         =   "frmMain.frx":A2D2
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   13
      Top             =   60
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picBlock 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   210
      Index           =   3
      Left            =   4800
      Picture         =   "frmMain.frx":A657
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   12
      Top             =   60
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H80000003&
      FillColor       =   &H00C0C0C0&
      Height          =   270
      Left            =   1620
      Top             =   3120
      Width           =   1070
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000003&
      FillColor       =   &H00C0C0C0&
      Height          =   270
      Left            =   840
      Top             =   3120
      Width           =   740
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000003&
      FillColor       =   &H00C0C0C0&
      Height          =   270
      Left            =   30
      Top             =   3120
      Width           =   765
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000003&
      FillColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2415
      Top             =   780
      Width           =   330
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   1740
      TabIndex        =   36
      Top             =   795
      Width           =   615
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   1740
      TabIndex        =   34
      Top             =   60
      Width           =   615
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1740
      TabIndex        =   33
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Lines"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1740
      TabIndex        =   32
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000003&
      FillColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2415
      Top             =   1560
      Width           =   480
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000003&
      FillColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2415
      Top             =   1185
      Width           =   930
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000003&
      FillColor       =   &H00C0C0C0&
      Height          =   3050
      Left            =   30
      Top             =   40
      Width           =   1550
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000003&
      FillColor       =   &H00C0C0C0&
      Height          =   630
      Left            =   2415
      Top             =   45
      Width           =   645
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuFileNewGame 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "Options"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboLevel_Click()
gReg.SaveSetting "startup", INI_START_LEVEL, cboLevel.Text 'remember which level (0 - 9) to start on
End Sub

Private Sub cboType_Click()
gReg.SaveSetting "startup", INI_GAME_TYPE, cboType.Text 'remember which game type to start with
End Sub

Private Sub chkMusic_Click()
gReg.SaveSetting "startup", INI_USE_MUSIC, chkMusic.Value 'Play Music?
End Sub

Private Sub chkSFX_Click()
gReg.SaveSetting "startup", INI_USE_SFX, chkSFX.Value 'Play Sound Effects?
End Sub

Private Sub Command1_Click()
Load frmHighScores
frmHighScores.Show 1
End Sub



Private Sub Form_KeyDown(keyCode As Integer, Shift As Integer)
'•••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'Keep a 'map' of keys that are down or were pressed since last time we checked
'•••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Select Case keyCode
  Case game.keyCodeRotateCounterclockwise
    game.keyDownRotateLeft = True
    game.keyWasPressedRotateLeft = True
  Case game.keyCodeRotateClockwise
    game.keyDownRotateRight = True
    game.keyWasPressedRotateRight = True
  Case game.keyCodeMoveLeft
    game.keyDownLeft = True
    game.keyWasPressedLeft = True
  Case game.keyCodeMoveRight
    game.keyDownRight = True
    game.keyWasPressedRight = True
  Case game.keyCodeDrop
    game.keyDownDown = True
    game.keyWasPressedDown = True
End Select
End Sub

Private Sub Form_KeyUp(keyCode As Integer, Shift As Integer)
'•••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'If a key is no longer pressed, remove it from the map
'•••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Select Case keyCode
  Case game.keyCodeRotateCounterclockwise ' 68 'D
    game.keyDownRotateLeft = False
  Case game.keyCodeRotateClockwise '70 'F
    game.keyDownRotateRight = False
  Case game.keyCodeMoveLeft '37
    game.keyDownLeft = False
  Case game.keyCodeMoveRight '39
    game.keyDownRight = False
  Case game.keyCodeDrop ' 40
    game.keyDownDown = False
End Select
'Movecount stores the number of piece moves that have occured while a button is down
'Reset it to zero when any key is released
game.moveCount = 0
End Sub

Private Sub Form_Load()
Dim intCounter As Integer

Randomize
game.rowCount = 20
game.colCount = 10
game.board = picBoard
game.tile = picBG
game.picGoodJob = picGoodJob   'Image to display when user did a good job
game.picGameOver = picGameOver 'Image to display when game is over
game.picNextPiece = Me.picNext 'Picture box that will display the next game piece
game.picErase = Me.picErase    'Image to display when a line is erased
game.piece.initialize          'Load all possible shapes and rotation states
game.piece.loadRandomShape     'Choose a random shape
game.piece.y = 0
game.piece.x = 4
game.nextPiece.initialize      'This piece will be the next piece that is played after the current piece
game.nextPiece.loadRandomShape 'Choose a random shape

'The 'blocks' are the images that will be used to display different colors for the shapes that drop
For intCounter = 0 To 3
  game.addBlock picBlock(intCounter)
Next intCounter

gStrINIFile = App.Path & "\tet.ini" 'set up the path to the INI file

'Initialize each of the score objects by telling them which pictures to get each number from
For intCounter = 0 To 10
  score.addNumber picNum(intCounter)
  scoreLines.addNumber picNum(intCounter)
  scoreLevel.addNumber picNum(intCounter)
Next intCounter
'Let each score object know where to display it's score
score.score = picScore
scoreLines.score = picLines
scoreLevel.score = picLevel

'••••••••••••••••••••••••••••••••••••••••••••••••v
'Load info from INI file
'••••••••••••••••••••••••••••••••••••••••••••••••v
With gReg
  cboType.Text = .GetSetting("startup", INI_GAME_TYPE, "Type A")
  cboLevel.Text = .GetSetting("startup", INI_START_LEVEL, "0")
  chkMusic.Value = .GetSetting("startup", INI_USE_MUSIC, "0")
  chkSFX.Value = .GetSetting("startup", INI_USE_SFX, "0")
  txtKeyMap(0).Text = showKeyMap(.GetSetting("startup", INI_KEYCODE_MOVE_LEFT, vbKeyLeft))
  txtKeyMap(1).Text = showKeyMap(.GetSetting("startup", INI_KEYCODE_MOVE_RIGHT, vbKeyRight))
  txtKeyMap(2).Text = showKeyMap(.GetSetting("startup", INI_KEYCODE_DROP, vbKeyDown))
  txtKeyMap(3).Text = showKeyMap(.GetSetting("startup", INI_KEYCODE_ROTATE_CLOCKWISE, vbKeyF))
  txtKeyMap(4).Text = showKeyMap(.GetSetting("startup", INI_KEYCODE_ROTATE_COUNTERCLOCKWISE, vbKeyD))
End With

End Sub



Private Sub gameLoop()
Dim intCounter          As Integer
Dim blnContinue         As Boolean
Dim lngLastTickCount    As Long
Dim intLinesRemoved     As Integer
Dim intLineScore        As Integer
Dim blnContinueGameLoop As Boolean

picBoard.SetFocus
lngLastTickCount = GetTickCount

game.drawNextPiece 'Draw the piece that is stored in the game.nextPiece pPiece array
blnContinueGameLoop = True
Do While blnContinueGameLoop
  intLinesRemoved = 0
  'if a floor collision will not occur
  If game.wallCollisionWillOccur(2, game.piece) = False Then
    If game.pieceCollisionWillOccur(2, game.piece) = False Then
      game.piece.y = game.piece.y + 1
    Else
      If game.pieceCollisionHasOccured(game.piece) = False Then
        blnContinueGameLoop = tryTransferPieceToBoard
      Else
        blnContinueGameLoop = False
      End If
    End If
  Else
    blnContinueGameLoop = tryTransferPieceToBoard
  End If
  game.drawPiece game.piece.shape, game.piece.state 'Draw the piece on the screen
  pause game.gameSpeed 'control speed of game loop
  'each level is 10 lines
  If game.level <> game.lines \ 10 Then
    game.level = game.lines \ 10
  End If
  
  'If Type B game, user is going backwards from 25 lines to 0
  If UCase(game.gameType) = "TYPE B" Then
    If game.lines = 0 Then
      blnContinueGameLoop = False
    End If
  End If
  
  'if user closed the form, end the game
  If Me.Visible = False Then
    If game.useMusic Then
      game.StopMIDIFile MIDI_FILE_0
    End If
    End
  End If
Loop

tmrKeyInput.Enabled = False
If UCase(game.gameType) = "TYPE B" Then
  If game.lines = 0 Then
    game.showGoodJob
  Else
    game.showGameOver
  End If
Else
  game.showGameOver
End If

If game.useMusic Then
  game.StopMIDIFile MIDI_FILE_0
End If

If isHighScore(game.score, game.gameType) Then
  Load frmHighScoreSave
  frmHighScoreSave.Left = 1740
  frmHighScoreSave.Top = 1800
  frmHighScoreSave.Show 1
End If
End Sub

Private Function tryTransferPieceToBoard() As Boolean
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'Purpose: transfers the information that is stored in the piece object to the
'         array that stores the game info
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

Dim intLinesRemoved As Integer
Dim intLineScore    As Integer
tryTransferPieceToBoard = True
'place the piece on the board, reload with a new piece
If game.transferPieceToBoard Then
  intLinesRemoved = game.eraseLines
  If intLinesRemoved <> 0 Then
    intLineScore = (game.level + 1) * (2 ^ intLinesRemoved) * 100
  Else
    intLineScore = 0
  End If
  If UCase(game.gameType) = "TYPE A" Then
    game.lines = game.lines + intLinesRemoved
  Else
    game.lines = game.lines - intLinesRemoved
  End If
  game.piece.y = 0
  game.piece.x = 4
  game.piece.state = 0
  game.piece.shape = game.nextPiece.shape
  game.piece.shapeColor = game.nextPiece.shapeColor
  game.score = game.score + 5000 \ game.levelSpeed + intLineScore
  
  game.gameSpeed = game.levelSpeed
  game.nextPiece.loadRandomShape
  score.drawScore game.score, 2
  If game.level < game.lines \ 10 Then
    game.level = game.lines \ 10
    scoreLevel.drawScore game.level, 2
  End If
  
  If intLinesRemoved <> 0 Then
    scoreLines.drawScore game.lines, 2
  End If
  game.drawNextPiece
Else
  tryTransferPieceToBoard = False
End If


If tryTransferPieceToBoard = True Then
  game.playSoundEffect App.Path & "\place.wav"
End If

End Function


Private Sub mnuFileNewGame_Click()
'Hide the options frame
fraOptions.Visible = False
Me.Height = 3825

game.gameType = gReg.GetSetting("startup", INI_GAME_TYPE, "Type A")
game.useMusic = gReg.GetSetting("startup", INI_USE_MUSIC, "0")
game.useSFX = gReg.GetSetting("startup", INI_USE_SFX, "0")
game.level = gReg.GetSetting("startup", INI_START_LEVEL, "0")
game.keyCodeMoveLeft = gReg.GetSetting("startup", INI_KEYCODE_MOVE_LEFT, vbKeyLeft)
game.keyCodeMoveRight = gReg.GetSetting("startup", INI_KEYCODE_MOVE_RIGHT, vbKeyRight)
game.keyCodeDrop = gReg.GetSetting("startup", INI_KEYCODE_DROP, vbKeyDown)
game.keyCodeRotateClockwise = gReg.GetSetting("startup", INI_KEYCODE_ROTATE_CLOCKWISE, vbKeyF)
game.keyCodeRotateCounterclockwise = gReg.GetSetting("startup", INI_KEYCODE_ROTATE_COUNTERCLOCKWISE, vbKeyD)

game.initialize
If UCase(game.gameType) = "TYPE A" Then
  game.lines = 0
Else
  game.lines = 25
End If

game.score = 0
game.gameSpeed = game.levelSpeed
Select Case UCase(game.gameType)
  Case "TYPE A"
    game.loadLevelTypeA
  Case Else
    game.loadLevelTypeB
End Select

If game.useMusic Then
  game.PlayMIDIFile MIDI_FILE_0
End If

picBoard.SetFocus
game.drawBackground
score.blankScoreBoard
score.drawScore game.score, 2

scoreLines.blankScoreBoard
scoreLines.drawScore game.lines, 2

scoreLevel.blankScoreBoard
scoreLevel.drawScore game.level, 2
Me.tmrKeyInput.Enabled = True
gameLoop
End Sub

Private Sub mnuFileOptions_Click()
If fraOptions.Visible = False Then
  fraOptions.Visible = True
  Me.Height = 7110
Else
  fraOptions.Visible = False
  Me.Height = 3825
End If
End Sub

Private Sub picHigh_Click()
Load frmHighScores
frmHighScores.Show 1
End Sub

Private Sub picOptions_Click()
mnuFileOptions_Click
End Sub

Private Sub picStart_Click()
mnuFileNewGame_Click
End Sub


Private Sub tmrKeyInput_Timer()

If game.keyWasPressedRotateLeft Then
  If game.wallCollisionWillOccur(4, game.piece) = False Then
    If game.pieceCollisionWillOccur(4, game.piece) = False Then
      game.piece.rotateLeft
      game.playSoundEffect App.Path & "\rotate.wav"
    End If
  End If
End If
If game.keyWasPressedRotateRight Then
  If game.wallCollisionWillOccur(3, game.piece) = False Then
    If game.pieceCollisionWillOccur(3, game.piece) = False Then
      game.piece.rotateRight
      game.playSoundEffect App.Path & "\rotate.wav"
    End If
  End If
End If
If game.keyDownLeft Or game.keyWasPressedLeft Then
  If game.wallCollisionWillOccur(0, game.piece) = False Then
    If game.pieceCollisionWillOccur(0, game.piece) = False Then
      If game.moveCount = 1 Then
        If GetTickCount - game.lastMove < 100 Then
          GoTo drawPiece
        End If
      End If
      game.piece.x = game.piece.x - 1
      game.moveCount = game.moveCount + 1
      game.lastMove = GetTickCount
    End If
  End If
End If
If game.keyDownRight Or game.keyWasPressedRight Then
  If game.wallCollisionWillOccur(1, game.piece) = False Then
    If game.pieceCollisionWillOccur(1, game.piece) = False Then
      If game.moveCount = 1 Then
        If GetTickCount - game.lastMove < 100 Then
          GoTo drawPiece
        End If
      End If
      game.piece.x = game.piece.x + 1
      game.moveCount = game.moveCount + 1
      game.lastMove = GetTickCount
    End If
  End If
End If

If game.keyDownDown Then
  If game.piece.y = 0 Then
    GoTo drawPiece
  End If
  If game.gameSpeed <> 20 Then
    game.gameSpeed = 20
    gBlnCancelPause = True 'cancel current game loop pause, if necessary
  End If
Else
  If game.gameSpeed <> game.levelSpeed Then
    game.gameSpeed = game.levelSpeed
  End If
End If

game.keyWasPressedDown = False
game.keyWasPressedLeft = False
game.keyWasPressedRight = False
game.keyWasPressedRotateLeft = False
game.keyWasPressedRotateRight = False
game.keyWasPressedUp = False

drawPiece:

game.drawPiece game.piece.shape, game.piece.state
End Sub

Private Sub txtKeyMap_KeyDown(Index As Integer, keyCode As Integer, Shift As Integer)
Select Case keyCode
  Case vbKeyA To vbKeyZ, vbKey0 To vbKey9
    txtKeyMap(Index).Text = "<" & Chr(keyCode) & ">"
  Case vbKeyLeft
    txtKeyMap(Index).Text = "<Left Arrow>"
  Case vbKeyUp
    txtKeyMap(Index).Text = "<Up Arrow>"
  Case vbKeyRight
    txtKeyMap(Index).Text = "<Right Arrow>"
  Case vbKeyDown
    txtKeyMap(Index).Text = "<Down Arrow>"
  Case vbKeySpace
    txtKeyMap(Index).Text = "<Space Bar>"
  Case vbKeyShift
    txtKeyMap(Index).Text = "<Shift Key>"
  Case vbKeyNumpad0 To vbKeyNumpad9
    txtKeyMap(Index).Text = "<Num Pad " & keyCode - 96 & ">"
  Case Else
    txtKeyMap(Index).Text = "<" & keyCode & ">"
End Select

Select Case Index
  Case 0
    gReg.SaveSetting "startup", INI_KEYCODE_MOVE_LEFT, keyCode
  Case 1
    gReg.SaveSetting "startup", INI_KEYCODE_MOVE_RIGHT, keyCode
  Case 2
    gReg.SaveSetting "startup", INI_KEYCODE_DROP, keyCode
  Case 3
    gReg.SaveSetting "startup", INI_KEYCODE_ROTATE_CLOCKWISE, keyCode
  Case 4
    gReg.SaveSetting "startup", INI_KEYCODE_ROTATE_COUNTERCLOCKWISE, keyCode
End Select

End Sub

Private Function showKeyMap(keyCode As Integer) As String
Select Case keyCode
  Case vbKeyA To vbKeyZ, vbKey0 To vbKey9
    showKeyMap = "<" & Chr(keyCode) & ">"
  Case vbKeyLeft
    showKeyMap = "<Left Arrow>"
  Case vbKeyUp
    showKeyMap = "<Up Arrow>"
  Case vbKeyRight
    showKeyMap = "<Right Arrow>"
  Case vbKeyDown
    showKeyMap = "<Down Arrow>"
  Case vbKeySpace
    showKeyMap = "<Space Bar>"
  Case vbKeyShift
    showKeyMap = "<Shift Key>"
  Case vbKeyNumpad0 To vbKeyNumpad9
    showKeyMap = "<Num Pad " & keyCode - 96 & ">"
  Case Else
    showKeyMap = "<" & keyCode & ">"
End Select
End Function
