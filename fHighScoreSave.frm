VERSION 5.00
Begin VB.Form frmHighScoreSave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "High Score!"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   Icon            =   "fHighScoreSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblInfo 
      Caption         =   "Congratulations!  You got a high score!"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   60
      Width           =   5835
   End
End
Attribute VB_Name = "frmHighScoreSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()

If txtName.Text = "" Then
  MsgBox "You must enter a name."
  Exit Sub
End If
Screen.MousePointer = vbHourglass
saveHighScore game.score, game.lines, txtName.Text, game.gameType
Screen.MousePointer = vbNormal
Me.Hide
DoEvents
Unload Me
End Sub

Private Sub Form_Load()
txtName.Text = "Player 1"
End Sub

Private Sub txtName_GotFocus()
txtName.SelStart = 0
txtName.SelLength = Len(txtName.Text)
End Sub
