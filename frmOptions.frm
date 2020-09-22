VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   3210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2100
      TabIndex        =   7
      Top             =   1380
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   1020
      TabIndex        =   6
      Top             =   1380
      Width           =   975
   End
   Begin VB.CheckBox chkSFX 
      Caption         =   "Sound Effects"
      Height          =   195
      Left            =   1080
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CheckBox chkMusic 
      Caption         =   "Music"
      Height          =   195
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   1035
   End
   Begin VB.ComboBox cboLevel 
      Height          =   315
      ItemData        =   "frmOptions.frx":0000
      Left            =   1080
      List            =   "frmOptions.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   420
      Width           =   675
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      ItemData        =   "frmOptions.frx":0044
      Left            =   1080
      List            =   "frmOptions.frx":004E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   1935
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Level:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   915
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Game Type:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Me.Hide
DoEvents
Unload Me
End Sub

Private Sub cmdOK_Click()
With gReg
  .SaveSetting "startup", INI_GAME_TYPE, cboType.Text
  .SaveSetting "startup", INI_START_LEVEL, cboLevel.Text
  .SaveSetting "startup", INI_USE_MUSIC, chkMusic.Value
  .SaveSetting "startup", INI_USE_SFX, chkSFX.Value
End With

Me.Hide
DoEvents
Unload Me
End Sub

Private Sub Form_Load()
With gReg
  cboType.Text = .GetSetting("startup", INI_GAME_TYPE, "Type A")
  cboLevel.Text = .GetSetting("startup", INI_START_LEVEL, "0")
  chkMusic.Value = .GetSetting("startup", INI_USE_MUSIC, "0")
  chkSFX.Value = .GetSetting("startup", INI_USE_SFX, "0")
End With
End Sub

