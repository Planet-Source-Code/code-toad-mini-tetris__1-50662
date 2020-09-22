VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmHighScores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "High Scores"
   ClientHeight    =   3345
   ClientLeft      =   315
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "fHighScores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   2955
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5212
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Type A"
      TabPicture(0)   =   "fHighScores.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lvwScores(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Type B"
      TabPicture(1)   =   "fHighScores.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwScores(1)"
      Tab(1).ControlCount=   1
      Begin MSComctlLib.ListView lvwScores 
         Height          =   2535
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   360
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Rank"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   4339
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Score"
            Object.Width           =   1109
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Lines"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ListView lvwScores 
         Height          =   2535
         Index           =   1
         Left            =   -74940
         TabIndex        =   3
         Top             =   360
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Rank"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   4339
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Score"
            Object.Width           =   1109
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Lines"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date"
            Object.Width           =   2646
         EndProperty
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   3000
      Width           =   1395
   End
End
Attribute VB_Name = "frmHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Me.Hide
DoEvents
Unload Me
End Sub

Private Sub Form_Load()
Dim item As ListItem
Dim intCounter As Integer
Dim intLvw As Integer
Dim intRank As Integer

For intCounter = 0 To 19
  If intCounter = 0 Or intCounter = 10 Then
    intRank = 1
  End If
  If intCounter < 10 Then
    intLvw = 0
  Else
    intLvw = 1
  End If
  Set item = lvwScores(intLvw).ListItems.Add(, , intRank)
  item.SubItems(1) = Decrypt(gReg.GetSetting("startup", "highScoreName" & intCounter), "test")
  item.SubItems(2) = Decrypt(gReg.GetSetting("startup", "highScoreValue" & intCounter), "test")
  item.SubItems(3) = Decrypt(gReg.GetSetting("startup", "highScoreLines" & intCounter), "test")
  item.SubItems(4) = Decrypt(gReg.GetSetting("startup", "highScoreDate" & intCounter, ""), "test")
  intRank = intRank + 1
Next intCounter
End Sub
