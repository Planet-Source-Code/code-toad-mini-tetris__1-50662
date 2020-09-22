VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function playSound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Const SND_VALID = &H1F            'valid flags  / ;Internal /
Private Const SND_VALIDFLAGS = &H17201F   'Set of valid flag bits.  Anything outside
Private Const SND_TYPE_MASK = &H170007
Private Const SND_SYNC = &H0              'play synchronously (default)
Private Const SND_RESOURCE = &H40004      'name is a resource name or atom
Private Const SND_RESERVED = &HFF000000   'In particular these flags are reserved
Private Const SND_PURGE = &H40            'purge non-static events for task
Private Const SND_NOWAIT = &H2000         'don't wait if the driver is busy
Private Const SND_NOSTOP = &H10           'don't stop any currently playing sound
Private Const SND_NODEFAULT = &H2         'silence not default, if sound not found
Private Const SND_MEMORY = &H4            'lpszSoundName points to a memory file
Private Const SND_LOOP = &H8              'loop the sound until next sndPlaySound
Private Const SND_FILENAME = &H20000      'name is a file name
Private Const SND_ASYNC = &H1             'play asynchronously
Private Const SND_APPLICATION = &H80      'look for application specific association
Private Const SND_ALIAS_START = 0         'must be > 4096 to keep strings in same section of resource file
Private Const SND_ALIAS_ID = &H110000     'name is a WIN.INI [sounds] entry identifier
Private Const SND_ALIAS = &H10000         'name is a WIN.INI [sounds] entry

Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long

Private pPicBoard     As PictureBox 'The playfield
Private pPicGoodJob   As PictureBox
Private pPicGameOver  As PictureBox

Private pPicTile      As PictureBox 'Picture that defines what background will look like
Private pPicBlock()   As PictureBox
Private pPicWall      As PictureBox 'Picture that defines what wall will look like
Private pPicNext      As PictureBox 'Picture that shows the next piece
Private pPicErase     As PictureBox
Private pRowCount     As Long       '# of rows in the playfield
Private pColCount     As Long       '# of cols in the playfield
Private pBoard()      As Integer    'An array that will hold information about the current board
                                    'pBoard(X,Y)=0 : this cell should display the background
                                    'pBoard(X,Y)=1 : this cell should display a wall
                                    'pBoard(X,Y)=2 : this cell should display food
Private pPiece        As New clsPiece
Private pPieceNext    As New clsPiece
Private pPieceArray() As Integer

Private pPlayerCount  As Integer    'The number of players in the game
Private pLevel        As Integer    'The level to load
Private pGameSpeed    As Long       'lower number, faster speed
Private pLevelSpeed   As Long
Private pScore        As Long
Private pLines        As Integer
Private pGameType     As String
Private pUseMusic     As Boolean
Private pUseSFX       As Boolean

Private pKeyDownRotateRight As Boolean
Private pKeyDownRotateLeft  As Boolean
Private pKeyDownLeft        As Boolean
Private pKeyDownRight       As Boolean
Private pKeyDownUp          As Boolean
Private pKeyDownDown        As Boolean

Private pkeyWasPressedRotateRight As Boolean
Private pkeyWasPressedRotateLeft  As Boolean
Private pkeyWasPressedLeft        As Boolean
Private pkeyWasPressedRight       As Boolean
Private pkeyWasPressedUp          As Boolean
Private pkeyWasPressedDown        As Boolean

Private pKeyCodeMoveLeft               As Long
Private pKeyCodeMoveRight              As Long
Private pKeyCodeDrop                   As Long
Private pKeyCodeRotateClockwise        As Long
Private pKeyCodeRotateCounterClockwise As Long
Private pMoveCount                     As Integer
Private pLastMove                      As Long

'pBoard(x,y) tile 0 on/off

Public Property Get keyDownRotateRight() As Boolean
keyDownRotateRight = pKeyDownRotateRight
End Property

Public Property Let keyDownRotateRight(keyDownRotateRight As Boolean)
pKeyDownRotateRight = keyDownRotateRight
End Property

Public Property Get keyCodeMoveLeft() As Long
keyCodeMoveLeft = pKeyCodeMoveLeft
End Property

Public Property Let keyCodeMoveLeft(keyCodeMoveLeft As Long)
pKeyCodeMoveLeft = keyCodeMoveLeft
End Property

Public Property Get keyCodeMoveRight() As Long
keyCodeMoveRight = pKeyCodeMoveRight
End Property

Public Property Let keyCodeMoveRight(keyCodeMoveRight As Long)
pKeyCodeMoveRight = keyCodeMoveRight
End Property

Public Property Get keyCodeDrop() As Long
keyCodeDrop = pKeyCodeDrop
End Property

Public Property Let keyCodeDrop(keyCodeDrop As Long)
pKeyCodeDrop = keyCodeDrop
End Property

Public Property Get keyCodeRotateClockwise() As Long
keyCodeRotateClockwise = pKeyCodeRotateClockwise
End Property

Public Property Let keyCodeRotateClockwise(keyCodeRotateClockwise As Long)
pKeyCodeRotateClockwise = keyCodeRotateClockwise
End Property

Public Property Get keyCodeRotateCounterclockwise() As Long
keyCodeRotateCounterclockwise = pKeyCodeRotateCounterClockwise
End Property

Public Property Let keyCodeRotateCounterclockwise(keyCodeRotateCounterclockwise As Long)
pKeyCodeRotateCounterClockwise = keyCodeRotateCounterclockwise
End Property



Public Property Get lastMove() As Long
lastMove = pLastMove
End Property

Public Property Let lastMove(lastMove As Long)
pLastMove = lastMove
End Property

Public Property Get keyDownRotateLeft() As Boolean
keyDownRotateLeft = pKeyDownRotateLeft
End Property

Public Property Let keyDownRotateLeft(keyDownRotateLeft As Boolean)
pKeyDownRotateLeft = keyDownRotateLeft
End Property

Public Property Get keyDownLeft() As Boolean
keyDownLeft = pKeyDownLeft
End Property

Public Property Let keyDownLeft(keyDownLeft As Boolean)
pKeyDownLeft = keyDownLeft
End Property

Public Property Get keyDownRight() As Boolean
keyDownRight = pKeyDownRight
End Property

Public Property Let keyDownRight(keyDownRight As Boolean)
pKeyDownRight = keyDownRight
End Property

Public Property Get keyDownUp() As Boolean
keyDownUp = pKeyDownUp
End Property

Public Property Let keyDownUp(keyDownUp As Boolean)
pKeyDownUp = keyDownUp
End Property

Public Property Get keyDownDown() As Boolean
keyDownDown = pKeyDownDown
End Property

Public Property Let keyDownDown(keyDownDown As Boolean)
pKeyDownDown = keyDownDown
End Property




Public Property Get keyWasPressedRotateRight() As Boolean
keyWasPressedRotateRight = pkeyWasPressedRotateRight
End Property

Public Property Let keyWasPressedRotateRight(keyWasPressedRotateRight As Boolean)
pkeyWasPressedRotateRight = keyWasPressedRotateRight
End Property

Public Property Get keyWasPressedRotateLeft() As Boolean
keyWasPressedRotateLeft = pkeyWasPressedRotateLeft
End Property

Public Property Let keyWasPressedRotateLeft(keyWasPressedRotateLeft As Boolean)
pkeyWasPressedRotateLeft = keyWasPressedRotateLeft
End Property

Public Property Get keyWasPressedLeft() As Boolean
keyWasPressedLeft = pkeyWasPressedLeft
End Property

Public Property Let keyWasPressedLeft(keyWasPressedLeft As Boolean)
pkeyWasPressedLeft = keyWasPressedLeft
End Property

Public Property Get keyWasPressedRight() As Boolean
keyWasPressedRight = pkeyWasPressedRight
End Property

Public Property Let keyWasPressedRight(keyWasPressedRight As Boolean)
pkeyWasPressedRight = keyWasPressedRight
End Property

Public Property Get keyWasPressedUp() As Boolean
keyWasPressedUp = pkeyWasPressedUp
End Property

Public Property Let keyWasPressedUp(keyWasPressedUp As Boolean)
pkeyWasPressedUp = keyWasPressedUp
End Property

Public Property Get keyWasPressedDown() As Boolean
keyWasPressedDown = pkeyWasPressedDown
End Property

Public Property Let keyWasPressedDown(keyWasPressedDown As Boolean)
pkeyWasPressedDown = keyWasPressedDown
End Property

Public Property Get useMusic() As Boolean
useMusic = pUseMusic
End Property

Public Property Let useMusic(useMusic As Boolean)
pUseMusic = useMusic
End Property

Public Property Get score() As Long
score = pScore
End Property

Public Property Let score(score As Long)
pScore = score
End Property

Public Property Get lines() As Integer
lines = pLines
End Property

Public Property Let lines(lines As Integer)
pLines = lines
End Property

Public Property Get moveCount() As Integer
moveCount = pMoveCount
End Property

Public Property Let moveCount(moveCount As Integer)
pMoveCount = moveCount
End Property


Public Property Get useSFX() As Boolean
useSFX = pUseSFX
End Property

Public Property Let useSFX(useSFX As Boolean)
pUseSFX = useSFX
End Property

Public Property Get gameType() As String
gameType = pGameType
End Property

Public Property Let gameType(gameType As String)
pGameType = gameType
End Property


Public Property Get piece() As clsPiece
Set piece = pPiece
End Property

Public Property Get nextPiece() As clsPiece
Set nextPiece = pPieceNext
End Property

Public Sub addBlock(pic As PictureBox)
Dim intUbound As Integer
On Error Resume Next
intUbound = -1
intUbound = UBound(pPicBlock)
If intUbound = -1 Then
  ReDim pPicBlock(0 To 0)
  Set pPicBlock(0) = pic
Else
  ReDim Preserve pPicBlock(0 To intUbound + 1)
  Set pPicBlock(UBound(pPicBlock)) = pic
End If

End Sub
Public Sub drawNextPiece()
Dim intRow  As Integer
Dim intCol  As Integer

pPieceArray = pPieceNext.getPieceArray
For intRow = 0 To 3
  For intCol = 0 To 3
    If pPieceArray(pPieceNext.shape, pPieceNext.state, intRow, intCol) <> 0 Then
      BitBlt pPicNext.hdc, (intCol) * Me.cellWidth, (intRow) * Me.cellHeight, pPicBlock(0).ScaleWidth, pPicBlock(0).ScaleHeight, pPicBlock(pPieceNext.shapeColor).hdc, 0, 0, vbSrcCopy
    Else 'This is not part of the piece, so draw whatever is on the board
      BitBlt pPicNext.hdc, (intCol) * Me.cellWidth, (intRow) * Me.cellHeight, pPicBlock(0).ScaleWidth, pPicBlock(0).ScaleHeight, pPicTile.hdc, 0, 0, vbSrcCopy
    End If
  Next intCol
Next intRow
End Sub

Public Sub drawPiece(intShape As Integer, intState As Integer)
Dim intRow  As Integer
Dim intCol  As Integer

pPieceArray = pPiece.getPieceArray
For intRow = 0 To 3
  For intCol = 0 To 3
    If pPieceArray(intShape, intState, intRow, intCol) <> 0 Then
      BitBlt pPicBoard.hdc, (pPiece.x + intCol) * Me.cellWidth, (pPiece.y + intRow) * Me.cellHeight, pPicBlock(0).ScaleWidth, pPicBlock(0).ScaleHeight, pPicBlock(pPiece.shapeColor).hdc, 0, 0, vbSrcCopy
    Else 'This is not part of the piece, so draw whatever is on the board
      If pPiece.y + intRow >= 0 And pPiece.y + intRow < Me.rowCount And pPiece.x + intCol >= 0 And pPiece.x + intCol < Me.colCount Then
        If pBoard(pPiece.x + intCol, pPiece.y + intRow) = 0 Then
          BitBlt pPicBoard.hdc, (pPiece.x + intCol) * Me.cellWidth, (pPiece.y + intRow) * Me.cellHeight, pPicBlock(0).ScaleWidth, pPicBlock(0).ScaleHeight, pPicTile.hdc, 0, 0, vbSrcCopy
        Else
          BitBlt pPicBoard.hdc, (pPiece.x + intCol) * Me.cellWidth, (pPiece.y + intRow) * Me.cellHeight, pPicBlock(0).ScaleWidth, pPicBlock(0).ScaleHeight, pPicBlock(pBoard(pPiece.x + intCol, pPiece.y + intRow) - 1).hdc, 0, 0, vbSrcCopy
        End If
      End If
    End If
  Next intCol
Next intRow

'Draw the game board on the cells around the piece
For intCol = 0 To 3
  If pPiece.y - 1 >= 0 And pPiece.x + intCol >= 0 And pPiece.x + intCol < Me.colCount Then
    If pBoard(pPiece.x + intCol, pPiece.y - 1) = 0 Then
      BitBlt pPicBoard.hdc, (pPiece.x + intCol) * Me.cellWidth, (pPiece.y - 1) * Me.cellHeight, pPicBlock(0).ScaleWidth, pPicBlock(0).ScaleHeight, pPicTile.hdc, 0, 0, vbSrcCopy
    Else
      BitBlt pPicBoard.hdc, (pPiece.x + intCol) * Me.cellWidth, (pPiece.y - 1) * Me.cellHeight, pPicBlock(0).ScaleWidth, pPicBlock(0).ScaleHeight, pPicBlock(pBoard(pPiece.x + intCol, pPiece.y - 1) - 1).hdc, 0, 0, vbSrcCopy
    End If
  End If
  
  If pPiece.y + 4 < Me.rowCount And pPiece.y + 4 >= 0 And pPiece.x + intCol >= 0 And pPiece.x + intCol < Me.colCount - 1 Then
    If pBoard(pPiece.x + intCol, pPiece.y + 4) = 0 Then
      BitBlt pPicBoard.hdc, (pPiece.x + intCol) * Me.cellWidth, (pPiece.y + 4) * Me.cellHeight, pPicBlock(0).ScaleWidth, pPicBlock(0).ScaleHeight, pPicTile.hdc, 0, 0, vbSrcCopy
    Else
      BitBlt pPicBoard.hdc, (pPiece.x + intCol) * Me.cellWidth, (pPiece.y + 4) * Me.cellHeight, pPicBlock(0).ScaleWidth, pPicBlock(0).ScaleHeight, pPicBlock(pBoard(pPiece.x + intCol, pPiece.y + 4) - 1).hdc, 0, 0, vbSrcCopy
    End If
  End If
Next intCol

For intRow = 0 To 3
  If pPiece.x - 1 >= 0 And pPiece.y + intRow >= 0 And pPiece.y + intRow < Me.rowCount Then
    If pBoard(pPiece.x - 1, pPiece.y + intRow) = 0 Then
      BitBlt pPicBoard.hdc, (pPiece.x - 1) * Me.cellWidth, (pPiece.y + intRow) * Me.cellHeight, pPicBlock(0).ScaleWidth, pPicBlock(0).ScaleHeight, pPicTile.hdc, 0, 0, vbSrcCopy
    Else
      BitBlt pPicBoard.hdc, (pPiece.x - 1) * Me.cellWidth, (pPiece.y + intRow) * Me.cellHeight, pPicBlock(0).ScaleWidth, pPicBlock(0).ScaleHeight, pPicBlock(pBoard(pPiece.x - 1, pPiece.y + intRow) - 1).hdc, 0, 0, vbSrcCopy
    End If
  End If
  
  If pPiece.x + 4 < Me.colCount And pPiece.x + 4 >= 0 And pPiece.y + intRow >= 0 And pPiece.y + intRow < Me.rowCount Then
    If pBoard(pPiece.x + 4, pPiece.y + intRow) = 0 Then
      BitBlt pPicBoard.hdc, (pPiece.x + 4) * Me.cellWidth, (pPiece.y + intRow) * Me.cellHeight, pPicBlock(0).ScaleWidth, pPicBlock(0).ScaleHeight, pPicTile.hdc, 0, 0, vbSrcCopy
    Else
      BitBlt pPicBoard.hdc, (pPiece.x + 4) * Me.cellWidth, (pPiece.y + intRow) * Me.cellHeight, pPicBlock(0).ScaleWidth, pPicBlock(0).ScaleHeight, pPicBlock(pBoard(pPiece.x + 4, pPiece.y + intRow) - 1).hdc, 0, 0, vbSrcCopy
    End If
  End If
Next intRow

End Sub

Public Function pieceCollisionWillOccur(intOperation As Integer, ByVal testPiece As clsPiece) As Boolean
Dim piece() As Integer
Dim cPiece  As New clsPiece
Dim intRow  As Integer
Dim intCol  As Integer


cPiece.initialize
cPiece.state = testPiece.state
cPiece.shape = testPiece.shape
cPiece.x = testPiece.x '- 1
cPiece.y = testPiece.y
piece = cPiece.getPieceArray


Select Case intOperation
  Case 0 'Trying to Move left
    cPiece.x = cPiece.x - 1
  Case 1 'trying to move right
    cPiece.x = cPiece.x + 1
  Case 2 'Trying to move down
    cPiece.y = cPiece.y + 1
  Case 3
    cPiece.rotateRight
  Case 4
    cPiece.rotateLeft
End Select

For intRow = 0 To 3
  For intCol = 0 To 3
    If piece(cPiece.shape, cPiece.state, intRow, intCol) = 1 Then
      If cPiece.x + intCol >= 0 And cPiece.x + intCol < Me.colCount And cPiece.y + intRow >= 0 And cPiece.y + intRow < Me.rowCount Then
        If pBoard(cPiece.x + intCol, cPiece.y + intRow) <> 0 Then
          pieceCollisionWillOccur = True
          Exit Function
        End If
      End If
    End If
  Next intCol
Next intRow


End Function

Public Function pieceCollisionHasOccured(ByVal testPiece As clsPiece) As Boolean
Dim piece() As Integer
Dim cPiece  As New clsPiece
Dim intRow  As Integer
Dim intCol  As Integer


cPiece.initialize
cPiece.state = testPiece.state
cPiece.shape = testPiece.shape
cPiece.x = testPiece.x '- 1
cPiece.y = testPiece.y
piece = cPiece.getPieceArray



For intRow = 0 To 3
  For intCol = 0 To 3
    If piece(cPiece.shape, cPiece.state, intRow, intCol) = 1 Then
      If cPiece.x + intCol >= 0 And cPiece.x + intCol < Me.colCount And cPiece.y + intRow >= 0 And cPiece.y + intRow < Me.rowCount Then
        If pBoard(cPiece.x + intCol, cPiece.y + intRow) <> 0 Then
          pieceCollisionHasOccured = True
          Exit Function
        End If
      End If
    End If
  Next intCol
Next intRow


End Function

Public Sub drawBackground()
Dim lngCounterY As Long
Dim lngCounterX As Long
Dim lngCellX    As Long
Dim lngCellY    As Long

'Loop through each row
For lngCounterY = 0 To Me.boardHeight - 1 Step Me.cellHeight
  'loop through each column
  lngCellX = 0
  For lngCounterX = 0 To Me.boardWidth - 1 Step Me.cellWidth
    Select Case pBoard(lngCellX, lngCellY)
      Case 0 'Background
        BitBlt pPicBoard.hdc, lngCounterX, lngCounterY, pPicTile.ScaleWidth, pPicTile.ScaleHeight, pPicTile.hdc, 0, 0, vbSrcCopy
      Case Else 'piece
        BitBlt pPicBoard.hdc, lngCounterX, lngCounterY, pPicTile.ScaleWidth, pPicTile.ScaleHeight, pPicBlock(pBoard(lngCellX, lngCellY) - 1).hdc, 0, 0, vbSrcCopy
    End Select
    lngCellX = lngCellX + 1
  Next lngCounterX
  lngCellY = lngCellY + 1
Next lngCounterY
End Sub

Public Function transferPieceToBoard() As Boolean
Dim intRow  As Integer
Dim intCol  As Integer

pPieceArray = pPiece.getPieceArray
transferPieceToBoard = True

For intRow = 0 To 3
  For intCol = 0 To 3
    If pPieceArray(pPiece.shape, pPiece.state, intRow, intCol) <> 0 Then
      If pPiece.y + intRow >= 0 And pPiece.y + intRow < Me.rowCount And pPiece.x + intCol >= 0 And pPiece.x + intCol <= Me.colCount Then
        pBoard(pPiece.x + intCol, pPiece.y + intRow) = pPiece.shapeColor + 1
      Else
        transferPieceToBoard = False
      End If
    End If
  Next intCol
Next intRow

End Function

Public Property Let board(picBoard As PictureBox)
Set pPicBoard = picBoard
pPicBoard.Width = Me.cellWidth * Me.colCount * Screen.TwipsPerPixelX
pPicBoard.Height = Me.cellHeight * Me.rowCount * Screen.TwipsPerPixelY
End Property

Public Property Let picGoodJob(pic As PictureBox)
Set pPicGoodJob = pic
End Property

Public Property Let picGameOver(pic As PictureBox)
Set pPicGameOver = pic
End Property

Public Property Let picNextPiece(picNextPiece As PictureBox)
Set pPicNext = picNextPiece
End Property

Public Property Let picErase(picErase As PictureBox)
Set pPicErase = picErase
End Property

Public Property Let tile(picTile As PictureBox)
Set pPicTile = picTile
End Property

Public Property Get cellWidth() As Long
cellWidth = Me.boardWidth \ Me.colCount 'scale width of the board picture divided by # columns
End Property

Public Property Get cellHeight() As Long
cellHeight = Me.boardHeight \ Me.rowCount 'scale height of the board picture divided by # rows
End Property

Public Property Get boardWidth() As Long
boardWidth = pPicBoard.ScaleWidth
End Property

Public Property Get boardHeight() As Long
boardHeight = pPicBoard.ScaleHeight
End Property

Public Property Get rowCount() As Long
rowCount = pRowCount
End Property

Public Property Let rowCount(rowCount As Long)
pRowCount = rowCount
End Property

Public Property Get gameSpeed() As Long
gameSpeed = pGameSpeed
End Property

Public Property Let gameSpeed(gameSpeed As Long)
pGameSpeed = gameSpeed
End Property

Public Property Get levelSpeed() As Long
If 500 - Me.level * 50 > 0 Then
  levelSpeed = 500 - Me.level * 40
Else
  levelSpeed = 0
End If
'levelSpeed = pLevelSpeed
End Property

'Public Property Let levelSpeed(levelSpeed As Long)
'pLevelSpeed = levelSpeed
'End Property

Public Property Get colCount() As Long
colCount = pColCount
End Property

Public Property Let colCount(colCount As Long)
pColCount = colCount
End Property

Public Property Get playerCount() As Long
playerCount = pPlayerCount
End Property

Public Property Let playerCount(playerCount As Long)
pPlayerCount = playerCount
End Property

Public Property Get level() As Long
level = pLevel
End Property

Public Property Let level(level As Long)
pLevel = level
End Property

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'Public subs and functions
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Public Sub drawGrid()
'This is just a test function, it just draws black lines around each cell
Dim lngCounter As Long

For lngCounter = 0 To Me.boardHeight Step Me.cellHeight
  BitBlt pPicBoard.hdc, 0, lngCounter, Me.boardWidth, 1, pPicBoard.hdc, 0, 0, vbBlackness
Next lngCounter

For lngCounter = 0 To Me.boardWidth Step Me.cellWidth
  BitBlt pPicBoard.hdc, lngCounter, 0, 1, Me.boardHeight, pPicBoard.hdc, 0, 0, vbBlackness
Next lngCounter

End Sub

Public Function wallCollisionWillOccur(intOperation As Integer, ByVal testPiece As clsPiece) As Boolean
Dim piece() As Integer
Dim cPiece  As New clsPiece
Dim intRow  As Integer
Dim intCol  As Integer


cPiece.initialize
cPiece.state = testPiece.state
cPiece.shape = testPiece.shape
cPiece.x = testPiece.x - 1
cPiece.y = testPiece.y
piece = cPiece.getPieceArray


'If trying to move piece left
If intOperation = 0 Then
  For intRow = 0 To 3
    For intCol = 0 To 3
      If piece(cPiece.shape, cPiece.state, intRow, intCol) = 1 Then
        If cPiece.x + intCol < 0 Then
          wallCollisionWillOccur = True
          Exit Function
        End If
      End If
    Next intCol
  Next intRow
End If

'Detect if collision will occur with right wall
If intOperation = 1 Then
  For intRow = 0 To 3
    For intCol = 3 To 0 Step -1
      If piece(cPiece.shape, cPiece.state, intRow, intCol) = 1 Then
        If cPiece.x + intCol + 2 >= game.colCount Then
          wallCollisionWillOccur = True
          Exit Function
        End If
      End If
    Next intCol
  Next intRow
End If

'Detect if collision will occur with floor
If intOperation = 2 Then
  For intRow = 3 To 0 Step -1
    For intCol = 0 To 3
      If piece(cPiece.shape, cPiece.state, intRow, intCol) = 1 Then
        If cPiece.y + intRow + 1 >= game.rowCount Then
          wallCollisionWillOccur = True
          Exit Function
        End If
      End If
    Next intCol
  Next intRow
End If

'Trying to rotate right
If intOperation = 3 Then
  cPiece.rotateRight
  'cPiece.x = cPiece.x - 1
  If Me.wallCollisionWillOccur(1, cPiece) Then
    wallCollisionWillOccur = True
    Exit Function
  End If
  cPiece.x = cPiece.x + 2
  If Me.wallCollisionWillOccur(0, cPiece) Then
    wallCollisionWillOccur = True
    Exit Function
  End If
  cPiece.y = cPiece.y - 1
  If Me.wallCollisionWillOccur(2, cPiece) Then
    wallCollisionWillOccur = True
    Exit Function
  End If
End If

'Trying to rotate left
If intOperation = 4 Then
  cPiece.rotateLeft
  'cPiece.x = cPiece.x - 1
  If Me.wallCollisionWillOccur(1, cPiece) Then
    wallCollisionWillOccur = True
    Exit Function
  End If
  cPiece.x = cPiece.x + 2
  If Me.wallCollisionWillOccur(0, cPiece) Then
    wallCollisionWillOccur = True
    Exit Function
  End If
  cPiece.y = cPiece.y - 1
  If Me.wallCollisionWillOccur(2, cPiece) Then
    wallCollisionWillOccur = True
    Exit Function
  End If
End If
End Function

Public Sub initialize()
Me.keyDownDown = False
Me.keyDownLeft = False
Me.keyDownRight = False
Me.keyDownUp = False
Me.keyDownRotateLeft = False
Me.keyDownRotateRight = False
Me.keyWasPressedDown = False
Me.keyWasPressedLeft = False
Me.keyWasPressedRight = False
Me.keyWasPressedUp = False
Me.keyWasPressedRotateLeft = False
Me.keyWasPressedRotateRight = False
End Sub

Public Sub loadLevelTypeA()
ReDim pBoard(0 To Me.colCount - 1, 0 To Me.rowCount - 1)
End Sub

Public Sub loadLevelTypeB()
Dim intRow As Integer
Dim intCol As Integer
Dim intShapeColor As Integer

ReDim pBoard(0 To Me.colCount - 1, 0 To Me.rowCount - 1)
For intRow = Me.rowCount - 1 To Me.rowCount - ((Me.level * 2) - (Me.level \ 2) + 0) Step -1
  For intCol = 0 To Me.colCount - 1
    intShapeColor = Int((6 - 0 + 1) * Rnd + 0)
    If intShapeColor > 3 Then
      pBoard(intCol, intRow) = 0
    Else
      pBoard(intCol, intRow) = intShapeColor
    End If
  Next intCol
Next intRow
End Sub
Private Sub Class_Initialize()
pPlayerCount = 1
pLevel = 1
End Sub

Public Function eraseLines() As Integer
'returns: number of lines removed

Dim intCol      As Integer
Dim intCol2     As Integer
Dim intRow      As Integer
Dim intRow2     As Integer
Dim rowComplete As Boolean
Dim intErased   As Integer

For intRow = 0 To Me.rowCount - 1
  rowComplete = True
  For intCol = 0 To Me.colCount - 1
    If pBoard(intCol, intRow) = 0 Then
      rowComplete = False
    End If
  Next intCol
  'if this row is completely full of blocks
  If rowComplete = True Then
    'show the lineErase pic
    BitBlt pPicBoard.hdc, 0, intRow * Me.cellHeight, pPicBoard.ScaleWidth, Me.cellHeight, pPicErase.hdc, 0, 0, vbSrcCopy
    pause 1
    intErased = intErased + 1
    For intRow2 = intRow - 1 To 0 Step -1
      For intCol2 = 0 To Me.colCount - 1
        pBoard(intCol2, intRow2 + 1) = pBoard(intCol2, intRow2)
      Next intCol2
    Next intRow2
    Me.drawBackground
  End If
Next intRow

eraseLines = intErased
End Function


Public Sub PlayMIDIFile(strFile As String)
mciExecute ("play " & strFile)
End Sub

Public Sub StopMIDIFile(strFile As String)
mciExecute ("stop " & strFile)
End Sub

Public Sub showGoodJob()
BitBlt pPicBoard.hdc, 0, 0, pPicBoard.ScaleWidth, pPicBoard.ScaleHeight, pPicGoodJob.hdc, 0, 0, vbSrcCopy
End Sub

Public Sub showGameOver()
BitBlt pPicBoard.hdc, 0, 0, pPicBoard.ScaleWidth, pPicBoard.ScaleHeight, pPicGameOver.hdc, 0, 0, vbSrcCopy
End Sub

Public Sub playSoundEffect(strFile As String)
If Me.useSFX Then
    playSound strFile, 0, SND_FILENAME Or SND_ASYNC
End If
End Sub
