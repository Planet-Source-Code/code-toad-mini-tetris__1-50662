VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPiece"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pPiece() As Integer
Private pShape As Integer 'Current Shape
Private pState As Integer 'Current State
Private pX     As Integer 'Shape's x pos on board
Private pY     As Integer 'Shape's y pos on board
Private pShapeColor As Integer

'ppiece(shape,state,introw,intcol)
'there are 7 shapes
'there are four states for each piece (1 state per rotation)

Public Property Get shape() As Integer
shape = pShape
End Property

Public Property Let shape(shape As Integer)
pShape = shape
End Property

Public Property Get shapeColor() As Integer
shapeColor = pShapeColor
End Property

Public Property Let shapeColor(shapeColor As Integer)
pShapeColor = shapeColor
End Property

Public Property Get x() As Integer
x = pX
End Property

Public Property Let x(x As Integer)
pX = x
End Property

Public Property Get y() As Integer
y = pY
End Property

Public Property Let y(y As Integer)
pY = y
End Property
Public Property Get state() As Integer
state = pState
End Property

Public Property Let state(state As Integer)
pState = state
End Property

Public Function getPieceArray() As Variant
getPieceArray = pPiece
End Function

Public Sub rotateRight()
'If in last state, rotate back to original state
If pState = 3 Then
  pState = 0
Else
  pState = pState + 1
End If
End Sub
Public Sub rotateLeft()
'if first state, go to last state
If pState = 0 Then
  pState = 3
Else
  pState = pState - 1
End If
End Sub

Public Sub loadRandomShape()
'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
pShape = Int((6 - 0 + 1) * Rnd + 0)
pShapeColor = Int((3 - 0 + 1) * Rnd + 0)
End Sub

Public Sub initialize()
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'Loads all possible shapes and states into the PPiece array
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Dim intRow   As Integer
Dim intCol   As Integer
Dim intState As Integer
Dim intShape As Integer


ReDim pPiece(0 To 6, 0 To 3, 0 To 3, 0 To 3)

For intShape = 0 To 6
  For intState = 0 To 3
    For intRow = 0 To 3
      For intCol = 0 To 3
        pPiece(intShape, intState, intRow, intCol) = 0
      Next intCol
    Next intRow
  Next intState
Next intShape

For intShape = 0 To 6
  loadShape intShape 'Initialize all states for each shape
Next intShape
End Sub

Private Sub loadShape(intShape As Integer)
Dim intState As Integer
Dim intRow   As Integer
Dim intCol   As Integer

Select Case intShape
  Case 0
    '0000  '0000  '0000  '0000
    '0110  '0110  '0110  '0110
    '0110  '0110  '0110  '0110
    '0000  '0000  '0000  '0000
    For intState = 0 To 3
      For intCol = 1 To 2
        For intRow = 1 To 2
          pPiece(intShape, intState, intRow, intCol) = 1
        Next intRow
      Next intCol
    Next intState
  Case 1
    'st 0  'st 1  'st 2  'st 3
    '0000  '0000  '0000  '0000
    '0011  '0010  '0011  '0010
    '0110  '0011  '0110  '0011
    '0000  '0001  '0000  '0001
    For intState = 0 To 2 Step 2
      pPiece(intShape, intState, 1, 2) = 1
      pPiece(intShape, intState, 1, 3) = 1
      pPiece(intShape, intState, 2, 1) = 1
      pPiece(intShape, intState, 2, 2) = 1
    Next intState
    
    For intState = 1 To 3 Step 2
      pPiece(intShape, intState, 1, 2) = 1
      pPiece(intShape, intState, 2, 2) = 1
      pPiece(intShape, intState, 2, 3) = 1
      pPiece(intShape, intState, 3, 3) = 1
    Next intState
  Case 2
    '0000  '0000  '0000  '0000
    '1100  '0100  '1100  '0100
    '0110  '1100  '0110  '1100
    '0000  '1000  '0000  '1000
    
    For intState = 0 To 2 Step 2
      pPiece(intShape, intState, 1, 0) = 1
      pPiece(intShape, intState, 1, 1) = 1
      pPiece(intShape, intState, 2, 1) = 1
      pPiece(intShape, intState, 2, 2) = 1
    Next intState
    
    For intState = 1 To 3 Step 2
      pPiece(intShape, intState, 1, 1) = 1
      pPiece(intShape, intState, 2, 0) = 1
      pPiece(intShape, intState, 2, 1) = 1
      pPiece(intShape, intState, 3, 0) = 1
    Next intState
  Case 3
    '0100 '0000 '0100 '0000
    '0100 '1111 '0100 '1111
    '0100 '0000 '0100 '0000
    '0100 '0000 '0100 '0000
    For intState = 0 To 2 Step 2
      For intRow = 0 To 3
        pPiece(intShape, intState, intRow, 1) = 1
      Next intRow
    Next intState
    
    For intState = 1 To 3 Step 2
      For intCol = 0 To 3
        pPiece(intShape, intState, 1, intCol) = 1
      Next intCol
    Next intState
  Case 4
    '0000 '0000 '0000 '0000
    '0010 '0010 '0000 '0010
    '0111 '0011 '0111 '0110
    '0000 '0010 '0010 '0010
    
    'load state 0
    pPiece(intShape, 0, 1, 2) = 1
    For intCol = 1 To 3
      pPiece(intShape, 0, 2, intCol) = 1
    Next intCol
    
    'load state 1
    pPiece(intShape, 1, 2, 3) = 1
    For intRow = 1 To 3
      pPiece(intShape, 1, intRow, 2) = 1
    Next intRow
    
    'load state 2
    pPiece(intShape, 2, 3, 2) = 1
    For intCol = 1 To 3
      pPiece(intShape, 2, 2, intCol) = 1
    Next intCol
        
    'load state 3
    pPiece(intShape, 3, 2, 1) = 1
    For intRow = 1 To 3
      pPiece(intShape, 3, intRow, 2) = 1
    Next intRow
  Case 5
    '0000 '0000 '0000 '0000
    '0100 '0000 '1100 '0010
    '0100 '1110 '0100 '1110
    '0110 '1000 '0100 '0000
    
    'load state 0
    pPiece(intShape, 0, 3, 2) = 1
    For intRow = 1 To 3
      pPiece(intShape, 0, intRow, 1) = 1
    Next intRow
    
    'load state 1
    pPiece(intShape, 1, 3, 0) = 1
    For intCol = 0 To 2
      pPiece(intShape, 1, 2, intCol) = 1
    Next intCol
    
    'load state 2
    pPiece(intShape, 2, 1, 0) = 1
    For intRow = 1 To 3
      pPiece(intShape, 2, intRow, 1) = 1
    Next intRow
    
    'load state 3
    pPiece(intShape, 3, 1, 2) = 1
    For intCol = 0 To 2
      pPiece(intShape, 3, 2, intCol) = 1
    Next intCol
  Case 6
    '0000 '0000 '0000 '0000
    '0010 '0100 '0011 '0000
    '0010 '0111 '0010 '0111
    '0110 '0000 '0010 '0001
    
    'load state 0
    pPiece(intShape, 0, 3, 1) = 1
    For intRow = 1 To 3
      pPiece(intShape, 0, intRow, 2) = 1
    Next intRow
    
    'load state 1
    pPiece(intShape, 1, 1, 1) = 1
    For intCol = 1 To 3
      pPiece(intShape, 1, 2, intCol) = 1
    Next intCol
    
    'load state 2
    pPiece(intShape, 2, 1, 3) = 1
    For intRow = 1 To 3
      pPiece(intShape, 2, intRow, 2) = 1
    Next intRow
    
    'load state 3
    pPiece(intShape, 3, 3, 3) = 1
    For intCol = 1 To 3
      pPiece(intShape, 3, 2, intCol) = 1
    Next intCol
End Select
End Sub
