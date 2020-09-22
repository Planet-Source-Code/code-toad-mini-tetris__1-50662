VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private pPicScore            As PictureBox 'Where to draw the score
Private pPicNumbers() As PictureBox '10 is blank pic


Public Sub addNumber(pic As PictureBox)
Dim intUbound As Integer
On Error Resume Next
intUbound = -1
intUbound = UBound(pPicNumbers)
If intUbound = -1 Then
  ReDim pPicNumbers(0 To 0)
  Set pPicNumbers(0) = pic
Else
  ReDim Preserve pPicNumbers(0 To intUbound + 1)
  Set pPicNumbers(UBound(pPicNumbers)) = pic
End If

End Sub

Public Property Let score(picScore As PictureBox)
Set pPicScore = picScore
End Property
Public Property Get cellWidth() As Integer
  cellWidth = pPicNumbers(0).ScaleWidth
End Property
Public Sub blankScoreBoard()
  StretchBlt pPicScore.hdc, 0, 0, pPicScore.ScaleWidth, pPicScore.ScaleHeight, pPicNumbers(10).hdc, 1, 1, 1, 1, vbSrcCopy
End Sub

Public Sub drawScore(ByVal strScore As String, intMethod As Integer)
Dim strTemp As String
Dim intCellWidth As Integer
Dim intCellLeft  As Integer
Dim intNumber    As Integer
Dim intCounter   As Integer
intCellWidth = Me.cellWidth
intCellLeft = pPicScore.ScaleWidth - intCellWidth
Do While Len(strScore) > 0
  strTemp = Right$(strScore, 1)
  If Len(strScore) > 1 Then
    strScore = Left$(strScore, Len(strScore) - 1)
  Else
    strScore = ""
  End If
  intNumber = strTemp
  Select Case intMethod
    Case 2 'standard
      BitBlt pPicScore.hdc, intCellLeft, 0, pPicNumbers(intNumber).ScaleWidth, pPicScore.ScaleHeight, pPicNumbers(intNumber).hdc, 0, 0, vbSrcCopy
    Case 1 'stretch
      For intCounter = 4 To 0 Step -1
        StretchBlt pPicScore.hdc, intCellLeft, intCounter, intCellWidth, pPicScore.ScaleHeight - (intCounter * 2), pPicNumbers(intNumber).hdc, 0, 0, pPicNumbers(intNumber).ScaleWidth, pPicNumbers(intNumber).ScaleHeight, vbSrcCopy
        pause 1
        'DoEvents
      Next intCounter
  End Select
  intCellLeft = intCellLeft - intCellWidth
Loop
End Sub
