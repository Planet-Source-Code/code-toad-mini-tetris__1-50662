Attribute VB_Name = "modGlobal"
Option Explicit
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public gBlnCancelPause As Boolean
Public game            As New clsBoard
Public score           As New clsScore
Public scoreLines      As New clsScore
Public scoreLevel      As New clsScore
Public gReg            As New clsRegSettings
Public gStrINIFile     As String

Public Const INI_USE_MUSIC   As String = "useMusic"
Public Const INI_USE_SFX     As String = "useSoundEffects"
Public Const INI_GAME_TYPE   As String = "gameType"
Public Const INI_START_LEVEL As String = "startLevel"
Public Const INI_KEYCODE_MOVE_LEFT As String = "KeyCodeMoveLeft"
Public Const INI_KEYCODE_MOVE_RIGHT As String = "KeyCodeMoveRight"
Public Const INI_KEYCODE_DROP As String = "KeyCodeDrop"
Public Const INI_KEYCODE_ROTATE_CLOCKWISE As String = "KeyCodeRotateClockwise"
Public Const INI_KEYCODE_ROTATE_COUNTERCLOCKWISE As String = "KeyCodeRotateCounterClockwise"

Public Const MIDI_FILE_0     As String = "c:\tet.mid"

Public Sub pause(lngMilliseconds As Long)
'Pauses for a specified # of milliseconds
Dim blnContinue As Boolean
Dim lngLastTickCount As Long
gBlnCancelPause = False
blnContinue = False
lngLastTickCount = GetTickCount
Do While blnContinue = False
  If gBlnCancelPause Then
    blnContinue = True
  End If
  If GetTickCount - lngLastTickCount >= lngMilliseconds Then
    lngLastTickCount = GetTickCount
    blnContinue = True
  End If
  DoEvents
Loop
End Sub

Public Function isHighScore(lngScore As Long, strGameType As String) As Boolean
Dim intCounter   As Integer
Dim intStart     As Integer
Dim lngHighScore As Long

Select Case UCase(strGameType)
  Case "TYPE A"
    intStart = 0
  Case Else
    intStart = 10
End Select
For intCounter = intStart To intStart + 9
    lngHighScore = Decrypt(gReg.GetSetting("startup", "highScoreValue" & intCounter, "01"), "test")
    'If this score is higher than at least one score in the list, it is a high score
    If lngScore > lngHighScore Then
      isHighScore = True
      Exit Function
    End If
Next intCounter
End Function

Public Function saveHighScore(lngScore As Long, intLines As Integer, strName As String, strGameType As String) As Boolean
Dim intCounter   As Integer
Dim intCounter2  As Integer
Dim intStart     As Integer
Dim intHighScore As Integer

Select Case UCase(strGameType)
  Case "TYPE A"
    intStart = 0
  Case Else
    intStart = 10
End Select

For intCounter = intStart To intStart + 9
    intHighScore = Decrypt(gReg.GetSetting("startup", "highScoreValue" & intCounter, "01"), "test")
    'If this score is higher than at least one score in the list, it is a high score
    If lngScore > intHighScore Then
      If intCounter <> intStart + 9 Then
        'Shift the old scores down
        For intCounter2 = intStart + 9 To intCounter + 1 Step -1
          gReg.SaveSetting "startup", "highScoreValue" & intCounter2, gReg.GetSetting("startup", "highScoreValue" & intCounter2 - 1)
          gReg.SaveSetting "startup", "highScoreName" & intCounter2, gReg.GetSetting("startup", "highScoreName" & intCounter2 - 1)
          gReg.SaveSetting "startup", "highScoreLines" & intCounter2, gReg.GetSetting("startup", "highScoreLines" & intCounter2 - 1)
          gReg.SaveSetting "startup", "highScoreDate" & intCounter2, gReg.GetSetting("startup", "highScoreDate" & intCounter2 - 1)
        Next intCounter2
             
      End If
      gReg.SaveSetting "startup", "highScoreValue" & intCounter, Encrypt(Str(lngScore), "test")
      gReg.SaveSetting "startup", "highScoreName" & intCounter, Encrypt(strName, "test")
      gReg.SaveSetting "startup", "highScoreLines" & intCounter, Encrypt(Str(intLines), "test")
      gReg.SaveSetting "startup", "highScoreDate" & intCounter, Encrypt(Format$(Now, "mm/dd/yy"), "test")
      Exit Function
    
    End If
Next intCounter
End Function

Function Encrypt(strEncrypt As String, strPassword As String) As String
'============================================================================='
''PURPOSE:  TO ENCRYPT A STRING                                              ''
''NOTE:     I DIDN'T WRITE THIS SUB, DON'T KNOW ANYTHING ABOUT IT            ''
'============================================================================='
Dim b As String
Dim s As String
Dim i As Long
Dim j As Long
Dim A1 As Long
Dim A2 As Long
Dim A3 As Long
Dim p As String
j = 1
For i = 1 To Len(strPassword)
  p = p & Asc(Mid$(strPassword, i, 1))
Next
For i = 1 To Len(strEncrypt)
  A1 = Asc(Mid(p, j, 1))
  j = j + 1: If j > Len(p) Then j = 1
  A2 = Asc(Mid$(strEncrypt, i, 1))
  A3 = A1 Xor A2
  b$ = Hex$(A3)
  If Len(b) < 2 Then b = "0" + b
  s = s + b
Next
Encrypt = s
End Function

Public Function Decrypt(strDecrypt, strPassword) As String
'============================================================================='
''PURPOSE:  TO ENCRYPT A STRING                                              ''
''NOTE:     I DIDN'T WRITE THIS SUB, DON'T KNOW ANYTHING ABOUT IT            ''
'============================================================================='

Dim b As String
Dim s As String
Dim i As Long
Dim j As Long
Dim A1 As Long
Dim A2 As Long
Dim A3 As Long
Dim p As String
j = 1
For i = 1 To Len(strPassword)
  p = p & Asc(Mid$(strPassword, i, 1))
Next
For i = 1 To Len(strDecrypt) Step 2
  A1 = Asc(Mid$(p, j, 1))
  j = j + 1: If j > Len(p) Then j = 1
  b = Mid$(strDecrypt, i, 2)
  A3 = Val("&H" + b)
  A2 = A1 Xor A3
  s = s + Chr$(A2)
Next
Decrypt = s
End Function


