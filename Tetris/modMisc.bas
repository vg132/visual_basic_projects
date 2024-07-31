Attribute VB_Name = "modMisc"
Public Const gridSize = 15
Public rotateKey As Integer
Public moveLeftKey As Integer
Public moveRightKey As Integer
Public moveDownKey As Integer
Public pauseKey As Integer
Public musicKey As Integer
Public showNextKey As Integer
Public gamePause As Boolean
Public startHeight As Integer
Public lines As Integer
Public Score As Long
Public speed As Integer

Type HS
    Name As String
    Score As Long
End Type

Public HighScore(4) As HS

Public Sub loadKeys()
    'Load keys to be used in the game
    rotateKey = GetSetting(App.Title, "Config", "RotateKey", 38) 'Default=Up Arrow
    moveLeftKey = GetSetting(App.Title, "Config", "LeftKey", 37) 'Default=Left Arrow
    moveRightKey = GetSetting(App.Title, "Config", "RightKey", 39) 'Default=Right Arrow
    moveDownKey = GetSetting(App.Title, "Config", "DownKey", 40) 'Default=Down Arrow
    showNextKey = GetSetting(App.Title, "Config", "ShowNextKey", 83) 'Default=s
    pauseKey = GetSetting(App.Title, "Config", "PauseKey", 80) 'Default=p
    musicKey = GetSetting(App.Title, "Config", "MusicKey", 77) 'Default m
End Sub

Public Sub addPoint(Points As Integer)
    Score = Score + Points
    frmMain.lblScore.Caption = Score
End Sub

Public Sub addLine()
    lines = lines + 1
    frmMain.lblLines.Caption = lines
End Sub

Public Sub loadHighScoreList()
Dim i As Integer
    For i = 0 To i = 4 Step 1
        HighScore(i).Name = modReg.GetValue(HKEY_CURRENT_USER, "Software\VG Software\Tetris\HighScore", "Name" & i)
        HighScore(i).Score = modReg.GetValue(HKEY_CURRENT_USER, "Software\VG Software\Tetris\HighScore", "Score" & i)
    Next i
End Sub

Public Sub checkScore()

End Sub
