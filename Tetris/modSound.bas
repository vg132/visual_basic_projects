Attribute VB_Name = "modSound"
'This code is from: Break Thru a Freeware VB game
'the game with source code can be downloaded from this url:
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=6135
Private sAlias As String
Private sFilename As String
Private bWait As Boolean
Private nReturn As Long

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Sub mmOpen(ByVal file As String)
    If sAlias <> "" Then
        mmClose
    End If
    sAlias = Right$(file, 3) & Minute(Now)
    If InStr(file, " ") Then file = Chr(34) & file & Chr(34)
    nReturn = mciSendString("Open " & file & " ALIAS " & sAlias & " TYPE Sequencer" & " wait", "", 0, 0)
End Sub

Public Sub mmClose()
    If sAlias = "" Then Exit Sub
    nReturn = mciSendString("Close " & sAlias, "", 0, 0)
    sAlias = ""
    sFilename = ""
End Sub

Public Sub mmPause()
    If sAlias = "" Then Exit Sub
    nReturn = mciSendString("Pause " & sAlias, "", 0, 0)
End Sub

Public Sub mmPlay()
    If sAlias = "" Then Exit Sub
    If bWait Then
        nReturn = mciSendString("Play " & sAlias & " wait", "", 0, 0)
    Else
        nReturn = mciSendString("Play " & sAlias, "", 0, 0)
    End If
End Sub

Public Sub mmStop()
    If sAlias = "" Then Exit Sub
    nReturn = mciSendString("Stop " & sAlias, "", 0, 0)
End Sub

Public Sub mmResume()
    nReturn = mciSendString("Play " & sAlias & "  notify", "", 0, 0)
End Sub

