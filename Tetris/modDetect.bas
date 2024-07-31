Attribute VB_Name = "modDetect"
Private gridMetrix(0 To 9, 0 To 19) As Integer
Private indexMetrix(0 To 9, 0 To 19) As Integer
Private i As Integer
Private ii As Integer

Public Function setObject(index As Integer, x As Integer, y As Integer) As Boolean
    If detectObject(x, y) = False Then
        gridMetrix(x, y) = 1
        indexMetrix(x, y) = index
        addPoint ((speed + 1) * (5 + lines + block.blockType) + y)
        setObject = True
    Else
        setObject = False
    End If
End Function

Public Function detectObject(x As Integer, y As Integer) As Boolean
    If y > 19 Then y = 19
    If x > 9 Then x = 9
    If y < 0 Then y = 0
    If x < 0 Then x = 0
    If gridMetrix(x, y) = 1 Then
        detectObject = True
    Else
        detectObject = False
    End If
End Function

Public Sub checkLine()
    For i = 0 To 19
        For ii = 0 To 9
            If gridMetrix(ii, i) = 0 Then Exit For
        Next
        If ii = 10 Then removeLine (i)
    Next
End Sub

Public Sub moveDown(toY As Integer)
    For i = toY To 0 Step -1
        For ii = 0 To 9
            If gridMetrix(ii, i) = 1 Then
                frmMain.imgBlock(indexMetrix(ii, i)).Top = frmMain.imgBlock(indexMetrix(ii, i)).Top + gridSize
                gridMetrix(ii, i) = 0
                gridMetrix(ii, i + 1) = 1
                indexMetrix(ii, i + 1) = indexMetrix(ii, i)
                indexMetrix(ii, i) = 0
            End If
        Next
    Next
End Sub

Public Sub removeLine(y As Integer)
    For x = 0 To 9
        Unload frmMain.imgBlock(indexMetrix(x, y))
        modBlock.setFreeBlockIndex indexMetrix(x, y)
        indexMetrix(x, y) = 0
        gridMetrix(x, y) = 0
    Next
    addLine
    addPoint ((speed + 1) * 23 * (lines + 1))
    If (lines = (speed * 10) + 10) And (speed < 8) Then speed = speed + 1
    moveDown (y - 1)
End Sub

Public Sub removeAllBlocks()
    On Error Resume Next
    For i = 0 To 19
        For ii = 0 To 9
            If indexMetrix(ii, i) <> 0 Then
                Unload frmMain.imgBlock(indexMetrix(ii, i))
                modBlock.setFreeBlockIndex indexMetrix(ii, i)
                indexMetrix(ii, i) = 0
                gridMetrix(ii, i) = 0
            End If
        Next
    Next
End Sub

Public Sub resetAllMetrix()
    For i = 0 To 19
        For ii = 0 To 9
            gridMetrix(ii, i) = 0
            indexMetrix(ii, i) = 0
        Next
    Next
End Sub
