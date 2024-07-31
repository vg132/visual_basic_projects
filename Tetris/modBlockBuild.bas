Attribute VB_Name = "modBlockBuild"
Public blockNext As blockInfo
Public block As blockInfo
Private Type xy
    x As Integer
    y As Integer
End Type

Public Enum Angle
    aUp = 0
    aRight = 1
    aDown = 2
    aLeft = 3
End Enum

Private Type blockInfo
    blockType As Integer '0=I, 1=L, 2=T, 3=Normal Z, 4=Mirror Z, 5=Block, 6=Mirror L
    viewAngle As Angle
    blockPos(0 To 3) As xy
End Type

Public Sub getNewBlock()
    With blockNext
        Randomize
        .blockType = Int(7 * Rnd)
        If .blockType = 0 Then
            'I
            .blockPos(0).x = frmMain.imgGrid.Left + 1 + (gridSize * 4)
            .blockPos(1).x = frmMain.imgGrid.Left + 1 + (gridSize * 5)
            .blockPos(2).x = frmMain.imgGrid.Left + 1 + (gridSize * 6)
            .blockPos(3).x = frmMain.imgGrid.Left + 1 + (gridSize * 7)
            .blockPos(0).y = frmMain.imgGrid.Top + 1
            .blockPos(1).y = frmMain.imgGrid.Top + 1
            .blockPos(2).y = frmMain.imgGrid.Top + 1
            .blockPos(3).y = frmMain.imgGrid.Top + 1
            .viewAngle = Angle.aUp
        ElseIf .blockType = 1 Then
            'L
            .blockPos(0).x = frmMain.imgGrid.Left + 1 + (gridSize * 4)
            .blockPos(1).x = frmMain.imgGrid.Left + 1 + (gridSize * 4)
            .blockPos(2).x = frmMain.imgGrid.Left + 1 + (gridSize * 5)
            .blockPos(3).x = frmMain.imgGrid.Left + 1 + (gridSize * 6)
            .blockPos(0).y = frmMain.imgGrid.Top + 16
            .blockPos(1).y = frmMain.imgGrid.Top + 1
            .blockPos(2).y = frmMain.imgGrid.Top + 1
            .blockPos(3).y = frmMain.imgGrid.Top + 1
            .viewAngle = Angle.aUp
        ElseIf .blockType = 2 Then
            'T
            .blockPos(0).x = frmMain.imgGrid.Left + 1 + (gridSize * 6)
            .blockPos(1).x = frmMain.imgGrid.Left + 1 + (gridSize * 5)
            .blockPos(2).x = frmMain.imgGrid.Left + 1 + (gridSize * 6)
            .blockPos(3).x = frmMain.imgGrid.Left + 1 + (gridSize * 7)
            .blockPos(0).y = frmMain.imgGrid.Top + 1
            .blockPos(1).y = frmMain.imgGrid.Top + 16
            .blockPos(2).y = frmMain.imgGrid.Top + 16
            .blockPos(3).y = frmMain.imgGrid.Top + 16
            .viewAngle = Angle.aUp
        ElseIf .blockType = 3 Then
            'Normal Z
            .blockPos(0).x = frmMain.imgGrid.Left + 1 + (gridSize * 4)
            .blockPos(1).x = frmMain.imgGrid.Left + 1 + (gridSize * 5)
            .blockPos(2).x = frmMain.imgGrid.Left + 1 + (gridSize * 5)
            .blockPos(3).x = frmMain.imgGrid.Left + 1 + (gridSize * 6)
            .blockPos(0).y = frmMain.imgGrid.Top + 16
            .blockPos(1).y = frmMain.imgGrid.Top + 16
            .blockPos(2).y = frmMain.imgGrid.Top + 1
            .blockPos(3).y = frmMain.imgGrid.Top + 1
            .viewAngle = Angle.aUp
        ElseIf .blockType = 4 Then
            'Onormalt Z
            .blockPos(0).x = frmMain.imgGrid.Left + 1 + (gridSize * 4)
            .blockPos(1).x = frmMain.imgGrid.Left + 1 + (gridSize * 5)
            .blockPos(2).x = frmMain.imgGrid.Left + 1 + (gridSize * 5)
            .blockPos(3).x = frmMain.imgGrid.Left + 1 + (gridSize * 6)
            .blockPos(0).y = frmMain.imgGrid.Top + 1
            .blockPos(1).y = frmMain.imgGrid.Top + 1
            .blockPos(2).y = frmMain.imgGrid.Top + 16
            .blockPos(3).y = frmMain.imgGrid.Top + 16
            .viewAngle = Angle.aUp
        ElseIf .blockType = 5 Then
            'Block
            .blockPos(0).x = frmMain.imgGrid.Left + 1 + (gridSize * 4)
            .blockPos(1).x = frmMain.imgGrid.Left + 1 + (gridSize * 5)
            .blockPos(2).x = frmMain.imgGrid.Left + 1 + (gridSize * 4)
            .blockPos(3).x = frmMain.imgGrid.Left + 1 + (gridSize * 5)
            .blockPos(0).y = frmMain.imgGrid.Top + 1
            .blockPos(1).y = frmMain.imgGrid.Top + 1
            .blockPos(2).y = frmMain.imgGrid.Top + 16
            .blockPos(3).y = frmMain.imgGrid.Top + 16
            .viewAngle = Angle.aUp
        ElseIf .blockType = 6 Then
            'L
            .blockPos(0).x = frmMain.imgGrid.Left + 1 + (gridSize * 4)
            .blockPos(1).x = frmMain.imgGrid.Left + 1 + (gridSize * 5)
            .blockPos(2).x = frmMain.imgGrid.Left + 1 + (gridSize * 6)
            .blockPos(3).x = frmMain.imgGrid.Left + 1 + (gridSize * 6)
            .blockPos(0).y = frmMain.imgGrid.Top + 1
            .blockPos(1).y = frmMain.imgGrid.Top + 1
            .blockPos(2).y = frmMain.imgGrid.Top + 1
            .blockPos(3).y = frmMain.imgGrid.Top + 16
            .viewAngle = Angle.aUp
        End If
    End With
End Sub
