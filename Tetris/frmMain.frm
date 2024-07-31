VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VG Tetris"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   376
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox imgBlock 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   4800
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   2160
      Top             =   4920
   End
   Begin VB.Label lblLines 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   5
      Top             =   1080
      Width           =   930
   End
   Begin VB.Label lblScore 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Lines:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label lblNext 
      AutoSize        =   -1  'True
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   488
      TabIndex        =   1
      Top             =   0
      Width           =   405
   End
   Begin VB.Image imgNext 
      Height          =   450
      Index           =   6
      Left            =   353
      Picture         =   "frmMain.frx":0E42
      Top             =   247
      Width           =   675
      Visible         =   0   'False
   End
   Begin VB.Image imgNext 
      Height          =   450
      Index           =   5
      Left            =   465
      Picture         =   "frmMain.frx":123B
      Top             =   247
      Width           =   450
      Visible         =   0   'False
   End
   Begin VB.Image imgNext 
      Height          =   450
      Index           =   4
      Left            =   353
      Picture         =   "frmMain.frx":15FC
      Top             =   247
      Width           =   675
      Visible         =   0   'False
   End
   Begin VB.Image imgNext 
      Height          =   450
      Index           =   3
      Left            =   353
      Picture         =   "frmMain.frx":19FD
      Top             =   247
      Width           =   675
      Visible         =   0   'False
   End
   Begin VB.Image imgNext 
      Height          =   450
      Index           =   2
      Left            =   353
      Picture         =   "frmMain.frx":1E00
      Top             =   247
      Width           =   675
      Visible         =   0   'False
   End
   Begin VB.Image imgNext 
      Height          =   450
      Index           =   1
      Left            =   353
      Picture         =   "frmMain.frx":21FF
      Top             =   247
      Width           =   675
      Visible         =   0   'False
   End
   Begin VB.Image imgNext 
      Height          =   225
      Index           =   0
      Left            =   240
      Picture         =   "frmMain.frx":25F7
      Top             =   360
      Width           =   900
      Visible         =   0   'False
   End
   Begin VB.Image imgBlockColor 
      Height          =   225
      Index           =   5
      Left            =   1320
      Picture         =   "frmMain.frx":29B7
      Top             =   5280
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.Image imgBlockColor 
      Height          =   225
      Index           =   2
      Left            =   600
      Picture         =   "frmMain.frx":2A33
      Top             =   5280
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.Image imgBlockColor 
      Height          =   225
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":2AAF
      Top             =   5280
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.Image imgBlockColor 
      Height          =   225
      Index           =   6
      Left            =   1560
      Picture         =   "frmMain.frx":2B2B
      Top             =   5280
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.Image imgBlockColor 
      Height          =   225
      Index           =   4
      Left            =   1080
      Picture         =   "frmMain.frx":2BA7
      Top             =   5280
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.Image imgBlockColor 
      Height          =   225
      Index           =   3
      Left            =   840
      Picture         =   "frmMain.frx":2C23
      Top             =   5280
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.Image imgBlockColor 
      Height          =   225
      Index           =   1
      Left            =   360
      Picture         =   "frmMain.frx":2C9F
      Top             =   5280
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.Image imgGrid 
      Height          =   4530
      Left            =   1680
      Picture         =   "frmMain.frx":2D1B
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2280
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MusicOn As Boolean

Private CurIndex(0 To 3) As Integer
Private blockPos(0 To 3) As blockPosType
Private showNextBlock As Boolean
Private Type blockPosType
    x As Integer
    y As Integer
End Type

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    moveBlock KeyCode
End Sub

Private Sub Form_Load()
'*************************************
'Function Name: Form_Load
'Use: Start a new Tetris game
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 2000-05-01
'*************************************

    gamePause = False
    'loadKeys
    Timer1.Interval = 500 - (speed * 50)
    setStartHeight (startHeight)
    getNewBlock
    newBlock
    MusicOn = True
End Sub

Private Sub imgBlock_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    moveBlock KeyCode
End Sub

Private Sub Timer1_Timer()
'*************************************
'Function Name: Timer1_Timer
'Use: Move block down one grid every time interval
'if block is not freedown then the block will be looked in
'its position and a new block will be created
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 2000-05-01
'*************************************
    Timer1.Interval = 500 - (speed * 50)
    If gamePause = False Then
        If freeDown = True Then
            For i = 0 To 3
                imgBlock(CurIndex(i)).Top = imgBlock(CurIndex(i)).Top + gridSize
            Next
            setBlockPos
        Else
            setBlockPos
            For i = 0 To 3
                If modDetect.setObject(CurIndex(i), blockPos(i).x, blockPos(i).y) = False Then
                    Unload Me
                End If
            Next
            modDetect.checkLine
            newBlock
        End If
    End If
End Sub

Private Sub moveBlock(KeyCode As Integer)
'*************************************
'Function Name: moveBlock
'Use: get user key input and make the right change
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 2000-05-01
'*************************************
    
    If KeyCode = pauseKey Then
        If gamePause = False Then
            gamePause = True
            frmPause.Show vbModal, frmMain
        Else
            gamePause = False
        End If
    End If
    If KeyCode = musicKey Then
        If MusicOn = True Then
            modSound.mmPause
            MusicOn = False
        Else
            modSound.mmResume
            MusicOn = True
        End If
    End If
    If gamePause = False Then
        If (KeyCode = moveLeftKey) And (freeLeft = True) Then
            For i = 0 To 3
                imgBlock(CurIndex(i)).Left = imgBlock(CurIndex(i)).Left - gridSize
            Next
        ElseIf (KeyCode = moveRightKey) And (freeRight = True) Then
            For i = 0 To 3
                imgBlock(CurIndex(i)).Left = imgBlock(CurIndex(i)).Left + gridSize
            Next
        ElseIf (KeyCode = moveDownKey) And (freeDown = True) Then
            For i = 0 To 3
                imgBlock(CurIndex(i)).Top = imgBlock(CurIndex(i)).Top + gridSize
            Next
        ElseIf KeyCode = rotateKey Then
            rotate
        ElseIf KeyCode = showNextKey Then
            If showNextBlock = True Then
                imgNext(blockNext.blockType).Visible = False
                showNextBlock = False
            Else
                imgNext(blockNext.blockType).Visible = True
                showNextBlock = True
            End If
        End If
    End If
    setBlockPos
End Sub

Private Sub setBlockPos()
'*************************************
'Function Name: setBlockPos
'Use: Save current block position into a variable
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 2000-05-01
'*************************************
    For i = 0 To 3
        blockPos(i).x = (imgBlock(CurIndex(i)).Left - imgGrid.Left) / gridSize
        blockPos(i).y = (imgBlock(CurIndex(i)).Top - imgGrid.Top) / gridSize
    Next
End Sub

Private Sub newBlock()
'*************************************
'Function Name: newBlock
'Use: Create a new block
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 2000-05-01
'*************************************
    block = blockNext
    modBlockBuild.getNewBlock
    For i = 0 To 3
        CurIndex(i) = getBlockIndex
        Load imgBlock(CurIndex(i))
        imgBlock(CurIndex(i)).Picture = imgBlockColor(block.blockType)
        imgBlock(CurIndex(i)).Visible = True
        imgBlock(CurIndex(i)).Top = block.blockPos(i).y
        imgBlock(CurIndex(i)).Left = block.blockPos(i).x
    Next
    If showNextBlock = True Then showNext
    setBlockPos
End Sub

Private Function freeDown() As Boolean
'*************************************
'Function Name: freeDown
'Use: Check if the block can be moved down
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 2000-05-01
'*************************************

Dim check As Boolean
    check = True
    For i = 0 To 3
        If Not imgBlock(CurIndex(i)).Top + imgBlock(CurIndex(i)).Width <= imgGrid.Height Then check = False
        If Not modDetect.detectObject(blockPos(i).x, blockPos(i).y + 1) = False Then check = False
    Next
    If check = True Then
        freeDown = True
    Else
        freeDown = False
    End If
End Function

Private Function freeLeft() As Boolean
'*************************************
'Function Name: freeLeft
'Use: Check if the block can be moved to the left
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 2000-05-01
'*************************************

Dim check As Boolean
    check = True
    For i = 0 To 3
        If Not imgBlock(CurIndex(i)).Left > imgGrid.Left + 1 Then check = False
        If Not modDetect.detectObject(blockPos(i).x - 1, blockPos(i).y) = False Then check = False
    Next
    If check = True Then
        freeLeft = True
    Else
        freeLeft = False
    End If
End Function

Private Function freeRight() As Boolean
'*************************************
'Function Name: freeRight
'Use: Check if the block can be moved to the right
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 2000-05-01
'*************************************

Dim check As Boolean
    check = True
    For i = 0 To 3
        If Not imgBlock(CurIndex(i)).Left + imgBlock(CurIndex(i)).Width < imgGrid.Left + imgGrid.Width - 1 Then check = False
        If Not modDetect.detectObject(blockPos(i).x + 1, blockPos(i).y) = False Then check = False
    Next
    If check = True Then
        freeRight = True
    Else
        freeRight = False
    End If
End Function

Private Sub rotate()
'*************************************
'Function Name: rotate
'Use: Rotate the block, first find out what type of block it is
'then call the right rotate function
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 2000-05-01
'*************************************

    If block.blockType = 0 Then
        rotI
    ElseIf block.blockType = 1 Then
        rotL1
    ElseIf block.blockType = 2 Then
        rotT
    ElseIf block.blockType = 3 Then
        rotZ1
    ElseIf block.blockType = 4 Then
        rotZ2
    ElseIf block.blockType = 6 Then
        rotL2
    End If
End Sub

Private Sub rotI()
'*************************************
'Function Name: rotI
'Use: Rotate the I shape block
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 2000-05-01
'*************************************

    If block.viewAngle = Angle.aUp Then
        If (modDetect.detectObject(blockPos(1).x, blockPos(1).y - 1) = False) _
        And (modDetect.detectObject(blockPos(1).x, blockPos(1).y + 1) = False) _
        And (modDetect.detectObject(blockPos(1).x, blockPos(1).y + 2) = False) _
        And (imgBlock(CurIndex(1)).Top + (gridSize * 2) < imgGrid.Top + imgGrid.Height - 1) _
        And (imgBlock(CurIndex(1)).Top - gridSize >= imgGrid.Top + 1) Then
            block.viewAngle = Angle.aLeft
            For i = 0 To 3
                imgBlock(CurIndex(i)).Left = imgBlock(CurIndex(1)).Left
                If i = 0 Then
                    imgBlock(CurIndex(i)).Top = imgBlock(CurIndex(i)).Top - gridSize
                ElseIf i = 2 Then
                    imgBlock(CurIndex(i)).Top = imgBlock(CurIndex(i)).Top + gridSize
                ElseIf i = 3 Then
                    imgBlock(CurIndex(i)).Top = imgBlock(CurIndex(i)).Top + (gridSize * 2)
                End If
            Next
        End If
    Else
        If (modDetect.detectObject(blockPos(1).x - 1, blockPos(1).y) = False) _
        And (modDetect.detectObject(blockPos(1).x + 1, blockPos(1).y) = False) _
        And (modDetect.detectObject(blockPos(1).x + 2, blockPos(1).y) = False) _
        And (imgBlock(CurIndex(1)).Left - gridSize > imgGrid.Left) _
        And (imgBlock(CurIndex(1)).Left + (gridSize * 2) < imgGrid.Left + imgGrid.Width - 1) Then
            block.viewAngle = Angle.aUp
            For i = 0 To 3
                imgBlock(CurIndex(i)).Top = imgBlock(CurIndex(1)).Top
                If i = 0 Then
                    imgBlock(CurIndex(i)).Left = imgBlock(CurIndex(i)).Left - gridSize
                ElseIf i = 2 Then
                    imgBlock(CurIndex(i)).Left = imgBlock(CurIndex(i)).Left + gridSize
                ElseIf i = 3 Then
                    imgBlock(CurIndex(i)).Left = imgBlock(CurIndex(i)).Left + (gridSize * 2)
                End If
            Next
        End If
    End If
End Sub

Private Sub rotT()
'*************************************
'Function Name: rotT
'Use: Rotate the T shape block
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 2000-05-01
'*************************************

    If block.viewAngle = Angle.aUp Then
        If (modDetect.detectObject(blockPos(2).x, blockPos(2).y + 1) = False) _
        And (imgBlock(CurIndex(2)).Top + imgBlock(CurIndex(2)).Height + gridSize < imgGrid.Height + imgGrid.Top - 1) Then
            block.viewAngle = Angle.aLeft
            imgBlock(CurIndex(3)).Top = imgBlock(CurIndex(3)).Top + gridSize
            imgBlock(CurIndex(3)).Left = imgBlock(CurIndex(3)).Left - gridSize
        End If
    ElseIf block.viewAngle = Angle.aLeft Then
        If (modDetect.detectObject(blockPos(2).x + 1, blockPos(2).y) = False) _
        And (imgBlock(CurIndex(2)).Left + imgBlock(CurIndex(2)).Width + gridSize <= imgGrid.Left + imgGrid.Width - 1) Then
            block.viewAngle = Angle.aDown
            imgBlock(CurIndex(0)).Top = imgBlock(CurIndex(0)).Top + gridSize
            imgBlock(CurIndex(0)).Left = imgBlock(CurIndex(0)).Left + gridSize
        End If
    ElseIf block.viewAngle = Angle.aDown Then
        If (modDetect.detectObject(blockPos(2).x, blockPos(2).y - 1) = False) _
        And (imgBlock(CurIndex(2)).Top + gridSize > imgGrid.Top + 1) Then
            block.viewAngle = Angle.aRight
            imgBlock(CurIndex(1)).Top = imgBlock(CurIndex(1)).Top - gridSize
            imgBlock(CurIndex(1)).Left = imgBlock(CurIndex(1)).Left + gridSize
        End If
    ElseIf block.viewAngle = Angle.aRight Then
        If (modDetect.detectObject(blockPos(2).x - 1, blockPos(2).y) = False) _
        And (imgBlock(CurIndex(2)).Left - gridSize >= imgGrid.Left + 1) Then
            block.viewAngle = Angle.aUp
            imgBlock(CurIndex(3)).Top = imgBlock(CurIndex(3)).Top - gridSize
            imgBlock(CurIndex(3)).Left = imgBlock(CurIndex(3)).Left + gridSize
            imgBlock(CurIndex(0)).Top = imgBlock(CurIndex(0)).Top - gridSize
            imgBlock(CurIndex(0)).Left = imgBlock(CurIndex(0)).Left - gridSize
            imgBlock(CurIndex(1)).Top = imgBlock(CurIndex(1)).Top + gridSize
            imgBlock(CurIndex(1)).Left = imgBlock(CurIndex(1)).Left - gridSize
        End If
    End If
End Sub

Private Sub rotZ1()
'*************************************
'Function Name: rotZ1
'Use: Rotate the Z shape block
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 2000-05-01
'*************************************

    If block.viewAngle = Angle.aUp Then
        If (modDetect.detectObject(blockPos(2).x - 1, blockPos(2).y) = False) _
        And (modDetect.detectObject(blockPos(2).x - 1, blockPos(2).y - 1) = False) _
        And (imgBlock(CurIndex(2)).Top - 15 >= imgGrid.Top + 1) Then
            block.viewAngle = Angle.aLeft
            imgBlock(CurIndex(0)).Top = imgBlock(CurIndex(0)).Top - (gridSize * 2)
            imgBlock(CurIndex(3)).Left = imgBlock(CurIndex(3)).Left - (gridSize * 2)
        End If
    ElseIf block.viewAngle = Angle.aLeft Then
        If (modDetect.detectObject(blockPos(1).x - 1, blockPos(1).y) = False) _
        And (modDetect.detectObject(blockPos(2).x + 1, blockPos(2).y) = False) _
        And (imgBlock(CurIndex(2)).Left + imgBlock(CurIndex(2)).Width + gridSize <= imgGrid.Left + imgGrid.Width - 1) Then
            block.viewAngle = Angle.aUp
            imgBlock(CurIndex(0)).Top = imgBlock(CurIndex(0)).Top + (gridSize * 2)
            imgBlock(CurIndex(3)).Left = imgBlock(CurIndex(3)).Left + (gridSize * 2)
        End If
    End If
End Sub

Private Sub rotZ2()
'*************************************
'Function Name: rotZ2
'Use: Rotate the other Z shape
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 2000-05-01
'*************************************

    If block.viewAngle = Angle.aUp Then
        If (modDetect.detectObject(blockPos(1).x + 1, blockPos(1).y) = False) _
        And (modDetect.detectObject(blockPos(1).x + 1, blockPos(1).y - 1) = False) _
        And (imgBlock(CurIndex(0)).Top - gridSize >= imgGrid.Top + 1) Then
            block.viewAngle = Angle.aLeft
            imgBlock(CurIndex(3)).Top = imgBlock(CurIndex(3)).Top - (gridSize * 2)
            imgBlock(CurIndex(0)).Left = imgBlock(CurIndex(0)).Left + (gridSize * 2)
        End If
    ElseIf block.viewAngle = Angle.aLeft Then
        If (modDetect.detectObject(blockPos(1).x - 1, blockPos(1).y) = False) _
        And (modDetect.detectObject(blockPos(2).x + 1, blockPos(2).y) = False) _
        And (imgBlock(CurIndex(1)).Left - gridSize >= imgGrid.Left + 1) Then
            block.viewAngle = Angle.aUp
            imgBlock(CurIndex(3)).Top = imgBlock(CurIndex(3)).Top + (gridSize * 2)
            imgBlock(CurIndex(0)).Left = imgBlock(CurIndex(0)).Left - (gridSize * 2)
        End If
    End If
End Sub

Private Sub rotL1()
'*************************************
'Function Name: rotL1
'Use: Rotate the L shape (green)
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 2000-05-01
'*************************************

    If block.viewAngle = Up Then
        If (modDetect.detectObject(blockPos(2).x, blockPos(2).y - 1) = False) _
        And (modDetect.detectObject(blockPos(2).x + 1, blockPos(2).y + 1) = False) _
        And (modDetect.detectObject(blockPos(2).x, blockPos(2).y + 1) = False) _
        And (imgBlock(CurIndex(2)).Top - gridSize >= imgGrid.Top + 1) Then
            block.viewAngle = Angle.aLeft
            imgBlock(CurIndex(0)).Left = imgBlock(CurIndex(0)).Left + (gridSize * 2)
            imgBlock(CurIndex(1)).Left = imgBlock(CurIndex(1)).Left + gridSize
            imgBlock(CurIndex(3)).Left = imgBlock(CurIndex(3)).Left - gridSize
            imgBlock(CurIndex(1)).Top = imgBlock(CurIndex(1)).Top + gridSize
            imgBlock(CurIndex(3)).Top = imgBlock(CurIndex(3)).Top - gridSize
        End If
    ElseIf block.viewAngle = Angle.aLeft Then

        If (modDetect.detectObject(blockPos(2).x + 1, blockPos(2).y) = False) _
        And (modDetect.detectObject(blockPos(2).x - 1, blockPos(2).y) = False) _
        And (modDetect.detectObject(blockPos(2).x + 1, blockPos(2).y - 1) = False) _
        And (modDetect.detectObject(blockPos(2).x - 1, blockPos(2).y) = False) _
        And (imgBlock(CurIndex(2)).Left - gridSize >= imgGrid.Left + 1) Then
            block.viewAngle = Angle.aDown
            imgBlock(CurIndex(0)).Top = imgBlock(CurIndex(0)).Top - (gridSize * 2)
            imgBlock(CurIndex(1)).Top = imgBlock(CurIndex(1)).Top - gridSize
            imgBlock(CurIndex(3)).Top = imgBlock(CurIndex(3)).Top + gridSize
            imgBlock(CurIndex(1)).Left = imgBlock(CurIndex(1)).Left + gridSize
            imgBlock(CurIndex(3)).Left = imgBlock(CurIndex(3)).Left - gridSize
        End If
    ElseIf block.viewAngle = Angle.aDown Then
        If (modDetect.detectObject(blockPos(2).x, blockPos(2).y + 1) = False) _
        And (modDetect.detectObject(blockPos(2).x, blockPos(2).y - 1) = False) _
        And (modDetect.detectObject(blockPos(2).x - 1, blockPos(2).y - 1) = False) _
        And (imgBlock(CurIndex(2)).Top + imgBlock(CurIndex(2)).Height + gridSize <= imgGrid.Top + imgGrid.Height - 1) Then
            block.viewAngle = Angle.aRight
            imgBlock(CurIndex(0)).Left = imgBlock(CurIndex(0)).Left - (gridSize * 2)
            imgBlock(CurIndex(1)).Top = imgBlock(CurIndex(1)).Top - gridSize
            imgBlock(CurIndex(3)).Top = imgBlock(CurIndex(3)).Top + gridSize
            imgBlock(CurIndex(1)).Left = imgBlock(CurIndex(1)).Left - gridSize
            imgBlock(CurIndex(3)).Left = imgBlock(CurIndex(3)).Left + gridSize
        End If
    ElseIf block.viewAngle = Angle.aRight Then
        If (modDetect.detectObject(blockPos(2).x - 1, blockPos(2).y + 1) = False) _
        And (modDetect.detectObject(blockPos(2).x - 1, blockPos(2).y) = False) _
        And (modDetect.detectObject(blockPos(2).x + 1, blockPos(2).y) = False) _
        And (imgBlock(CurIndex(2)).Left + imgBlock(CurIndex(2)).Width + gridSize <= imgGrid.Left + imgGrid.Width - 1) Then
            block.viewAngle = Angle.aUp
            imgBlock(CurIndex(0)).Top = imgBlock(CurIndex(0)).Top + (gridSize * 2)
            imgBlock(CurIndex(1)).Top = imgBlock(CurIndex(1)).Top + gridSize
            imgBlock(CurIndex(3)).Top = imgBlock(CurIndex(3)).Top - gridSize
            imgBlock(CurIndex(1)).Left = imgBlock(CurIndex(1)).Left - gridSize
            imgBlock(CurIndex(3)).Left = imgBlock(CurIndex(3)).Left + gridSize
       End If
    End If
End Sub

Private Sub rotL2()
'*************************************
'Function Name: rotL2
'Use: Rotate the L shape (brown)
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 2000-05-01
'*************************************

    If block.viewAngle = Angle.aUp Then
        If (modDetect.detectObject(blockPos(1).x, blockPos(1).y - 1) = False) _
        And (modDetect.detectObject(blockPos(1).x + 1, blockPos(1).y - 1) = False) _
        And (modDetect.detectObject(blockPos(1).x, blockPos(1).y + 1) = False) _
        And (imgBlock(CurIndex(1)).Top - 15 >= imgGrid.Top + 1) Then
            block.viewAngle = Angle.aLeft
            imgBlock(CurIndex(0)).Top = imgBlock(CurIndex(0)).Top + gridSize
            imgBlock(CurIndex(0)).Left = imgBlock(CurIndex(0)).Left + gridSize
            imgBlock(CurIndex(2)).Top = imgBlock(CurIndex(2)).Top - gridSize
            imgBlock(CurIndex(2)).Left = imgBlock(CurIndex(2)).Left - gridSize
            imgBlock(CurIndex(3)).Top = imgBlock(CurIndex(3)).Top - (gridSize * 2)
        End If
    ElseIf block.viewAngle = Angle.aLeft Then
        If (modDetect.detectObject(blockPos(1).x - 1, blockPos(1).y - 1) = False) _
        And (modDetect.detectObject(blockPos(1).x - 1, blockPos(1).y) = False) _
        And (modDetect.detectObject(blockPos(1).x + 1, blockPos(1).y) = False) _
        And (imgBlock(CurIndex(1)).Left - gridSize >= imgGrid.Left + 1) Then
            block.viewAngle = Angle.aDown
            imgBlock(CurIndex(0)).Top = imgBlock(CurIndex(0)).Top - gridSize
            imgBlock(CurIndex(0)).Left = imgBlock(CurIndex(0)).Left - gridSize
            imgBlock(CurIndex(2)).Left = imgBlock(CurIndex(2)).Left - gridSize
            imgBlock(CurIndex(3)).Top = imgBlock(CurIndex(3)).Top + gridSize
        End If
    ElseIf block.viewAngle = Angle.aDown Then
        If (modDetect.detectObject(blockPos(1).x - 1, blockPos(1).y + 1) = False) _
        And (imgBlock(CurIndex(1)).Top + imgBlock(CurIndex(1)).Height + gridSize <= imgGrid.Top + imgGrid.Height - 1) Then
            block.viewAngle = Angle.aRight
            imgBlock(CurIndex(0)).Top = imgBlock(CurIndex(0)).Top + gridSize
            imgBlock(CurIndex(2)).Left = imgBlock(CurIndex(2)).Left + gridSize
            imgBlock(CurIndex(3)).Top = imgBlock(CurIndex(3)).Top + gridSize
            imgBlock(CurIndex(3)).Left = imgBlock(CurIndex(3)).Left - gridSize
        End If
    ElseIf block.viewAngle = Angle.aRight Then
        If (modDetect.detectObject(blockPos(1).x + 1, blockPos(1).y + 1) = False) _
        And (modDetect.detectObject(blockPos(1).x + 1, blockPos(1).y) = False) _
        And (modDetect.detectObject(blockPos(1).x - 1, blockPos(1).y) = False) _
        And (imgBlock(CurIndex(1)).Left + imgBlock(CurIndex(1)).Width + gridSize <= imgGrid.Left + imgGrid.Width - 1) Then
            block.viewAngle = Angle.aUp
            imgBlock(CurIndex(0)).Top = imgBlock(CurIndex(0)).Top - gridSize
            imgBlock(CurIndex(2)).Left = imgBlock(CurIndex(2)).Left + gridSize
            imgBlock(CurIndex(2)).Top = imgBlock(CurIndex(2)).Top + gridSize
            imgBlock(CurIndex(3)).Left = imgBlock(CurIndex(3)).Left + gridSize
        End If
    End If
End Sub

Public Sub setStartHeight(startHeight As Integer)
'*************************************
'Function Name: setStartHeight
'Use: set random blocks at the start of the game if
'user have selected this
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 2000-05-01
'*************************************

Dim retVal As Integer
Dim tmpIndex As Integer
Dim ii As Integer
Dim i As Integer
    For i = 19 To (20 - startHeight) Step -1
        For ii = 0 To 9
            Randomize
            retVal = Int(3 * Rnd)
            If retVal = 1 Then
                tmpIndex = getBlockIndex
                modDetect.setObject tmpIndex, ii, i
                Load imgBlock(tmpIndex)
                imgBlock(tmpIndex).Picture = imgBlockColor(Int(6 * Rnd))
                imgBlock(tmpIndex).Left = imgGrid.Left + (gridSize * ii) + 1
                imgBlock(tmpIndex).Top = imgGrid.Top + (gridSize * i) + 1
                imgBlock(tmpIndex).Visible = True
            End If
        Next
    Next
End Sub

Private Sub showNext()
'*************************************
'Function Name: showNext
'Use: show the next block
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 2000-05-01
'*************************************
    
    imgNext(block.blockType).Visible = False
    imgNext(blockNext.blockType).Visible = True
End Sub
