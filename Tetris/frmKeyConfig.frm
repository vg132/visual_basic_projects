VERSION 5.00
Begin VB.Form frmKeyConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tetris Key Config"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "frmKeyConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   3000
      TabIndex        =   0
      Top             =   600
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Top             =   960
      Width           =   1035
   End
   Begin VB.CommandButton cmdResetKeys 
      Caption         =   "Reconfig "
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label lblKey 
      Height          =   195
      Index           =   6
      Left            =   1560
      TabIndex        =   16
      Top             =   1560
      Width           =   1380
   End
   Begin VB.Label lblKeyName 
      Caption         =   "Music:"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   1380
   End
   Begin VB.Label lblKey 
      Height          =   195
      Index           =   4
      Left            =   1560
      TabIndex        =   14
      Top             =   1080
      Width           =   1380
   End
   Begin VB.Label lblKeyName 
      Caption         =   "Show/Hide next:"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   1380
   End
   Begin VB.Label lblKey 
      Height          =   195
      Index           =   2
      Left            =   1560
      TabIndex        =   12
      Top             =   600
      Width           =   1380
   End
   Begin VB.Label lblKey 
      Height          =   195
      Index           =   5
      Left            =   1560
      TabIndex        =   11
      Top             =   1320
      Width           =   1380
   End
   Begin VB.Label lblKey 
      Height          =   195
      Index           =   3
      Left            =   1560
      TabIndex        =   10
      Top             =   840
      Width           =   1380
   End
   Begin VB.Label lblKey 
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   9
      Top             =   360
      Width           =   1380
   End
   Begin VB.Label lblKey 
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   8
      Top             =   120
      Width           =   1380
   End
   Begin VB.Label lblKeyName 
      Caption         =   "Pause game:"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1380
   End
   Begin VB.Label lblKeyName 
      Caption         =   "Rotate block:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1380
   End
   Begin VB.Label lblKeyName 
      Caption         =   "Move block down:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1380
   End
   Begin VB.Label lblKeyName 
      Caption         =   "Move block right:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1380
   End
   Begin VB.Label lblKeyName 
      AutoSize        =   -1  'True
      Caption         =   "Move block left:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1380
   End
End
Attribute VB_Name = "frmKeyConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SaveSetting App.Title, "Config", "LeftKey", lblKey(0).Tag
    SaveSetting App.Title, "Config", "RightKey", lblKey(1).Tag
    SaveSetting App.Title, "Config", "DownKey", lblKey(2).Tag
    SaveSetting App.Title, "Config", "RotateKey", lblKey(3).Tag
    SaveSetting App.Title, "Config", "ShowNextKey", lblKey(4).Tag
    SaveSetting App.Title, "Config", "PauseKey", lblKey(5).Tag
    SaveSetting App.Title, "Config", "MusicKey", lblKey(6).Tag
    DoEvents
    loadKeys
    Unload Me
End Sub

Private Sub cmdResetKeys_Click()
    frmReadKey.Show vbModal, frmKeyConfig
End Sub

Private Sub Form_Load()
    lblKey(0).Tag = moveLeftKey
    lblKey(1).Tag = moveRightKey
    lblKey(2).Tag = moveDownKey
    lblKey(3).Tag = rotateKey
    lblKey(4).Tag = showNextKey
    lblKey(5).Tag = pauseKey
    lblKey(6).Tag = musicKey
    For i = 0 To 6
        lblKey(i).Caption = getKeyName(lblKey(i).Tag)
    Next
End Sub
