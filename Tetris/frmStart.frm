VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tetris"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3945
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   315
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   1035
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   315
      Left            =   2820
      TabIndex        =   17
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Frame fraHeight 
      Caption         =   "Start Height"
      Height          =   1335
      Left            =   2400
      TabIndex        =   10
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton optHeight 
         Caption         =   "15"
         Height          =   195
         Index           =   5
         Left            =   720
         TabIndex        =   16
         Top             =   840
         Width           =   495
      End
      Begin VB.OptionButton optHeight 
         Caption         =   "12"
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   15
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optHeight 
         Caption         =   "9"
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton optHeight 
         Caption         =   "6"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   495
      End
      Begin VB.OptionButton optHeight 
         Caption         =   "3"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optHeight 
         Caption         =   "0"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame fraSpeed 
      Caption         =   "Start Speed"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton optSpeed 
         Caption         =   "4"
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton optSpeed 
         Caption         =   "9"
         Height          =   195
         Index           =   8
         Left            =   1320
         TabIndex        =   8
         Top             =   840
         Width           =   495
      End
      Begin VB.OptionButton optSpeed 
         Caption         =   "8"
         Height          =   195
         Index           =   7
         Left            =   1320
         TabIndex        =   7
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optSpeed 
         Caption         =   "7"
         Height          =   195
         Index           =   6
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton optSpeed 
         Caption         =   "6"
         Height          =   195
         Index           =   5
         Left            =   720
         TabIndex        =   5
         Top             =   840
         Width           =   495
      End
      Begin VB.OptionButton optSpeed 
         Caption         =   "5"
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optSpeed 
         Caption         =   "3"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   495
      End
      Begin VB.OptionButton optSpeed 
         Caption         =   "2"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optSpeed 
         Caption         =   "1"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuKeyConfig 
         Caption         =   "&Key Config..."
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ProgramDir As String

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdStart_Click()
Dim i As Integer
    frmStart.Hide
    For i = 0 To 8
        If optSpeed(i).Value = True Then speed = i
    Next
    For i = 0 To 5
        If optHeight(i).Value = True Then startHeight = optHeight(i).Caption
    Next
    Randomize
    i = Int(3 * Rnd) + 1
    modSound.mmOpen ProgramDir + "music\" & i & ".mid"
    modSound.mmPlay
    frmMain.Show vbModal, frmStart
    modSound.mmClose
    frmStart.Show
    Unload frmMain
    resetAllMetrix
    resetBlockIndex
End Sub

Private Sub Form_Load()
    ProgramDir = App.Path
    If Right(ProgramDir, 1) <> "\" Then ProgramDir = ProgramDir & "\"
    loadKeys
    loadHighScoreList
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuKeyConfig_Click()
    frmKeyConfig.Show vbModal, frmStart
End Sub
