VERSION 5.00
Begin VB.Form frmReadKey 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblKey 
      Height          =   195
      Index           =   6
      Left            =   1560
      TabIndex        =   13
      Top             =   1560
      Width           =   1380
   End
   Begin VB.Label lblKeyName 
      Caption         =   "Music:"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1380
   End
   Begin VB.Label lblKey 
      Height          =   195
      Index           =   4
      Left            =   1560
      TabIndex        =   11
      Top             =   1080
      Width           =   1380
   End
   Begin VB.Label lblKeyName 
      Caption         =   "Show/Hide next:"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1380
   End
   Begin VB.Label lblKeyName 
      AutoSize        =   -1  'True
      Caption         =   "Move block left:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1380
   End
   Begin VB.Label lblKeyName 
      Caption         =   "Move block right:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   1380
   End
   Begin VB.Label lblKeyName 
      Caption         =   "Move block down:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   600
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
      Caption         =   "Pause game:"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1380
   End
   Begin VB.Label lblKey 
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   1380
   End
   Begin VB.Label lblKey 
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   1380
   End
   Begin VB.Label lblKey 
      Height          =   195
      Index           =   3
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   1380
   End
   Begin VB.Label lblKey 
      Height          =   195
      Index           =   5
      Left            =   1560
      TabIndex        =   1
      Top             =   1320
      Width           =   1380
   End
   Begin VB.Label lblKey 
      Height          =   195
      Index           =   2
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   1380
   End
End
Attribute VB_Name = "frmReadKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private countKeys As Integer
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If countkey <> -1 Then
        lblKey(countKeys) = getKeyName(KeyCode)
        frmKeyConfig.lblKey(countKeys) = getKeyName(KeyCode)
        frmKeyConfig.lblKey(countKeys).Tag = KeyCode
        If countKeys = 6 Then
            DoEvents
            Sleep (500)
            Unload Me
        Else
            lblKeyName(countKeys).ForeColor = vbBlack
            countKeys = countKeys + 1
            lblKeyName(countKeys).ForeColor = vbYellow
        End If
    End If
End Sub

Private Sub Form_Load()
    countKeys = 0
    lblKeyName(countKeys).ForeColor = vbYellow
End Sub
