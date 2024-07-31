VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmYatzy 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Super YATZY"
   ClientHeight    =   7200
   ClientLeft      =   1830
   ClientTop       =   2115
   ClientWidth     =   8745
   Icon            =   "Yatzy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Yatzy.frx":0442
   ScaleHeight     =   7200
   ScaleWidth      =   8745
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   145
      Top             =   6945
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtScore2 
      Height          =   285
      Left            =   120
      TabIndex        =   144
      Text            =   "90"
      Top             =   5880
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.TextBox txtName2 
      Height          =   285
      Left            =   1320
      TabIndex        =   143
      Text            =   "Trench"
      Top             =   5880
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.TextBox txtDate2 
      Height          =   285
      Left            =   2520
      TabIndex        =   142
      Text            =   "1997-09-07"
      Top             =   5880
      Width           =   1095
      Visible         =   0   'False
   End
   Begin VB.TextBox txtScore3 
      Height          =   285
      Left            =   120
      TabIndex        =   141
      Text            =   "80"
      Top             =   6240
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.TextBox txtName3 
      Height          =   285
      Left            =   1320
      TabIndex        =   140
      Text            =   "Larsson"
      Top             =   6240
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.TextBox txtDate3 
      Height          =   285
      Left            =   2520
      TabIndex        =   139
      Text            =   "1997-09-07"
      Top             =   6240
      Width           =   1095
      Visible         =   0   'False
   End
   Begin VB.TextBox txtDate1 
      Height          =   285
      Left            =   2520
      TabIndex        =   138
      Text            =   "1997-09-07"
      Top             =   5520
      Width           =   1095
      Visible         =   0   'False
   End
   Begin VB.TextBox txtName1 
      Height          =   285
      Left            =   1320
      TabIndex        =   137
      Text            =   "NOE"
      Top             =   5520
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.TextBox txtScore1 
      Height          =   285
      Left            =   120
      TabIndex        =   136
      Text            =   "100"
      Top             =   5520
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.Timer trmTimer 
      Interval        =   10
      Left            =   0
      Top             =   3720
   End
   Begin VB.CommandButton cmdKlara 
      Caption         =   "Klar"
      Height          =   375
      Left            =   840
      TabIndex        =   130
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdNästa 
      Caption         =   "Nästa Spelare"
      Height          =   375
      Left            =   720
      TabIndex        =   129
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtUndre5 
      Height          =   285
      Left            =   7800
      TabIndex        =   128
      Top             =   6120
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.TextBox txtUndre4 
      Height          =   285
      Left            =   6960
      TabIndex        =   127
      Top             =   6120
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.TextBox txtUndre3 
      Height          =   285
      Left            =   6120
      TabIndex        =   126
      Top             =   6120
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.TextBox txtUndre2 
      Height          =   285
      Left            =   5280
      TabIndex        =   125
      Top             =   6120
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.TextBox txtUndre1 
      Height          =   285
      Left            =   4440
      TabIndex        =   124
      Top             =   6120
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.TextBox txtÖvre5 
      Height          =   285
      Left            =   7800
      TabIndex        =   123
      Top             =   5880
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.TextBox txtÖvre4 
      Height          =   285
      Left            =   6960
      TabIndex        =   122
      Top             =   5880
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.TextBox txtÖvre3 
      Height          =   285
      Left            =   6120
      TabIndex        =   121
      Top             =   5880
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.TextBox txtÖvre2 
      Height          =   285
      Left            =   5280
      TabIndex        =   120
      Top             =   5880
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.TextBox txtÖvre1 
      Height          =   285
      Left            =   4440
      TabIndex        =   119
      Top             =   5880
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmd2Summa5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   118
      Top             =   4440
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdYatzy5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   117
      Top             =   4200
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdChans5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   116
      Top             =   3960
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdKåk5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   115
      Top             =   3720
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdStor5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   114
      Top             =   3480
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdLiten5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   113
      Top             =   3240
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdFyrtal5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   112
      Top             =   3000
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdTretal5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   111
      Top             =   2760
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmd2Par5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   110
      Top             =   2520
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmd1Par5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   109
      Top             =   2280
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdBonus5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   108
      Top             =   2040
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSumma5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   107
      Top             =   1800
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSexor5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   106
      Top             =   1560
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdFemmor5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   105
      Top             =   1320
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdFyror5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   104
      Top             =   1080
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdTreor5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   103
      Top             =   840
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdTvåor5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   102
      Top             =   600
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdEttor5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   101
      Top             =   360
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmd2Summa4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   100
      Top             =   4440
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdYatzy4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   99
      Top             =   4200
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdChans4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   98
      Top             =   3960
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdKåk4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   97
      Top             =   3720
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdStor4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   96
      Top             =   3480
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdLiten4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   95
      Top             =   3240
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdFyrtal4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   94
      Top             =   3000
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdTretal4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   93
      Top             =   2760
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmd2Par4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   92
      Top             =   2520
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmd1Par4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   91
      Top             =   2280
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdBonus4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   90
      Top             =   2040
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSumma4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   89
      Top             =   1800
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSexor4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   88
      Top             =   1560
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdFemmor4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   87
      Top             =   1320
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdFyror4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   86
      Top             =   1080
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdTreor4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   85
      Top             =   840
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdTvåor4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   84
      Top             =   600
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdEttor4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   83
      Top             =   360
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmd2Summa3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   82
      Top             =   4440
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdYatzy3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   81
      Top             =   4200
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdChans3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   80
      Top             =   3960
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdKåk3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   79
      Top             =   3720
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdStor3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   78
      Top             =   3480
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdLiten3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   77
      Top             =   3240
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdFyrtal3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   76
      Top             =   3000
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdTretal3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   75
      Top             =   2760
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmd2Par3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   74
      Top             =   2520
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmd1Par3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   73
      Top             =   2280
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdBonus3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   72
      Top             =   2040
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSumma3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   71
      Top             =   1800
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSexor3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   70
      Top             =   1560
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdFemmor3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   69
      Top             =   1320
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdFyror3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   68
      Top             =   1080
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdTreor3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   67
      Top             =   840
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdTvåor3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   66
      Top             =   600
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdEttor3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   65
      Top             =   360
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmd2Summa2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   64
      Top             =   4440
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdYatzy2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   63
      Top             =   4200
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdChans2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   62
      Top             =   3960
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdKåk2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   61
      Top             =   3720
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdStor2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   60
      Top             =   3480
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdLiten2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   59
      Top             =   3240
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdFyrtal2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   58
      Top             =   3000
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdTretal2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   57
      Top             =   2760
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmd2Par2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   56
      Top             =   2520
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmd1Par2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   55
      Top             =   2280
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdBonus2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   54
      Top             =   2040
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSumma2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   53
      Top             =   1800
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSexor2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   52
      Top             =   1560
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdFemmor2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   51
      Top             =   1320
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdFyror2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   50
      Top             =   1080
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdTreor2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   49
      Top             =   840
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdTvåor2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   48
      Top             =   600
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdEttor2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   47
      Top             =   360
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdKasta 
      BackColor       =   &H00000000&
      Caption         =   "Kasta"
      Height          =   375
      Left            =   840
      TabIndex        =   46
      Top             =   2520
      Width           =   975
   End
   Begin VB.CheckBox chk5 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2280
      TabIndex        =   45
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox chk4 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1800
      TabIndex        =   44
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox chk3 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1320
      TabIndex        =   43
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox chk2 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   840
      TabIndex        =   42
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox chk1 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   41
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton cmd2Summa1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   35
      Top             =   4440
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdYatzy1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   34
      ToolTipText     =   "5 lika tärningar"
      Top             =   4200
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdChans1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   33
      ToolTipText     =   "Vad som hellst"
      Top             =   3960
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdKåk1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   32
      ToolTipText     =   "3 lika och 2 lika"
      Top             =   3720
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdStor1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   31
      ToolTipText     =   "Tärningarna visar  2, 3, 4, 5, 6"
      Top             =   3480
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdLiten1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   30
      ToolTipText     =   "Tärningarna visar 1, 2, 3, 4, 5"
      Top             =   3240
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdFyrtal1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   29
      ToolTipText     =   "4 lika tärningar"
      Top             =   3000
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdTretal1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   28
      ToolTipText     =   "3 lika tärningar"
      Top             =   2760
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmd2Par1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   27
      ToolTipText     =   "2 lika plus ytterligare 2 lika"
      Top             =   2520
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmd1Par1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   26
      ToolTipText     =   "2 tärningar som är lika"
      Top             =   2280
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdBonus1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   25
      Top             =   2040
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSumma1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   24
      Top             =   1800
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSexor1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   23
      ToolTipText     =   "Tre sexor behövs för att få ett X"
      Top             =   1560
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdFemmor1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   22
      ToolTipText     =   "Tre femmor behövs för att få ett X"
      Top             =   1320
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdFyror1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   21
      ToolTipText     =   "Tre fyror behövs för att få ett X"
      Top             =   1080
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdTreor1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      ToolTipText     =   "Tre treor behövs för att få ett X"
      Top             =   840
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdTvåor1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   19
      ToolTipText     =   "Tre tvåor behövs för att få ett X"
      Top             =   600
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdEttor1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   18
      ToolTipText     =   "Tre ettor behövs för att få ett X"
      Top             =   360
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.Line Line20 
      X1              =   4440
      X2              =   4440
      Y1              =   360
      Y2              =   4680
   End
   Begin VB.Label lblSpelare5 
      Height          =   255
      Left            =   7800
      TabIndex        =   135
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblSpelare4 
      Height          =   255
      Left            =   6960
      TabIndex        =   134
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblSpelare3 
      Height          =   255
      Left            =   6120
      TabIndex        =   133
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblSpelare2 
      Height          =   255
      Left            =   5280
      TabIndex        =   132
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblSpelare1 
      Height          =   255
      Left            =   4440
      TabIndex        =   131
      Top             =   120
      Width           =   855
   End
   Begin VB.Image imgVisa2 
      Height          =   375
      Index           =   4
      Left            =   1200
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgVisa2 
      Height          =   375
      Index           =   3
      Left            =   240
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgVisa2 
      Height          =   375
      Index           =   2
      Left            =   720
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgVisa2 
      Height          =   375
      Index           =   0
      Left            =   1680
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgVisa2 
      Height          =   375
      Index           =   1
      Left            =   2160
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgVisa 
      Height          =   375
      Index           =   1
      Left            =   2160
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgVisa 
      Height          =   375
      Index           =   2
      Left            =   720
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgVisa 
      Height          =   375
      Index           =   3
      Left            =   240
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgVisa 
      Height          =   375
      Index           =   4
      Left            =   1200
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgTär 
      Height          =   300
      Index           =   3
      Left            =   720
      Picture         =   "Yatzy.frx":0FE4
      Top             =   0
      Width           =   300
      Visible         =   0   'False
   End
   Begin VB.Image imgTär 
      Height          =   300
      Index           =   5
      Left            =   1080
      Picture         =   "Yatzy.frx":105F
      Top             =   0
      Width           =   300
      Visible         =   0   'False
   End
   Begin VB.Image imgTär 
      Height          =   300
      Index           =   2
      Left            =   1440
      Picture         =   "Yatzy.frx":10E0
      Top             =   0
      Width           =   300
      Visible         =   0   'False
   End
   Begin VB.Image imgTär 
      Height          =   300
      Index           =   4
      Left            =   1800
      Picture         =   "Yatzy.frx":1165
      Top             =   0
      Width           =   300
      Visible         =   0   'False
   End
   Begin VB.Image imgVisa 
      Height          =   375
      Index           =   0
      Left            =   1680
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgTär 
      Height          =   300
      Index           =   0
      Left            =   0
      Picture         =   "Yatzy.frx":11EF
      Top             =   0
      Width           =   300
      Visible         =   0   'False
   End
   Begin VB.Image imgTär 
      Height          =   300
      Index           =   1
      Left            =   360
      Picture         =   "Yatzy.frx":1260
      Top             =   0
      Width           =   300
      Visible         =   0   'False
   End
   Begin VB.Line Line19 
      X1              =   3360
      X2              =   4440
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line18 
      X1              =   3360
      X2              =   4440
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line17 
      X1              =   3360
      X2              =   4440
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line16 
      X1              =   3360
      X2              =   4440
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line15 
      X1              =   3360
      X2              =   4440
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line14 
      X1              =   3360
      X2              =   4440
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line13 
      X1              =   3360
      X2              =   4440
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line12 
      X1              =   3360
      X2              =   4440
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line11 
      X1              =   3360
      X2              =   4440
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line10 
      X1              =   3360
      X2              =   4440
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line9 
      X1              =   3360
      X2              =   4440
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line8 
      X1              =   3360
      X2              =   4440
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line7 
      X1              =   3360
      X2              =   4440
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line6 
      X1              =   3360
      X2              =   4440
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line5 
      X1              =   3360
      X2              =   4440
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line4 
      X1              =   3360
      X2              =   3360
      Y1              =   360
      Y2              =   4680
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   3360
      X2              =   4440
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   3360
      X2              =   4440
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      X1              =   3360
      X2              =   4440
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line1 
      X1              =   3360
      X2              =   4440
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label lblTärning5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   40
      Top             =   480
      Width           =   255
      Visible         =   0   'False
   End
   Begin VB.Label lblTärning2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   39
      Top             =   480
      Width           =   255
      Visible         =   0   'False
   End
   Begin VB.Label lblTärning1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   480
      Width           =   255
      Visible         =   0   'False
   End
   Begin VB.Label lblTärning4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   37
      Top             =   480
      Width           =   255
      Visible         =   0   'False
   End
   Begin VB.Label lblTärning3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   36
      Top             =   480
      Width           =   255
      Visible         =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Ettor"
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label18 
      Caption         =   "Summa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label17 
      Caption         =   "YATZY"
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label16 
      Caption         =   "Chans"
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label15 
      Caption         =   "Kåk"
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Stor Stege"
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Liten Stege"
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Fyrtal"
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Tretal"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "1 par"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Bonus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Summa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Sexor"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Femmor"
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Fyror"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Treor"
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Tvåor"
      Height          =   255
      Left            =   3480
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "2 par"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.Menu mnuArkiv 
      Caption         =   "&Arkiv"
      Begin VB.Menu mnuNytt 
         Caption         =   "&Nytt Spel"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Skriv ut resultat"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScore 
         Caption         =   "High Score"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAvsluta 
         Caption         =   "&Avsluta"
      End
   End
   Begin VB.Menu mnuSpelare 
      Caption         =   "Spelare"
      Begin VB.Menu mnuEn 
         Caption         =   "1 Spelare"
      End
      Begin VB.Menu mnutvå 
         Caption         =   "2 Spelare"
      End
      Begin VB.Menu mnuTre 
         Caption         =   "3 Spelare"
      End
      Begin VB.Menu mnuFyra 
         Caption         =   "4 Spelare"
      End
      Begin VB.Menu mnuFem 
         Caption         =   "5 Spelare"
      End
   End
   Begin VB.Menu mnuHjälp 
      Caption         =   "Hjälp"
      Begin VB.Menu mnuOm 
         Caption         =   "Om"
      End
   End
End
Attribute VB_Name = "frmYatzy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Kast As Integer
Dim Counter As Integer
Dim Response As Integer
Dim Poäng1 As Integer
Dim Poäng2 As Integer

Dim Poäng As Integer
Dim Poäng3 As Integer
Dim Poäng4 As Integer
Dim Poäng5 As Integer
Dim Poäng6 As Integer

Dim Tärning1 As Integer
Dim Tärning2 As Integer
Dim Tärning3 As Integer
Dim Tärning4 As Integer
Dim Tärning5 As Integer
Dim Tärning As Integer
Dim Current As Integer
Dim Visa As Integer
Dim Animering As Integer
Dim Fusk As Integer
Dim Spelare As Integer
Dim SpelareNr As Integer

Dim gScore As Score
Dim gFileNum As Integer
Dim gRecordLen As Long
Dim gCurrentRecord As Long
Dim gLastRecord As Long

Private Sub cmd1Par1_Click() 'Spelare 1, ett par
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = 0
    If Tärning1 = Tärning2 And (Tärning < Tärning1 + Tärning2) Then
        Tärning = Tärning1 + Tärning2
    End If
    
    If Tärning1 = Tärning3 And (Tärning < Tärning1 + Tärning3) Then
        Tärning = Tärning1 + Tärning3
    End If
    
    If Tärning1 = Tärning4 And (Tärning < Tärning1 + Tärning4) Then
        Tärning = Tärning1 + Tärning4
    End If

    If Tärning1 = Tärning5 And (Tärning < Tärning1 + Tärning5) Then
        Tärning = Tärning1 + Tärning5
    End If
    
    
    If Tärning2 = Tärning1 And (Tärning < Tärning2 + Tärning1) Then
        Tärning = Tärning2 + Tärning1
    End If
    
    If Tärning2 = Tärning3 And (Tärning < Tärning2 + Tärning3) Then
        Tärning = Tärning2 + Tärning3
    End If
    
    If Tärning2 = Tärning4 And (Tärning < Tärning2 + Tärning4) Then
        Tärning = Tärning2 + Tärning4
    End If

    If Tärning2 = Tärning5 And (Tärning < Tärning2 + Tärning5) Then
        Tärning = Tärning2 + Tärning5
    End If
    
    If Tärning3 = Tärning2 And (Tärning < Tärning3 + Tärning2) Then
        Tärning = Tärning3 + Tärning2
    End If
    
    If Tärning3 = Tärning1 And (Tärning < Tärning1 + Tärning3) Then
        Tärning = Tärning1 + Tärning3
    End If
    
    If Tärning3 = Tärning4 And (Tärning < Tärning3 + Tärning4) Then
        Tärning = Tärning3 + Tärning4
    End If

    If Tärning3 = Tärning5 And (Tärning < Tärning3 + Tärning5) Then
        Tärning = Tärning3 + Tärning5
    End If
    
    If Tärning4 = Tärning2 And (Tärning < Tärning4 + Tärning2) Then
        Tärning = Tärning4 + Tärning2
    End If
    
    If Tärning4 = Tärning3 And (Tärning < Tärning4 + Tärning3) Then
        Tärning = Tärning4 + Tärning3
    End If
    
    If Tärning4 = Tärning1 And (Tärning < Tärning1 + Tärning4) Then
        Tärning = Tärning1 + Tärning4
    End If

    If Tärning4 = Tärning5 And (Tärning < Tärning4 + Tärning5) Then
        Tärning = Tärning4 + Tärning5
    End If
    
    If Tärning5 = Tärning2 And (Tärning < Tärning5 + Tärning2) Then
        Tärning = Tärning5 + Tärning2
    End If
    
    If Tärning5 = Tärning3 And (Tärning < Tärning5 + Tärning3) Then
        Tärning = Tärning5 + Tärning3
    End If
    
    If Tärning5 = Tärning4 And (Tärning < Tärning5 + Tärning4) Then
        Tärning = Tärning5 + Tärning4
    End If

    If Tärning5 = Tärning1 And (Tärning < Tärning1 + Tärning5) Then
        Tärning = Tärning1 + Tärning5
    End If
    If Tärning > 0 Then
        cmd1Par1.Caption = Tärning
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        txtUndre1.Text = txtUndre1.Text + Tärning
        cmd1Par1.Enabled = False
        Exit Sub
    End If
    If Tärning = 0 Then
        Response = MsgBox("Du har inget par! Vill du strycka ett par?", vbYesNo, "Medelande")
            If Response = vbYes Then
                cmd1Par1.Caption = "-"
                cmdNästa.Enabled = True
                cmdKasta.Enabled = False
                cmd1Par1.Enabled = False
            End If
            If Response = vbNo Then
                Fusk = 1
            End If
        Exit Sub
    End If
    End If
End Sub

Private Sub cmd1Par2_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = 0
    If Tärning1 = Tärning2 And (Tärning < Tärning1 + Tärning2) Then
        Tärning = Tärning1 + Tärning2
    End If
    If Tärning1 = Tärning3 And (Tärning < Tärning1 + Tärning3) Then
        Tärning = Tärning1 + Tärning3
    End If
    If Tärning1 = Tärning4 And (Tärning < Tärning1 + Tärning4) Then
        Tärning = Tärning1 + Tärning4
    End If
    If Tärning1 = Tärning5 And (Tärning < Tärning1 + Tärning5) Then
        Tärning = Tärning1 + Tärning5
    End If
    If Tärning2 = Tärning1 And (Tärning < Tärning2 + Tärning1) Then
        Tärning = Tärning2 + Tärning1
    End If
    If Tärning2 = Tärning3 And (Tärning < Tärning2 + Tärning3) Then
        Tärning = Tärning2 + Tärning3
    End If
    If Tärning2 = Tärning4 And (Tärning < Tärning2 + Tärning4) Then
        Tärning = Tärning2 + Tärning4
    End If
    If Tärning2 = Tärning5 And (Tärning < Tärning2 + Tärning5) Then
        Tärning = Tärning2 + Tärning5
    End If
    If Tärning3 = Tärning2 And (Tärning < Tärning3 + Tärning2) Then
        Tärning = Tärning3 + Tärning2
    End If
    If Tärning3 = Tärning1 And (Tärning < Tärning1 + Tärning3) Then
        Tärning = Tärning1 + Tärning3
    End If
    If Tärning3 = Tärning4 And (Tärning < Tärning3 + Tärning4) Then
        Tärning = Tärning3 + Tärning4
    End If
    If Tärning3 = Tärning5 And (Tärning < Tärning3 + Tärning5) Then
        Tärning = Tärning3 + Tärning5
    End If
    If Tärning4 = Tärning2 And (Tärning < Tärning4 + Tärning2) Then
        Tärning = Tärning4 + Tärning2
    End If
    If Tärning4 = Tärning3 And (Tärning < Tärning4 + Tärning3) Then
        Tärning = Tärning4 + Tärning3
    End If
    If Tärning4 = Tärning1 And (Tärning < Tärning1 + Tärning4) Then
        Tärning = Tärning1 + Tärning4
    End If
    If Tärning4 = Tärning5 And (Tärning < Tärning4 + Tärning5) Then
        Tärning = Tärning4 + Tärning5
    End If
    If Tärning5 = Tärning2 And (Tärning < Tärning5 + Tärning2) Then
        Tärning = Tärning5 + Tärning2
    End If
    If Tärning5 = Tärning3 And (Tärning < Tärning5 + Tärning3) Then
        Tärning = Tärning5 + Tärning3
    End If
    If Tärning5 = Tärning4 And (Tärning < Tärning5 + Tärning4) Then
        Tärning = Tärning5 + Tärning4
    End If
    If Tärning5 = Tärning1 And (Tärning < Tärning1 + Tärning5) Then
        Tärning = Tärning1 + Tärning5
    End If
    If Tärning > 0 Then
        cmd1Par2.Caption = Tärning
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        txtUndre2.Text = txtUndre2.Text + Tärning
        cmd1Par2.Enabled = False
        Exit Sub
    End If
    If Tärning = 0 Then
        Response = MsgBox("Du har inget par! Vill du strycka ett par?", vbYesNo, "Medelande")
            If Response = vbYes Then
                cmd1Par2.Caption = "-"
                cmdNästa.Enabled = True
                cmdKasta.Enabled = False
                cmd1Par2.Enabled = False
            End If
            If Response = vbNo Then
                Fusk = 1
            End If
        Exit Sub
    End If
    End If
End Sub

Private Sub cmd1Par3_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = 0
    If Tärning1 = Tärning2 And (Tärning < Tärning1 + Tärning2) Then
        Tärning = Tärning1 + Tärning2
    End If
    
    If Tärning1 = Tärning3 And (Tärning < Tärning1 + Tärning3) Then
        Tärning = Tärning1 + Tärning3
    End If
    
    If Tärning1 = Tärning4 And (Tärning < Tärning1 + Tärning4) Then
        Tärning = Tärning1 + Tärning4
    End If

    If Tärning1 = Tärning5 And (Tärning < Tärning1 + Tärning5) Then
        Tärning = Tärning1 + Tärning5
    End If
    
    
    If Tärning2 = Tärning1 And (Tärning < Tärning2 + Tärning1) Then
        Tärning = Tärning2 + Tärning1
    End If
    
    If Tärning2 = Tärning3 And (Tärning < Tärning2 + Tärning3) Then
        Tärning = Tärning2 + Tärning3
    End If
    
    If Tärning2 = Tärning4 And (Tärning < Tärning2 + Tärning4) Then
        Tärning = Tärning2 + Tärning4
    End If

    If Tärning2 = Tärning5 And (Tärning < Tärning2 + Tärning5) Then
        Tärning = Tärning2 + Tärning5
    End If
    
    If Tärning3 = Tärning2 And (Tärning < Tärning3 + Tärning2) Then
        Tärning = Tärning3 + Tärning2
    End If
    
    If Tärning3 = Tärning1 And (Tärning < Tärning1 + Tärning3) Then
        Tärning = Tärning1 + Tärning3
    End If
    
    If Tärning3 = Tärning4 And (Tärning < Tärning3 + Tärning4) Then
        Tärning = Tärning3 + Tärning4
    End If

    If Tärning3 = Tärning5 And (Tärning < Tärning3 + Tärning5) Then
        Tärning = Tärning3 + Tärning5
    End If
    
    If Tärning4 = Tärning2 And (Tärning < Tärning4 + Tärning2) Then
        Tärning = Tärning4 + Tärning2
    End If
    
    If Tärning4 = Tärning3 And (Tärning < Tärning4 + Tärning3) Then
        Tärning = Tärning4 + Tärning3
    End If
    
    If Tärning4 = Tärning1 And (Tärning < Tärning1 + Tärning4) Then
        Tärning = Tärning1 + Tärning4
    End If

    If Tärning4 = Tärning5 And (Tärning < Tärning4 + Tärning5) Then
        Tärning = Tärning4 + Tärning5
    End If
    
    If Tärning5 = Tärning2 And (Tärning < Tärning5 + Tärning2) Then
        Tärning = Tärning5 + Tärning2
    End If
    
    If Tärning5 = Tärning3 And (Tärning < Tärning5 + Tärning3) Then
        Tärning = Tärning5 + Tärning3
    End If
    
    If Tärning5 = Tärning4 And (Tärning < Tärning5 + Tärning4) Then
        Tärning = Tärning5 + Tärning4
    End If

    If Tärning5 = Tärning1 And (Tärning < Tärning1 + Tärning5) Then
        Tärning = Tärning1 + Tärning5
    End If
    If Tärning > 0 Then
        cmd1Par3.Caption = Tärning
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        txtUndre3.Text = txtUndre3.Text + Tärning
        cmd1Par3.Enabled = False
        Exit Sub
    End If
    If Tärning = 0 Then
        Response = MsgBox("Du har inget par! Vill du strycka ett par?", vbYesNo, "Medelande")
            If Response = vbYes Then
                cmd1Par3.Caption = "-"
                cmdNästa.Enabled = True
                cmdKasta.Enabled = False
                cmd1Par3.Enabled = False
            End If
            If Response = vbNo Then
                Fusk = 1
            End If
        Exit Sub
    End If
    End If
End Sub

Private Sub cmd1Par4_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = 0
    If Tärning1 = Tärning2 And (Tärning < Tärning1 + Tärning2) Then
        Tärning = Tärning1 + Tärning2
    End If
    
    If Tärning1 = Tärning3 And (Tärning < Tärning1 + Tärning3) Then
        Tärning = Tärning1 + Tärning3
    End If
    
    If Tärning1 = Tärning4 And (Tärning < Tärning1 + Tärning4) Then
        Tärning = Tärning1 + Tärning4
    End If

    If Tärning1 = Tärning5 And (Tärning < Tärning1 + Tärning5) Then
        Tärning = Tärning1 + Tärning5
    End If
    
    
    If Tärning2 = Tärning1 And (Tärning < Tärning2 + Tärning1) Then
        Tärning = Tärning2 + Tärning1
    End If
    
    If Tärning2 = Tärning3 And (Tärning < Tärning2 + Tärning3) Then
        Tärning = Tärning2 + Tärning3
    End If
    
    If Tärning2 = Tärning4 And (Tärning < Tärning2 + Tärning4) Then
        Tärning = Tärning2 + Tärning4
    End If

    If Tärning2 = Tärning5 And (Tärning < Tärning2 + Tärning5) Then
        Tärning = Tärning2 + Tärning5
    End If
    
    If Tärning3 = Tärning2 And (Tärning < Tärning3 + Tärning2) Then
        Tärning = Tärning3 + Tärning2
    End If
    
    If Tärning3 = Tärning1 And (Tärning < Tärning1 + Tärning3) Then
        Tärning = Tärning1 + Tärning3
    End If
    
    If Tärning3 = Tärning4 And (Tärning < Tärning3 + Tärning4) Then
        Tärning = Tärning3 + Tärning4
    End If

    If Tärning3 = Tärning5 And (Tärning < Tärning3 + Tärning5) Then
        Tärning = Tärning3 + Tärning5
    End If
    
    If Tärning4 = Tärning2 And (Tärning < Tärning4 + Tärning2) Then
        Tärning = Tärning4 + Tärning2
    End If
    
    If Tärning4 = Tärning3 And (Tärning < Tärning4 + Tärning3) Then
        Tärning = Tärning4 + Tärning3
    End If
    
    If Tärning4 = Tärning1 And (Tärning < Tärning1 + Tärning4) Then
        Tärning = Tärning1 + Tärning4
    End If

    If Tärning4 = Tärning5 And (Tärning < Tärning4 + Tärning5) Then
        Tärning = Tärning4 + Tärning5
    End If
    
    If Tärning5 = Tärning2 And (Tärning < Tärning5 + Tärning2) Then
        Tärning = Tärning5 + Tärning2
    End If
    
    If Tärning5 = Tärning3 And (Tärning < Tärning5 + Tärning3) Then
        Tärning = Tärning5 + Tärning3
    End If
    
    If Tärning5 = Tärning4 And (Tärning < Tärning5 + Tärning4) Then
        Tärning = Tärning5 + Tärning4
    End If

    If Tärning5 = Tärning1 And (Tärning < Tärning1 + Tärning5) Then
        Tärning = Tärning1 + Tärning5
    End If
    If Tärning > 0 Then
        cmd1Par4.Caption = Tärning
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        txtUndre4.Text = txtUndre4.Text + Tärning
        cmd1Par4.Enabled = False
        Exit Sub
    End If
    If Tärning = 0 Then
        Response = MsgBox("Du har inget par! Vill du strycka ett par?", vbYesNo, "Medelande")
            If Response = vbYes Then
                cmd1Par4.Caption = "-"
                cmdNästa.Enabled = True
                cmdKasta.Enabled = False
                cmd1Par4.Enabled = False
            End If
            If Response = vbNo Then
                Fusk = 1
            End If
        Exit Sub
    End If
    End If
End Sub

Private Sub cmd1Par5_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = 0
    If Tärning1 = Tärning2 And (Tärning < Tärning1 + Tärning2) Then
        Tärning = Tärning1 + Tärning2
    End If
    
    If Tärning1 = Tärning3 And (Tärning < Tärning1 + Tärning3) Then
        Tärning = Tärning1 + Tärning3
    End If
    
    If Tärning1 = Tärning4 And (Tärning < Tärning1 + Tärning4) Then
        Tärning = Tärning1 + Tärning4
    End If

    If Tärning1 = Tärning5 And (Tärning < Tärning1 + Tärning5) Then
        Tärning = Tärning1 + Tärning5
    End If
    
    
    If Tärning2 = Tärning1 And (Tärning < Tärning2 + Tärning1) Then
        Tärning = Tärning2 + Tärning1
    End If
    
    If Tärning2 = Tärning3 And (Tärning < Tärning2 + Tärning3) Then
        Tärning = Tärning2 + Tärning3
    End If
    
    If Tärning2 = Tärning4 And (Tärning < Tärning2 + Tärning4) Then
        Tärning = Tärning2 + Tärning4
    End If

    If Tärning2 = Tärning5 And (Tärning < Tärning2 + Tärning5) Then
        Tärning = Tärning2 + Tärning5
    End If
    
    If Tärning3 = Tärning2 And (Tärning < Tärning3 + Tärning2) Then
        Tärning = Tärning3 + Tärning2
    End If
    
    If Tärning3 = Tärning1 And (Tärning < Tärning1 + Tärning3) Then
        Tärning = Tärning1 + Tärning3
    End If
    
    If Tärning3 = Tärning4 And (Tärning < Tärning3 + Tärning4) Then
        Tärning = Tärning3 + Tärning4
    End If

    If Tärning3 = Tärning5 And (Tärning < Tärning3 + Tärning5) Then
        Tärning = Tärning3 + Tärning5
    End If
    
    If Tärning4 = Tärning2 And (Tärning < Tärning4 + Tärning2) Then
        Tärning = Tärning4 + Tärning2
    End If
    
    If Tärning4 = Tärning3 And (Tärning < Tärning4 + Tärning3) Then
        Tärning = Tärning4 + Tärning3
    End If
    
    If Tärning4 = Tärning1 And (Tärning < Tärning1 + Tärning4) Then
        Tärning = Tärning1 + Tärning4
    End If

    If Tärning4 = Tärning5 And (Tärning < Tärning4 + Tärning5) Then
        Tärning = Tärning4 + Tärning5
    End If
    
    If Tärning5 = Tärning2 And (Tärning < Tärning5 + Tärning2) Then
        Tärning = Tärning5 + Tärning2
    End If
    
    If Tärning5 = Tärning3 And (Tärning < Tärning5 + Tärning3) Then
        Tärning = Tärning5 + Tärning3
    End If
    
    If Tärning5 = Tärning4 And (Tärning < Tärning5 + Tärning4) Then
        Tärning = Tärning5 + Tärning4
    End If

    If Tärning5 = Tärning1 And (Tärning < Tärning1 + Tärning5) Then
        Tärning = Tärning1 + Tärning5
    End If
    If Tärning > 0 Then
        cmd1Par5.Caption = Tärning
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        txtUndre5.Text = txtUndre5.Text + Tärning
        cmd1Par5.Enabled = False
        Exit Sub
    End If
    If Tärning = 0 Then
        Response = MsgBox("Du har inget par! Vill du strycka ett par?", vbYesNo, "Medelande")
            If Response = vbYes Then
                cmd1Par5.Caption = "-"
                cmdNästa.Enabled = True
                cmdKasta.Enabled = False
                cmd1Par5.Enabled = False
            End If
            If Response = vbNo Then
                Fusk = 1
            End If
        Exit Sub
    End If
    End If
End Sub

Private Sub cmd2Par1_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = 0
 
    If Tärning1 = Tärning2 Then
        If Tärning3 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare1
            Exit Sub
        End If
        If Tärning3 = Tärning5 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning5
            ParSpelare1
            Exit Sub
        End If
        If Tärning4 = Tärning5 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
            ParSpelare1
            Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning4 Then
        If Tärning1 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare1
            Exit Sub
        End If
        If Tärning2 = Tärning5 Then
            Tärning = Tärning4 + Tärning2 + Tärning3 + Tärning5
            ParSpelare1
            Exit Sub
        End If
        If Tärning1 = Tärning5 Then
            Tärning = Tärning1 + Tärning3 + Tärning5 + Tärning4
            ParSpelare1
            Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning5 Then
        If Tärning2 = Tärning4 Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
            ParSpelare1
            Exit Sub
        End If
        If Tärning4 = Tärning1 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare1
            Exit Sub
        End If
        If Tärning1 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare1
            Exit Sub
        End If
    End If

    If Tärning4 = Tärning5 Then
        If Tärning2 = Tärning3 Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
            ParSpelare1
            Exit Sub
        End If
        If Tärning2 = Tärning1 Then
            Tärning = Tärning1 + Tärning4 + Tärning2 + Tärning5
            ParSpelare1
            Exit Sub
        End If
        If Tärning1 = Tärning3 Then
            Tärning = Tärning1 + Tärning3 + Tärning5 + Tärning4
            ParSpelare1
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning1 Then
        If Tärning2 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare1
            Exit Sub
        End If
        If Tärning4 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare1
            Exit Sub
        End If
        If Tärning5 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare1
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning4 Then
        If Tärning2 = Tärning5 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
            ParSpelare1
            Exit Sub
        End If
        If Tärning3 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare1
            Exit Sub
        End If
        If Tärning3 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning4 + Tärning3
            ParSpelare1
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning5 Then
        If Tärning2 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
            ParSpelare1
            Exit Sub
        End If
        If Tärning4 = Tärning3 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare1
            Exit Sub
        End If
        If Tärning3 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare1
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning2 Then
        If Tärning1 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare1
            Exit Sub
        End If
        If Tärning4 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare1
            Exit Sub
        End If
        If Tärning5 = Tärning1 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare1
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning4 Then
        If Tärning1 = Tärning3 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare1
            Exit Sub
        End If
        If Tärning1 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning2 + Tärning5
            ParSpelare1
            Exit Sub
        End If
        If Tärning3 = Tärning5 Then
            Tärning = Tärning4 + Tärning2 + Tärning5 + Tärning3
            ParSpelare1
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning5 Then
        If Tärning3 = Tärning4 Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
            ParSpelare1
            Exit Sub
        End If
        If Tärning4 = Tärning1 Then
            Tärning = Tärning1 + Tärning4 + Tärning2 + Tärning5
            ParSpelare1
            Exit Sub
        End If
        If Tärning1 = Tärning3 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare1
            Exit Sub
        End If
    End If
    Response = MsgBox("Du har inte två par! Vill du strycka två par?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmd2Par1.Enabled = False
        cmd2Par1.Caption = "-"
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmd2Par2_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = 0
 
    If Tärning1 = Tärning2 Then
        If Tärning3 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare2
            Exit Sub
        End If
        If Tärning3 = Tärning5 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning5
            ParSpelare2
            Exit Sub
        End If
        If Tärning4 = Tärning5 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
            ParSpelare2
            Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning4 Then
        If Tärning1 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare2
            Exit Sub
        End If
        If Tärning2 = Tärning5 Then
            Tärning = Tärning4 + Tärning2 + Tärning3 + Tärning5
            ParSpelare2
            Exit Sub
        End If
        If Tärning1 = Tärning5 Then
            Tärning = Tärning1 + Tärning3 + Tärning5 + Tärning4
            ParSpelare2
            Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning5 Then
        If Tärning2 = Tärning4 Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
            ParSpelare2
            Exit Sub
        End If
        If Tärning4 = Tärning1 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare2
            Exit Sub
        End If
        If Tärning1 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare2
            Exit Sub
        End If
    End If

    If Tärning4 = Tärning5 Then
        If Tärning2 = Tärning3 Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
            ParSpelare2
            Exit Sub
        End If
        If Tärning2 = Tärning1 Then
            Tärning = Tärning1 + Tärning4 + Tärning2 + Tärning5
            ParSpelare2
            Exit Sub
        End If
        If Tärning1 = Tärning3 Then
            Tärning = Tärning1 + Tärning3 + Tärning5 + Tärning4
            ParSpelare2
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning1 Then
        If Tärning2 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare2
            Exit Sub
        End If
        If Tärning4 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare2
            Exit Sub
        End If
        If Tärning5 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare2
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning4 Then
        If Tärning2 = Tärning5 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
            ParSpelare2
            Exit Sub
        End If
        If Tärning3 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare2
            Exit Sub
        End If
        If Tärning3 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning4 + Tärning3
            ParSpelare2
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning5 Then
        If Tärning2 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
            ParSpelare2
            Exit Sub
        End If
        If Tärning4 = Tärning3 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare2
            Exit Sub
        End If
        If Tärning3 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare2
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning2 Then
        If Tärning1 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare2
            Exit Sub
        End If
        If Tärning4 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare2
            Exit Sub
        End If
        If Tärning5 = Tärning1 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare2
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning4 Then
        If Tärning1 = Tärning3 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare2
            Exit Sub
        End If
        If Tärning1 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning2 + Tärning5
            ParSpelare2
            Exit Sub
        End If
        If Tärning3 = Tärning5 Then
            Tärning = Tärning4 + Tärning2 + Tärning5 + Tärning3
            ParSpelare2
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning5 Then
        If Tärning3 = Tärning4 Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
            ParSpelare2
            Exit Sub
        End If
        If Tärning4 = Tärning1 Then
            Tärning = Tärning1 + Tärning4 + Tärning2 + Tärning5
            ParSpelare2
            Exit Sub
        End If
        If Tärning1 = Tärning3 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare2
            Exit Sub
        End If
    End If
    Response = MsgBox("Du har inte två par! Vill du strycka två par?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmd2Par2.Enabled = False
        cmd2Par2.Caption = "-"
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmd2Par3_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = 0
 
    If Tärning1 = Tärning2 Then
        If Tärning3 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare3
            Exit Sub
        End If
        If Tärning3 = Tärning5 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning5
            ParSpelare3
            Exit Sub
        End If
        If Tärning4 = Tärning5 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
            ParSpelare3
            Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning4 Then
        If Tärning1 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare3
            Exit Sub
        End If
        If Tärning2 = Tärning5 Then
            Tärning = Tärning4 + Tärning2 + Tärning3 + Tärning5
            ParSpelare3
            Exit Sub
        End If
        If Tärning1 = Tärning5 Then
            Tärning = Tärning1 + Tärning3 + Tärning5 + Tärning4
            ParSpelare3
            Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning5 Then
        If Tärning2 = Tärning4 Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
            ParSpelare3
            Exit Sub
        End If
        If Tärning4 = Tärning1 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare3
            Exit Sub
        End If
        If Tärning1 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare3
            Exit Sub
        End If
    End If

    If Tärning4 = Tärning5 Then
        If Tärning2 = Tärning3 Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
            ParSpelare3
            Exit Sub
        End If
        If Tärning2 = Tärning1 Then
            Tärning = Tärning1 + Tärning4 + Tärning2 + Tärning5
            ParSpelare3
            Exit Sub
        End If
        If Tärning1 = Tärning3 Then
            Tärning = Tärning1 + Tärning3 + Tärning5 + Tärning4
            ParSpelare3
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning1 Then
        If Tärning2 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare3
            Exit Sub
        End If
        If Tärning4 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare3
            Exit Sub
        End If
        If Tärning5 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare3
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning4 Then
        If Tärning2 = Tärning5 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
            ParSpelare3
            Exit Sub
        End If
        If Tärning3 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare3
            Exit Sub
        End If
        If Tärning3 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning4 + Tärning3
            ParSpelare3
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning5 Then
        If Tärning2 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
            ParSpelare3
            Exit Sub
        End If
        If Tärning4 = Tärning3 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare3
            Exit Sub
        End If
        If Tärning3 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare3
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning2 Then
        If Tärning1 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare3
            Exit Sub
        End If
        If Tärning4 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare3
            Exit Sub
        End If
        If Tärning5 = Tärning1 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare3
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning4 Then
        If Tärning1 = Tärning3 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare3
            Exit Sub
        End If
        If Tärning1 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning2 + Tärning5
            ParSpelare3
            Exit Sub
        End If
        If Tärning3 = Tärning5 Then
            Tärning = Tärning4 + Tärning2 + Tärning5 + Tärning3
            ParSpelare3
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning5 Then
        If Tärning3 = Tärning4 Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
            ParSpelare3
            Exit Sub
        End If
        If Tärning4 = Tärning1 Then
            Tärning = Tärning1 + Tärning4 + Tärning2 + Tärning5
            ParSpelare3
            Exit Sub
        End If
        If Tärning1 = Tärning3 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare3
            Exit Sub
        End If
    End If
    Response = MsgBox("Du har inte två par! Vill du strycka två par?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmd2Par3.Enabled = False
        cmd2Par3.Caption = "-"
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmd2Par4_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = 0
 
    If Tärning1 = Tärning2 Then
        If Tärning3 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare4
            Exit Sub
        End If
        If Tärning3 = Tärning5 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning5
            ParSpelare4
            Exit Sub
        End If
        If Tärning4 = Tärning5 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
            ParSpelare4
            Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning4 Then
        If Tärning1 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare4
            Exit Sub
        End If
        If Tärning2 = Tärning5 Then
            Tärning = Tärning4 + Tärning2 + Tärning3 + Tärning5
            ParSpelare4
            Exit Sub
        End If
        If Tärning1 = Tärning5 Then
            Tärning = Tärning1 + Tärning3 + Tärning5 + Tärning4
            ParSpelare4
            Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning5 Then
        If Tärning2 = Tärning4 Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
            ParSpelare4
            Exit Sub
        End If
        If Tärning4 = Tärning1 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare4
            Exit Sub
        End If
        If Tärning1 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare4
            Exit Sub
        End If
    End If

    If Tärning4 = Tärning5 Then
        If Tärning2 = Tärning3 Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
            ParSpelare4
            Exit Sub
        End If
        If Tärning2 = Tärning1 Then
            Tärning = Tärning1 + Tärning4 + Tärning2 + Tärning5
            ParSpelare4
            Exit Sub
        End If
        If Tärning1 = Tärning3 Then
            Tärning = Tärning1 + Tärning3 + Tärning5 + Tärning4
            ParSpelare4
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning1 Then
        If Tärning2 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare4
            Exit Sub
        End If
        If Tärning4 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare4
            Exit Sub
        End If
        If Tärning5 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare4
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning4 Then
        If Tärning2 = Tärning5 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
            ParSpelare4
            Exit Sub
        End If
        If Tärning3 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare4
            Exit Sub
        End If
        If Tärning3 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning4 + Tärning3
            ParSpelare4
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning5 Then
        If Tärning2 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
            ParSpelare4
            Exit Sub
        End If
        If Tärning4 = Tärning3 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare4
            Exit Sub
        End If
        If Tärning3 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare4
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning2 Then
        If Tärning1 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare4
            Exit Sub
        End If
        If Tärning4 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare4
            Exit Sub
        End If
        If Tärning5 = Tärning1 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare4
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning4 Then
        If Tärning1 = Tärning3 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare4
            Exit Sub
        End If
        If Tärning1 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning2 + Tärning5
            ParSpelare4
            Exit Sub
        End If
        If Tärning3 = Tärning5 Then
            Tärning = Tärning4 + Tärning2 + Tärning5 + Tärning3
            ParSpelare4
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning5 Then
        If Tärning3 = Tärning4 Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
            ParSpelare4
            Exit Sub
        End If
        If Tärning4 = Tärning1 Then
            Tärning = Tärning1 + Tärning4 + Tärning2 + Tärning5
            ParSpelare4
            Exit Sub
        End If
        If Tärning1 = Tärning3 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare4
            Exit Sub
        End If
    End If
    Response = MsgBox("Du har inte två par! Vill du strycka två par?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmd2Par4.Enabled = False
        cmd2Par4.Caption = "-"
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmd2Par5_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = 0
 
    If Tärning1 = Tärning2 Then
        If Tärning3 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare5
            Exit Sub
        End If
        If Tärning3 = Tärning5 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning5
            ParSpelare5
            Exit Sub
        End If
        If Tärning4 = Tärning5 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
            ParSpelare5
            Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning4 Then
        If Tärning1 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare5
            Exit Sub
        End If
        If Tärning2 = Tärning5 Then
            Tärning = Tärning4 + Tärning2 + Tärning3 + Tärning5
            ParSpelare5
            Exit Sub
        End If
        If Tärning1 = Tärning5 Then
            Tärning = Tärning1 + Tärning3 + Tärning5 + Tärning4
            ParSpelare5
            Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning5 Then
        If Tärning2 = Tärning4 Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
            ParSpelare5
            Exit Sub
        End If
        If Tärning4 = Tärning1 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare5
            Exit Sub
        End If
        If Tärning1 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare5
            Exit Sub
        End If
    End If

    If Tärning4 = Tärning5 Then
        If Tärning2 = Tärning3 Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
            ParSpelare5
            Exit Sub
        End If
        If Tärning2 = Tärning1 Then
            Tärning = Tärning1 + Tärning4 + Tärning2 + Tärning5
            ParSpelare5
            Exit Sub
        End If
        If Tärning1 = Tärning3 Then
            Tärning = Tärning1 + Tärning3 + Tärning5 + Tärning4
            ParSpelare5
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning1 Then
        If Tärning2 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare5
            Exit Sub
        End If
        If Tärning4 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare5
            Exit Sub
        End If
        If Tärning5 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare5
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning4 Then
        If Tärning2 = Tärning5 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
            ParSpelare5
            Exit Sub
        End If
        If Tärning3 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare5
            Exit Sub
        End If
        If Tärning3 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning4 + Tärning3
            ParSpelare5
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning5 Then
        If Tärning2 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
            ParSpelare5
            Exit Sub
        End If
        If Tärning4 = Tärning3 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare5
            Exit Sub
        End If
        If Tärning3 = Tärning2 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare5
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning2 Then
        If Tärning1 = Tärning4 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare5
            Exit Sub
        End If
        If Tärning4 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning3 + Tärning5
            ParSpelare5
            Exit Sub
        End If
        If Tärning5 = Tärning1 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare5
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning4 Then
        If Tärning1 = Tärning3 Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
            ParSpelare5
            Exit Sub
        End If
        If Tärning1 = Tärning5 Then
            Tärning = Tärning1 + Tärning4 + Tärning2 + Tärning5
            ParSpelare5
            Exit Sub
        End If
        If Tärning3 = Tärning5 Then
            Tärning = Tärning4 + Tärning2 + Tärning5 + Tärning3
            ParSpelare5
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning5 Then
        If Tärning3 = Tärning4 Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
            ParSpelare5
            Exit Sub
        End If
        If Tärning4 = Tärning1 Then
            Tärning = Tärning1 + Tärning4 + Tärning2 + Tärning5
            ParSpelare5
            Exit Sub
        End If
        If Tärning1 = Tärning3 Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning3
            ParSpelare5
            Exit Sub
        End If
    End If
    Response = MsgBox("Du har inte två par! Vill du strycka två par?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmd2Par5.Enabled = False
        cmd2Par5.Caption = "-"
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmdChans1_Click() 'Chansen för spelare 1
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
    cmdChans1.Caption = Tärning
    txtUndre1.Text = txtUndre1.Text + Tärning
    cmdKasta.Enabled = False
    cmdNästa.Enabled = True
    cmdChans1.Enabled = False
    End If
End Sub

Private Sub cmdChans2_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
    cmdChans2.Caption = Tärning
    txtUndre2.Text = txtUndre2.Text + Tärning
    cmdKasta.Enabled = False
    cmdNästa.Enabled = True
    cmdChans2.Enabled = False
    End If
End Sub

Private Sub cmdChans3_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
    cmdChans3.Caption = Tärning
    txtUndre3.Text = txtUndre3.Text + Tärning
    cmdKasta.Enabled = False
    cmdNästa.Enabled = True
    cmdChans3.Enabled = False
    End If
End Sub

Private Sub cmdChans4_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
    cmdChans4.Caption = Tärning
    txtUndre4.Text = txtUndre4.Text + Tärning
    cmdKasta.Enabled = False
    cmdNästa.Enabled = True
    cmdChans4.Enabled = False
    End If
End Sub

Private Sub cmdChans5_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
    cmdChans5.Caption = Tärning
    txtUndre5.Text = txtUndre5.Text + Tärning
    cmdKasta.Enabled = False
    cmdNästa.Enabled = True
    cmdChans5.Enabled = False
    End If
End Sub

Private Sub cmdEttor1_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 1 Then Counter = Counter + 1 'Spelare 1 ettor
    If lblTärning2.Caption = 1 Then Counter = Counter + 1
    If lblTärning3.Caption = 1 Then Counter = Counter + 1
    If lblTärning4.Caption = 1 Then Counter = Counter + 1
    If lblTärning5.Caption = 1 Then Counter = Counter + 1
    If Counter > 0 Then
        cmdEttor1.Caption = Counter
        txtÖvre1.Text = Counter + txtÖvre1.Text
        cmdEttor1.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en etta!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdEttor2_Click()
    If Fusk = 1 Then 'Ettor för spelare 2
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 1 Then Counter = Counter + 1
    If lblTärning2.Caption = 1 Then Counter = Counter + 1
    If lblTärning3.Caption = 1 Then Counter = Counter + 1
    If lblTärning4.Caption = 1 Then Counter = Counter + 1
    If lblTärning5.Caption = 1 Then Counter = Counter + 1
    If Counter > 0 Then
        cmdEttor2.Caption = Counter
        txtÖvre2.Text = Counter + txtÖvre2.Text
        cmdEttor2.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en etta!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdEttor3_Click()
    If Fusk = 1 Then 'Ettor för spelare 3
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 1 Then Counter = Counter + 1
    If lblTärning2.Caption = 1 Then Counter = Counter + 1
    If lblTärning3.Caption = 1 Then Counter = Counter + 1
    If lblTärning4.Caption = 1 Then Counter = Counter + 1
    If lblTärning5.Caption = 1 Then Counter = Counter + 1
    If Counter > 0 Then
        cmdEttor3.Caption = Counter
        txtÖvre3.Text = Counter + txtÖvre3.Text
        cmdEttor3.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en etta!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdEttor4_Click()
    If Fusk = 1 Then 'Ettor för spelare 4
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 1 Then Counter = Counter + 1
    If lblTärning2.Caption = 1 Then Counter = Counter + 1
    If lblTärning3.Caption = 1 Then Counter = Counter + 1
    If lblTärning4.Caption = 1 Then Counter = Counter + 1
    If lblTärning5.Caption = 1 Then Counter = Counter + 1
    If Counter > 0 Then
        cmdEttor4.Caption = Counter
        txtÖvre4.Text = Counter + txtÖvre4.Text
        cmdEttor4.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en etta!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdEttor5_Click()
    If Fusk = 1 Then 'Ettor för spelare 5
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 1 Then Counter = Counter + 1
    If lblTärning2.Caption = 1 Then Counter = Counter + 1
    If lblTärning3.Caption = 1 Then Counter = Counter + 1
    If lblTärning4.Caption = 1 Then Counter = Counter + 1
    If lblTärning5.Caption = 1 Then Counter = Counter + 1
    If Counter > 0 Then
        cmdEttor5.Caption = Counter
        txtÖvre5.Text = Counter + txtÖvre5.Text
        cmdEttor5.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en etta!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdFemmor1_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 5 Then Counter = Counter + 5 'Spelare 1 Femmor
    If lblTärning2.Caption = 5 Then Counter = Counter + 5
    If lblTärning3.Caption = 5 Then Counter = Counter + 5
    If lblTärning4.Caption = 5 Then Counter = Counter + 5
    If lblTärning5.Caption = 5 Then Counter = Counter + 5
    If Counter > 0 Then
        cmdFemmor1.Caption = Counter
        txtÖvre1.Text = Counter + txtÖvre1.Text
        cmdFemmor1.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en femma!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdFemmor2_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 5 Then Counter = Counter + 5 'Spelare 1 Femmor
    If lblTärning2.Caption = 5 Then Counter = Counter + 5
    If lblTärning3.Caption = 5 Then Counter = Counter + 5
    If lblTärning4.Caption = 5 Then Counter = Counter + 5
    If lblTärning5.Caption = 5 Then Counter = Counter + 5
    If Counter > 0 Then
        cmdFemmor2.Caption = Counter
        txtÖvre2.Text = Counter + txtÖvre2.Text
        cmdFemmor2.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en femma!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdFemmor3_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 5 Then Counter = Counter + 5 'Spelare 1 Femmor
    If lblTärning2.Caption = 5 Then Counter = Counter + 5
    If lblTärning3.Caption = 5 Then Counter = Counter + 5
    If lblTärning4.Caption = 5 Then Counter = Counter + 5
    If lblTärning5.Caption = 5 Then Counter = Counter + 5
    If Counter > 0 Then
        cmdFemmor3.Caption = Counter
        txtÖvre3.Text = Counter + txtÖvre3.Text
        cmdFemmor3.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en femma!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdFemmor4_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 5 Then Counter = Counter + 5 'Spelare 1 Femmor
    If lblTärning2.Caption = 5 Then Counter = Counter + 5
    If lblTärning3.Caption = 5 Then Counter = Counter + 5
    If lblTärning4.Caption = 5 Then Counter = Counter + 5
    If lblTärning5.Caption = 5 Then Counter = Counter + 5
    If Counter > 0 Then
        cmdFemmor4.Caption = Counter
        txtÖvre4.Text = Counter + txtÖvre4.Text
        cmdFemmor4.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en femma!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdFemmor5_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 5 Then Counter = Counter + 5 'Spelare 1 Femmor
    If lblTärning2.Caption = 5 Then Counter = Counter + 5
    If lblTärning3.Caption = 5 Then Counter = Counter + 5
    If lblTärning4.Caption = 5 Then Counter = Counter + 5
    If lblTärning5.Caption = 5 Then Counter = Counter + 5
    If Counter > 0 Then
        cmdFemmor5.Caption = Counter
        txtÖvre5.Text = Counter + txtÖvre5.Text
        cmdFemmor5.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en femma!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdFyror1_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 4 Then Counter = Counter + 4 'Spelare 1 Fyror
    If lblTärning2.Caption = 4 Then Counter = Counter + 4
    If lblTärning3.Caption = 4 Then Counter = Counter + 4
    If lblTärning4.Caption = 4 Then Counter = Counter + 4
    If lblTärning5.Caption = 4 Then Counter = Counter + 4
    If Counter > 0 Then
        cmdFyror1.Caption = Counter
        txtÖvre1.Text = Counter + txtÖvre1.Text
        cmdFyror1.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en Fyra!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdFyror2_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 4 Then Counter = Counter + 4 'Spelare 1 Fyror
    If lblTärning2.Caption = 4 Then Counter = Counter + 4
    If lblTärning3.Caption = 4 Then Counter = Counter + 4
    If lblTärning4.Caption = 4 Then Counter = Counter + 4
    If lblTärning5.Caption = 4 Then Counter = Counter + 4
    If Counter > 0 Then
        cmdFyror2.Caption = Counter
        txtÖvre2.Text = Counter + txtÖvre2.Text
        cmdFyror2.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en Fyra!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdFyror3_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 4 Then Counter = Counter + 4 'Spelare 1 Fyror
    If lblTärning2.Caption = 4 Then Counter = Counter + 4
    If lblTärning3.Caption = 4 Then Counter = Counter + 4
    If lblTärning4.Caption = 4 Then Counter = Counter + 4
    If lblTärning5.Caption = 4 Then Counter = Counter + 4
    If Counter > 0 Then
        cmdFyror3.Caption = Counter
        txtÖvre3.Text = Counter + txtÖvre3.Text
        cmdFyror3.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en Fyra!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdFyror4_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 4 Then Counter = Counter + 4 'Spelare 4 Fyror
    If lblTärning2.Caption = 4 Then Counter = Counter + 4
    If lblTärning3.Caption = 4 Then Counter = Counter + 4
    If lblTärning4.Caption = 4 Then Counter = Counter + 4
    If lblTärning5.Caption = 4 Then Counter = Counter + 4
    If Counter > 0 Then
        cmdFyror4.Caption = Counter
        txtÖvre4.Text = Counter + txtÖvre4.Text
        cmdFyror4.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en Fyra!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdFyror5_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 4 Then Counter = Counter + 4 'Spelare 5 Fyror
    If lblTärning2.Caption = 4 Then Counter = Counter + 4
    If lblTärning3.Caption = 4 Then Counter = Counter + 4
    If lblTärning4.Caption = 4 Then Counter = Counter + 4
    If lblTärning5.Caption = 4 Then Counter = Counter + 4
    If Counter > 0 Then
        cmdFyror5.Caption = Counter
        txtÖvre5.Text = Counter + txtÖvre5.Text
        cmdFyror5.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en Fyra!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdFyrtal1_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning = 0                 'Fyrtal spelare 1
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
   
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning3) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
        Fyrtal
        Exit Sub
    End If
    If (Tärning1 = Tärning5) And (Tärning1 = Tärning3) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning3 + Tärning4 + Tärning5
        Fyrtal
        Exit Sub
    End If
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning3) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning5
        Fyrtal
        Exit Sub
    End If
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning5) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
        Fyrtal
        Exit Sub
    End If
    If (Tärning5 = Tärning2) And (Tärning5 = Tärning3) And (Tärning5 = Tärning4) Then
        Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
        Fyrtal
        Exit Sub
    End If
    Response = MsgBox("Du har inget fyrtal! Vill du strycka fyrtal?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmdFyrtal1.Enabled = False
        cmdFyrtal1.Caption = "-"
        Exit Sub
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmdFyrtal2_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning = 0                 'Fyrtal spelare 2
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
   
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning3) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
        Fyrtal2
        Exit Sub
    End If
    If (Tärning1 = Tärning5) And (Tärning1 = Tärning3) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning3 + Tärning4 + Tärning5
        Fyrtal2
        Exit Sub
    End If
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning3) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning5
        Fyrtal2
        Exit Sub
    End If
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning5) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
        Fyrtal2
        Exit Sub
    End If
    If (Tärning5 = Tärning2) And (Tärning5 = Tärning3) And (Tärning5 = Tärning4) Then
        Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
        Fyrtal2
        Exit Sub
    End If
    Response = MsgBox("Du har inget fyrtal! Vill du strycka fyrtal?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmdFyrtal2.Enabled = False
        cmdFyrtal2.Caption = "-"
        Exit Sub
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmdFyrtal3_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning = 0                 'Fyrtal spelare 3
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
   
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning3) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
        Fyrtal3
        Exit Sub
    End If
    If (Tärning1 = Tärning5) And (Tärning1 = Tärning3) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning3 + Tärning4 + Tärning5
        Fyrtal3
        Exit Sub
    End If
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning3) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning5
        Fyrtal3
        Exit Sub
    End If
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning5) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
        Fyrtal3
        Exit Sub
    End If
    If (Tärning5 = Tärning2) And (Tärning5 = Tärning3) And (Tärning5 = Tärning4) Then
        Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
        Fyrtal3
        Exit Sub
    End If
    Response = MsgBox("Du har inget fyrtal! Vill du strycka fyrtal?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmdFyrtal3.Enabled = False
        cmdFyrtal3.Caption = "-"
        Exit Sub
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmdFyrtal4_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning = 0                 'Fyrtal spelare 4
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
   
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning3) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
        Fyrtal4
        Exit Sub
    End If
    If (Tärning1 = Tärning5) And (Tärning1 = Tärning3) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning3 + Tärning4 + Tärning5
        Fyrtal4
        Exit Sub
    End If
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning3) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning5
        Fyrtal4
        Exit Sub
    End If
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning5) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
        Fyrtal4
        Exit Sub
    End If
    If (Tärning5 = Tärning2) And (Tärning5 = Tärning3) And (Tärning5 = Tärning4) Then
        Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
        Fyrtal4
        Exit Sub
    End If
    Response = MsgBox("Du har inget fyrtal! Vill du strycka fyrtal?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmdFyrtal4.Enabled = False
        cmdFyrtal4.Caption = "-"
        Exit Sub
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmdFyrtal5_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning = 0                 'Fyrtal spelare 1
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
   
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning3) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4
        Fyrtal5
        Exit Sub
    End If
    If (Tärning1 = Tärning5) And (Tärning1 = Tärning3) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning3 + Tärning4 + Tärning5
        Fyrtal5
        Exit Sub
    End If
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning3) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning5
        Fyrtal5
        Exit Sub
    End If
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning5) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4
        Fyrtal5
        Exit Sub
    End If
    If (Tärning5 = Tärning2) And (Tärning5 = Tärning3) And (Tärning5 = Tärning4) Then
        Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4
        Fyrtal5
        Exit Sub
    End If
    Response = MsgBox("Du har inget fyrtal! Vill du strycka fyrtal?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmdFyrtal5.Enabled = False
        cmdFyrtal5.Caption = "-"
        Exit Sub
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmdKasta_Click()
    If Kast < 3 Then
       Fusk = 1
       imgVisa(0).Visible = True
       imgVisa(1).Visible = True
       imgVisa(2).Visible = True
       imgVisa(3).Visible = True
       imgVisa(4).Visible = True
        
        If chk1.Value = 0 Then imgVisa2(3).Visible = False
        If chk2.Value = 0 Then imgVisa2(2).Visible = False
        If chk3.Value = 0 Then imgVisa2(4).Visible = False
        If chk4.Value = 0 Then imgVisa2(0).Visible = False
        If chk5.Value = 0 Then imgVisa2(1).Visible = False
        
        Animering = 0
        Visa = 0
        Current = 0
        Kast = Kast + 1 'Tärningen
        Randomize
        If chk1.Value = 0 Then lblTärning1.Caption = Int((6 * Rnd) + 1)
        If chk2.Value = 0 Then lblTärning2.Caption = Int((6 * Rnd) + 1)
        If chk3.Value = 0 Then lblTärning3.Caption = Int((6 * Rnd) + 1)
        If chk4.Value = 0 Then lblTärning4.Caption = Int((6 * Rnd) + 1)
        If chk5.Value = 0 Then lblTärning5.Caption = Int((6 * Rnd) + 1)
    End If
    If Kast > 2 Then
        cmdKasta.Enabled = False
    End If
End Sub

Private Sub cmdKlara_Click()
    Response = MsgBox("Är du säker på att du vill avslute detta parti?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdSumma1.Caption = txtÖvre1.Text  'Spelare 1's Totala poäng
        If txtÖvre1.Text > 62 Then
            cmdBonus1.Caption = 50
            txtÖvre1.Text = txtÖvre1.Text + 50
        End If
        Poäng1 = txtUndre1.Text
        Poäng2 = txtÖvre1.Text
        Poäng = Poäng1 + Poäng2
        cmd2Summa1.Caption = Poäng
        
        cmdSumma2.Caption = txtÖvre2.Text 'Spelare 2's totala Poäng
        If txtÖvre2.Text > 62 Then
            cmdBonus2.Caption = 50
            txtÖvre2.Text = txtÖvre2.Text + 50
        End If
        Poäng1 = txtUndre2.Text
        Poäng2 = txtÖvre2.Text
        Poäng3 = Poäng1 + Poäng2
        cmd2Summa2.Caption = Poäng3
    
        
        cmdSumma3.Caption = txtÖvre3.Text   'Spelare 3's Totala Poäng
        If txtÖvre3.Text > 62 Then
            cmdBonus3.Caption = 50
            txtÖvre3.Text = txtÖvre3.Text + 50
        End If
        Poäng1 = txtUndre3.Text
        Poäng2 = txtÖvre3.Text
        Poäng4 = Poäng1 + Poäng2
        cmd2Summa3.Caption = Poäng4
    
        cmdSumma4.Caption = txtÖvre4.Text 'Spelare 4's Totala Poäng
        If txtÖvre4.Text > 62 Then
            cmdBonus4.Caption = 50
            txtÖvre4.Text = txtÖvre4.Text + 50
        End If
        Poäng1 = txtUndre4.Text
        Poäng2 = txtÖvre4.Text
        Poäng5 = Poäng1 + Poäng2
        cmd2Summa4.Caption = Poäng5
        cmd2Summa4.FontSize = 12
        cmd2Summa4.FontBold = True
    
        cmdSumma5.Caption = txtÖvre5.Text   'Spelare 5's Totala Poäng
        If txtÖvre5.Text > 62 Then
            cmdBonus5.Caption = 50
            txtÖvre5.Text = txtÖvre5.Text + 50
        End If
        Poäng1 = txtUndre5.Text
        Poäng2 = txtÖvre5.Text
        Poäng6 = Poäng1 + Poäng2
        cmd2Summa5.Caption = Poäng6
    
        cmdKasta.Enabled = False
        cmdNästa.Enabled = False
        cmdKlara.Enabled = False
    End If
    
    'High Score

    If Poäng > txtScore3.Text Then
        If cmd2Summa1.Caption > txtScore2.Text Then
            If cmd2Summa1.Caption > txtScore1.Text Then
                txtScore1.Text = cmd2Summa1.Caption
                txtName1.Text = lblSpelare1.Caption
                txtDate1.Text = Date
                Exit Sub
            End If
            txtScore2.Text = cmd2Summa1.Caption
            txtName2.Text = lblSpelare1.Caption
            txtDate2.Text = Date
            Exit Sub
        End If
        txtScore3.Text = cmd2Summa1.Caption
        txtName3.Text = lblSpelare1.Caption
        txtDate3.Text = Date
        Exit Sub
    End If

    If Poäng3 > txtScore3.Text Then
        If cmd2Summa2.Caption > txtScore2.Text Then
            If cmd2Summa2.Caption > txtScore1.Text Then
                txtScore1.Text = cmd2Summa2.Caption
                txtName1.Text = lblSpelare2.Caption
                txtDate1.Text = Date
                Exit Sub
            End If
            txtScore2.Text = cmd2Summa2.Caption
            txtName2.Text = lblSpelare2.Caption
            txtDate2.Text = Date
            Exit Sub
        End If
        txtScore3.Text = cmd2Summa2.Caption
        txtName3.Text = lblSpelare2.Caption
        txtDate3.Text = Date
        Exit Sub
    End If

    If Poäng > txtScore3.Text Then
        If cmd2Summa1.Caption > txtScore2.Text Then
            If cmd2Summa3.Caption > txtScore1.Text Then
                txtScore1.Text = cmd2Summa3.Caption
                txtName1.Text = lblSpelare3.Caption
                txtDate1.Text = Date
                Exit Sub
            End If
            txtScore2.Text = cmd2Summa3.Caption
            txtName2.Text = lblSpelare3.Caption
            txtDate2.Text = Date
            Exit Sub
        End If
        txtScore3.Text = cmd2Summa3.Caption
        txtName3.Text = lblSpelare3.Caption
        txtDate3.Text = Date
        Exit Sub
    End If

    If Poäng5 > txtScore3.Text Then
        If cmd2Summa4.Caption > txtScore2.Text Then
            If cmd2Summa4.Caption > txtScore1.Text Then
                txtScore1.Text = cmd2Summa4.Caption
                txtName1.Text = lblSpelare4.Caption
                txtDate1.Text = Date
                Exit Sub
            End If
            txtScore2.Text = cmd2Summa4.Caption
            txtName2.Text = lblSpelare4.Caption
            txtDate2.Text = Date
            Exit Sub
        End If
        txtScore3.Text = cmd2Summa4.Caption
        txtName3.Text = lblSpelare4.Caption
        txtDate3.Text = Date
        Exit Sub
    End If

    If Poäng6 > txtScore3.Text Then
        If cmd2Summa5.Caption > txtScore2.Text Then
            If cmd2Summa5.Caption > txtScore1.Text Then
                txtScore1.Text = cmd2Summa5.Caption
                txtName1.Text = lblSpelare5.Caption
                txtDate1.Text = Date
                Exit Sub
            End If
            txtScore2.Text = cmd2Summa5.Caption
            txtName2.Text = lblSpelare5.Caption
            txtDate2.Text = Date
            Exit Sub
        End If
        txtScore3.Text = cmd2Summa5.Caption
        txtName3.Text = lblSpelare5.Caption
        txtDate3.Text = Date
        Exit Sub
    End If

End Sub

Private Sub cmdKåk1_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption  'Kåk, Spelare ett
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = 0
 
    If Tärning1 = Tärning2 Then
        If (Tärning3 = Tärning4) And (Tärning3 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk
        Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning4 Then
        If (Tärning1 = Tärning2) And (Tärning1 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk
            Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning5 Then
        If (Tärning2 = Tärning4) And (Tärning2 = Tärning1) Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4 + Tärning1
            Kåk
            Exit Sub
        End If
    End If

    If Tärning4 = Tärning5 Then
        If (Tärning2 = Tärning3) And (Tärning2 = Tärning1) Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4 + Tärning1
            Kåk
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning1 Then
        If (Tärning2 = Tärning4) And (Tärning3 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning4 Then
        If (Tärning2 = Tärning5) And (Tärning2 = Tärning3) Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4 + Tärning3
            Kåk
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning5 Then
        If (Tärning2 = Tärning4) And (Tärning2 = Tärning3) Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4 + Tärning3
            Kåk
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning2 Then
        If (Tärning1 = Tärning4) And (Tärning1 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning4 Then
        If (Tärning1 = Tärning3) And (Tärning1 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning5 Then
        If (Tärning3 = Tärning4) And (Tärning3 = Tärning1) Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4 + Tärning1
            Kåk
            Exit Sub
        End If
    End If
    Response = MsgBox("Du har ingen kåk! Vill du strycka kåk?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmdKåk1.Enabled = False
        cmdKåk1.Caption = "-"
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmdKåk2_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption  'Kåk, Spelare ett
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = 0
 
    If Tärning1 = Tärning2 Then
        If (Tärning3 = Tärning4) And (Tärning3 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk2
        Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning4 Then
        If (Tärning1 = Tärning2) And (Tärning1 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk2
            Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning5 Then
        If (Tärning2 = Tärning4) And (Tärning2 = Tärning1) Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4 + Tärning1
            Kåk2
            Exit Sub
        End If
    End If

    If Tärning4 = Tärning5 Then
        If (Tärning2 = Tärning3) And (Tärning2 = Tärning1) Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4 + Tärning1
            Kåk2
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning1 Then
        If (Tärning2 = Tärning4) And (Tärning3 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk2
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning4 Then
        If (Tärning2 = Tärning5) And (Tärning2 = Tärning3) Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4 + Tärning3
            Kåk2
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning5 Then
        If (Tärning2 = Tärning4) And (Tärning2 = Tärning3) Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4 + Tärning3
            Kåk2
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning2 Then
        If (Tärning1 = Tärning4) And (Tärning1 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk2
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning4 Then
        If (Tärning1 = Tärning3) And (Tärning1 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk2
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning5 Then
        If (Tärning3 = Tärning4) And (Tärning3 = Tärning1) Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4 + Tärning1
            Kåk2
            Exit Sub
        End If
    End If
    Response = MsgBox("Du har ingen kåk! Vill du strycka kåk?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmdKåk2.Enabled = False
        cmdKåk2.Caption = "-"
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmdKåk3_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption  'Kåk, Spelare ett
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = 0
 
    If Tärning1 = Tärning2 Then
        If (Tärning3 = Tärning4) And (Tärning3 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk3
        Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning4 Then
        If (Tärning1 = Tärning2) And (Tärning1 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk3
            Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning5 Then
        If (Tärning2 = Tärning4) And (Tärning2 = Tärning1) Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4 + Tärning1
            Kåk3
            Exit Sub
        End If
    End If

    If Tärning4 = Tärning5 Then
        If (Tärning2 = Tärning3) And (Tärning2 = Tärning1) Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4 + Tärning1
            Kåk3
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning1 Then
        If (Tärning2 = Tärning4) And (Tärning3 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk3
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning4 Then
        If (Tärning2 = Tärning5) And (Tärning2 = Tärning3) Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4 + Tärning3
            Kåk3
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning5 Then
        If (Tärning2 = Tärning4) And (Tärning2 = Tärning3) Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4 + Tärning3
            Kåk3
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning2 Then
        If (Tärning1 = Tärning4) And (Tärning1 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk3
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning4 Then
        If (Tärning1 = Tärning3) And (Tärning1 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk3
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning5 Then
        If (Tärning3 = Tärning4) And (Tärning3 = Tärning1) Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4 + Tärning1
            Kåk3
            Exit Sub
        End If
    End If
    Response = MsgBox("Du har ingen kåk! Vill du strycka kåk?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmdKåk3.Enabled = False
        cmdKåk3.Caption = "-"
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmdKåk4_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption  'Kåk, Spelare ett
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = 0
 
    If Tärning1 = Tärning2 Then
        If (Tärning3 = Tärning4) And (Tärning3 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk4
        Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning4 Then
        If (Tärning1 = Tärning2) And (Tärning1 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk4
            Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning5 Then
        If (Tärning2 = Tärning4) And (Tärning2 = Tärning1) Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4 + Tärning1
            Kåk4
            Exit Sub
        End If
    End If

    If Tärning4 = Tärning5 Then
        If (Tärning2 = Tärning3) And (Tärning2 = Tärning1) Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4 + Tärning1
            Kåk4
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning1 Then
        If (Tärning2 = Tärning4) And (Tärning3 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk4
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning4 Then
        If (Tärning2 = Tärning5) And (Tärning2 = Tärning3) Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4 + Tärning3
            Kåk4
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning5 Then
        If (Tärning2 = Tärning4) And (Tärning2 = Tärning3) Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4 + Tärning3
            Kåk4
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning2 Then
        If (Tärning1 = Tärning4) And (Tärning1 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk4
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning4 Then
        If (Tärning1 = Tärning3) And (Tärning1 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk4
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning5 Then
        If (Tärning3 = Tärning4) And (Tärning3 = Tärning1) Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4 + Tärning1
            Kåk4
            Exit Sub
        End If
    End If
    Response = MsgBox("Du har ingen kåk! Vill du strycka kåk?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmdKåk4.Enabled = False
        cmdKåk4.Caption = "-"
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmdKåk5_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning1 = lblTärning1.Caption  'Kåk, Spelare ett
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    Tärning = 0
 
    If Tärning1 = Tärning2 Then
        If (Tärning3 = Tärning4) And (Tärning3 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk5
        Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning4 Then
        If (Tärning1 = Tärning2) And (Tärning1 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk5
            Exit Sub
        End If
    End If
        
    If Tärning3 = Tärning5 Then
        If (Tärning2 = Tärning4) And (Tärning2 = Tärning1) Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4 + Tärning1
            Kåk5
            Exit Sub
        End If
    End If

    If Tärning4 = Tärning5 Then
        If (Tärning2 = Tärning3) And (Tärning2 = Tärning1) Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4 + Tärning1
            Kåk5
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning1 Then
        If (Tärning2 = Tärning4) And (Tärning3 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk5
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning4 Then
        If (Tärning2 = Tärning5) And (Tärning2 = Tärning3) Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4 + Tärning3
            Kåk5
            Exit Sub
        End If
    End If

    If Tärning1 = Tärning5 Then
        If (Tärning2 = Tärning4) And (Tärning2 = Tärning3) Then
            Tärning = Tärning1 + Tärning2 + Tärning5 + Tärning4 + Tärning3
            Kåk5
            Exit Sub
        End If
    End If

    If Tärning3 = Tärning2 Then
        If (Tärning1 = Tärning4) And (Tärning1 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk5
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning4 Then
        If (Tärning1 = Tärning3) And (Tärning1 = Tärning5) Then
            Tärning = Tärning1 + Tärning2 + Tärning3 + Tärning4 + Tärning5
            Kåk5
            Exit Sub
        End If
    End If

    If Tärning2 = Tärning5 Then
        If (Tärning3 = Tärning4) And (Tärning3 = Tärning1) Then
            Tärning = Tärning5 + Tärning2 + Tärning3 + Tärning4 + Tärning1
            Kåk5
            Exit Sub
        End If
    End If
    Response = MsgBox("Du har ingen kåk! Vill du strycka kåk?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmdKåk5.Enabled = False
        cmdKåk5.Caption = "-"
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmdLiten1_Click() 'Liten stege
    If Fusk = 1 Then
    Fusk = 2
    If lblTärning1.Caption = 1 Or lblTärning2.Caption = 1 Or lblTärning3.Caption = 1 Or lblTärning4.Caption = 1 Or lblTärning5.Caption = 1 Then
        If lblTärning1.Caption = 2 Or lblTärning2.Caption = 2 Or lblTärning3.Caption = 2 Or lblTärning4.Caption = 2 Or lblTärning5.Caption = 2 Then
            If lblTärning1.Caption = 3 Or lblTärning2.Caption = 3 Or lblTärning3.Caption = 3 Or lblTärning4.Caption = 3 Or lblTärning5.Caption = 3 Then
                If lblTärning1.Caption = 4 Or lblTärning2.Caption = 4 Or lblTärning3.Caption = 4 Or lblTärning4.Caption = 4 Or lblTärning5.Caption = 4 Then
                    If lblTärning1.Caption = 5 Or lblTärning2.Caption = 5 Or lblTärning3.Caption = 5 Or lblTärning4.Caption = 5 Or lblTärning5.Caption = 5 Then
                        cmdLiten1.Enabled = False
                    Else
                        Liten
                        Exit Sub
                    End If
                Else
                    Liten
                    Exit Sub
                End If
            Else
                Liten
                Exit Sub
            End If
        Else
            Liten
            Exit Sub
        End If
    Else
        Liten
        Exit Sub
    End If
    cmdLiten1.Caption = 15
    txtUndre1.Text = txtUndre1.Text + 15
    cmdKasta.Enabled = False
    cmdNästa.Enabled = True
    End If
End Sub

Private Sub cmdLiten2_Click()
    If Fusk = 1 Then
    Fusk = 2
    If lblTärning1.Caption = 1 Or lblTärning2.Caption = 1 Or lblTärning3.Caption = 1 Or lblTärning4.Caption = 1 Or lblTärning5.Caption = 1 Then
        If lblTärning1.Caption = 2 Or lblTärning2.Caption = 2 Or lblTärning3.Caption = 2 Or lblTärning4.Caption = 2 Or lblTärning5.Caption = 2 Then
            If lblTärning1.Caption = 3 Or lblTärning2.Caption = 3 Or lblTärning3.Caption = 3 Or lblTärning4.Caption = 3 Or lblTärning5.Caption = 3 Then
                If lblTärning1.Caption = 4 Or lblTärning2.Caption = 4 Or lblTärning3.Caption = 4 Or lblTärning4.Caption = 4 Or lblTärning5.Caption = 4 Then
                    If lblTärning1.Caption = 5 Or lblTärning2.Caption = 5 Or lblTärning3.Caption = 5 Or lblTärning4.Caption = 5 Or lblTärning5.Caption = 5 Then
                        cmdLiten2.Enabled = False
                    Else
                        Liten2
                        Exit Sub
                    End If
                Else
                    Liten2
                    Exit Sub
                End If
            Else
                Liten2
                Exit Sub
            End If
        Else
            Liten2
            Exit Sub
        End If
    Else
        Liten2
        Exit Sub
    End If
    cmdLiten2.Caption = 15
    txtUndre2.Text = txtUndre2.Text + 15
    cmdKasta.Enabled = False
    cmdNästa.Enabled = True
    End If
End Sub

Private Sub cmdLiten3_Click()
    If Fusk = 1 Then
    Fusk = 2
    If lblTärning1.Caption = 1 Or lblTärning2.Caption = 1 Or lblTärning3.Caption = 1 Or lblTärning4.Caption = 1 Or lblTärning5.Caption = 1 Then
        If lblTärning1.Caption = 2 Or lblTärning2.Caption = 2 Or lblTärning3.Caption = 2 Or lblTärning4.Caption = 2 Or lblTärning5.Caption = 2 Then
            If lblTärning1.Caption = 3 Or lblTärning2.Caption = 3 Or lblTärning3.Caption = 3 Or lblTärning4.Caption = 3 Or lblTärning5.Caption = 3 Then
                If lblTärning1.Caption = 4 Or lblTärning2.Caption = 4 Or lblTärning3.Caption = 4 Or lblTärning4.Caption = 4 Or lblTärning5.Caption = 4 Then
                    If lblTärning1.Caption = 5 Or lblTärning2.Caption = 5 Or lblTärning3.Caption = 5 Or lblTärning4.Caption = 5 Or lblTärning5.Caption = 5 Then
                        cmdLiten3.Enabled = False
                    Else
                        Liten3
                        Exit Sub
                    End If
                Else
                    Liten3
                    Exit Sub
                End If
            Else
                Liten3
                Exit Sub
            End If
        Else
            Liten3
            Exit Sub
        End If
    Else
        Liten3
        Exit Sub
    End If
    cmdLiten3.Caption = 15
    txtUndre3.Text = txtUndre3.Text + 15
    cmdKasta.Enabled = False
    cmdNästa.Enabled = True
    End If
End Sub

Private Sub cmdLiten4_Click()
    If Fusk = 1 Then
    Fusk = 2
    If lblTärning1.Caption = 1 Or lblTärning2.Caption = 1 Or lblTärning3.Caption = 1 Or lblTärning4.Caption = 1 Or lblTärning5.Caption = 1 Then
        If lblTärning1.Caption = 2 Or lblTärning2.Caption = 2 Or lblTärning3.Caption = 2 Or lblTärning4.Caption = 2 Or lblTärning5.Caption = 2 Then
            If lblTärning1.Caption = 3 Or lblTärning2.Caption = 3 Or lblTärning3.Caption = 3 Or lblTärning4.Caption = 3 Or lblTärning5.Caption = 3 Then
                If lblTärning1.Caption = 4 Or lblTärning2.Caption = 4 Or lblTärning3.Caption = 4 Or lblTärning4.Caption = 4 Or lblTärning5.Caption = 4 Then
                    If lblTärning1.Caption = 5 Or lblTärning2.Caption = 5 Or lblTärning3.Caption = 5 Or lblTärning4.Caption = 5 Or lblTärning5.Caption = 5 Then
                        cmdLiten4.Enabled = False
                    Else
                        Liten4
                        Exit Sub
                    End If
                Else
                    Liten4
                    Exit Sub
                End If
            Else
                Liten4
                Exit Sub
            End If
        Else
            Liten4
            Exit Sub
        End If
    Else
        Liten4
        Exit Sub
    End If
    cmdLiten4.Caption = 15
    txtUndre4.Text = txtUndre4.Text + 15
    cmdKasta.Enabled = False
    cmdNästa.Enabled = True
    End If
End Sub

Private Sub cmdLiten5_Click()
    If Fusk = 1 Then
    Fusk = 2
    If lblTärning1.Caption = 1 Or lblTärning2.Caption = 1 Or lblTärning3.Caption = 1 Or lblTärning4.Caption = 1 Or lblTärning5.Caption = 1 Then
        If lblTärning1.Caption = 2 Or lblTärning2.Caption = 2 Or lblTärning3.Caption = 2 Or lblTärning4.Caption = 2 Or lblTärning5.Caption = 2 Then
            If lblTärning1.Caption = 3 Or lblTärning2.Caption = 3 Or lblTärning3.Caption = 3 Or lblTärning4.Caption = 3 Or lblTärning5.Caption = 3 Then
                If lblTärning1.Caption = 4 Or lblTärning2.Caption = 4 Or lblTärning3.Caption = 4 Or lblTärning4.Caption = 4 Or lblTärning5.Caption = 4 Then
                    If lblTärning1.Caption = 5 Or lblTärning2.Caption = 5 Or lblTärning3.Caption = 5 Or lblTärning4.Caption = 5 Or lblTärning5.Caption = 5 Then
                        cmdLiten5.Enabled = False
                    Else
                        Liten5
                        Exit Sub
                    End If
                Else
                    Liten5
                    Exit Sub
                End If
            Else
                Liten5
                Exit Sub
            End If
        Else
            Liten5
            Exit Sub
        End If
    Else
        Liten5
        Exit Sub
    End If
    cmdLiten5.Caption = 15
    txtUndre5.Text = txtUndre5.Text + 15
    cmdKasta.Enabled = False
    cmdNästa.Enabled = True
    End If
End Sub

Private Sub cmdNästa_Click()
    imgVisa2(0).Visible = False
    imgVisa2(1).Visible = False
    imgVisa2(2).Visible = False
    imgVisa2(3).Visible = False
    imgVisa2(4).Visible = False
    imgVisa(0).Visible = False
    imgVisa(1).Visible = False
    imgVisa(2).Visible = False
    imgVisa(3).Visible = False
    imgVisa(4).Visible = False
    Kast = 0
    cmdNästa.Enabled = False
    cmdKasta.Enabled = True
    lblTärning1.Caption = ""
    lblTärning2.Caption = ""
    lblTärning3.Caption = ""
    lblTärning4.Caption = ""
    lblTärning5.Caption = ""
    chk1.Value = 0
    chk2.Value = 0
    chk3.Value = 0
    chk4.Value = 0
    chk5.Value = 0
End Sub

Private Sub cmdSexor1_Click()
    If Fusk = 1 Then
    Counter = 0
    Fusk = 2
    If lblTärning1.Caption = 6 Then Counter = Counter + 6
    If lblTärning2.Caption = 6 Then Counter = Counter + 6
    If lblTärning3.Caption = 6 Then Counter = Counter + 6
    If lblTärning4.Caption = 6 Then Counter = Counter + 6
    If lblTärning5.Caption = 6 Then Counter = Counter + 6
    If Counter > 0 Then
        cmdSexor1.Caption = Counter
        txtÖvre1.Text = Counter + txtÖvre1.Text
        cmdSexor1.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en sexa!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdSexor2_Click()
    If Fusk = 1 Then
    Counter = 0
    Fusk = 2
    If lblTärning1.Caption = 6 Then Counter = Counter + 6
    If lblTärning2.Caption = 6 Then Counter = Counter + 6
    If lblTärning3.Caption = 6 Then Counter = Counter + 6
    If lblTärning4.Caption = 6 Then Counter = Counter + 6
    If lblTärning5.Caption = 6 Then Counter = Counter + 6
    If Counter > 0 Then
        cmdSexor2.Caption = Counter
        txtÖvre2.Text = Counter + txtÖvre2.Text
        cmdSexor2.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en sexa!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdSexor3_Click()
    If Fusk = 1 Then
    Counter = 0
    Fusk = 2
    If lblTärning1.Caption = 6 Then Counter = Counter + 6
    If lblTärning2.Caption = 6 Then Counter = Counter + 6
    If lblTärning3.Caption = 6 Then Counter = Counter + 6
    If lblTärning4.Caption = 6 Then Counter = Counter + 6
    If lblTärning5.Caption = 6 Then Counter = Counter + 6
    If Counter > 0 Then
        cmdSexor3.Caption = Counter
        txtÖvre3.Text = Counter + txtÖvre3.Text
        cmdSexor3.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en sexa!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdSexor4_Click()
    If Fusk = 1 Then
    Counter = 0
    Fusk = 2
    If lblTärning1.Caption = 6 Then Counter = Counter + 6
    If lblTärning2.Caption = 6 Then Counter = Counter + 6
    If lblTärning3.Caption = 6 Then Counter = Counter + 6
    If lblTärning4.Caption = 6 Then Counter = Counter + 6
    If lblTärning5.Caption = 6 Then Counter = Counter + 6
    If Counter > 0 Then
        cmdSexor4.Caption = Counter
        txtÖvre4.Text = Counter + txtÖvre4.Text
        cmdSexor4.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en sexa!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdSexor5_Click()
    If Fusk = 1 Then
    Counter = 0
    Fusk = 2
    If lblTärning1.Caption = 6 Then Counter = Counter + 6
    If lblTärning2.Caption = 6 Then Counter = Counter + 6
    If lblTärning3.Caption = 6 Then Counter = Counter + 6
    If lblTärning4.Caption = 6 Then Counter = Counter + 6
    If lblTärning5.Caption = 6 Then Counter = Counter + 6
    If Counter > 0 Then
        cmdSexor5.Caption = Counter
        txtÖvre5.Text = Counter + txtÖvre5.Text
        cmdSexor5.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en sexa!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdStor1_Click()
    If Fusk = 1 Then
    Fusk = 2
    If lblTärning1.Caption = 6 Or lblTärning2.Caption = 6 Or lblTärning3.Caption = 6 Or lblTärning4.Caption = 6 Or lblTärning5.Caption = 6 Then
        If lblTärning1.Caption = 2 Or lblTärning2.Caption = 2 Or lblTärning3.Caption = 2 Or lblTärning4.Caption = 2 Or lblTärning5.Caption = 2 Then
            If lblTärning1.Caption = 3 Or lblTärning2.Caption = 3 Or lblTärning3.Caption = 3 Or lblTärning4.Caption = 3 Or lblTärning5.Caption = 3 Then
                If lblTärning1.Caption = 4 Or lblTärning2.Caption = 4 Or lblTärning3.Caption = 4 Or lblTärning4.Caption = 4 Or lblTärning5.Caption = 4 Then
                    If lblTärning1.Caption = 5 Or lblTärning2.Caption = 5 Or lblTärning3.Caption = 5 Or lblTärning4.Caption = 5 Or lblTärning5.Caption = 5 Then
                        cmdStor1.Enabled = False
                    Else
                        Stor
                        Exit Sub
                    End If
                Else
                    Stor
                    Exit Sub
                End If
            Else
                Stor
                Exit Sub
            End If
        Else
            Stor
            Exit Sub
        End If
    Else
        Stor
        Exit Sub
    End If
    cmdStor1.Caption = 20
    txtUndre1.Text = txtUndre1.Text + 20
    cmdKasta.Enabled = False
    cmdNästa.Enabled = True
    End If
End Sub

Private Sub cmdStor2_Click()
    If Fusk = 1 Then
    Fusk = 2
    If lblTärning1.Caption = 6 Or lblTärning2.Caption = 6 Or lblTärning3.Caption = 6 Or lblTärning4.Caption = 6 Or lblTärning5.Caption = 6 Then
        If lblTärning1.Caption = 2 Or lblTärning2.Caption = 2 Or lblTärning3.Caption = 2 Or lblTärning4.Caption = 2 Or lblTärning5.Caption = 2 Then
            If lblTärning1.Caption = 3 Or lblTärning2.Caption = 3 Or lblTärning3.Caption = 3 Or lblTärning4.Caption = 3 Or lblTärning5.Caption = 3 Then
                If lblTärning1.Caption = 4 Or lblTärning2.Caption = 4 Or lblTärning3.Caption = 4 Or lblTärning4.Caption = 4 Or lblTärning5.Caption = 4 Then
                    If lblTärning1.Caption = 5 Or lblTärning2.Caption = 5 Or lblTärning3.Caption = 5 Or lblTärning4.Caption = 5 Or lblTärning5.Caption = 5 Then
                        cmdStor2.Enabled = False
                    Else
                        Stor2
                        Exit Sub
                    End If
                Else
                    Stor2
                    Exit Sub
                End If
            Else
                Stor2
                Exit Sub
            End If
        Else
            Stor2
            Exit Sub
        End If
    Else
        Stor2
        Exit Sub
    End If
    cmdStor2.Caption = 20
    txtUndre2.Text = txtUndre2.Text + 20
    cmdKasta.Enabled = False
    cmdNästa.Enabled = True
    End If
End Sub

Private Sub cmdStor3_Click()
    If Fusk = 1 Then
    Fusk = 2
    If lblTärning1.Caption = 6 Or lblTärning2.Caption = 6 Or lblTärning3.Caption = 6 Or lblTärning4.Caption = 6 Or lblTärning5.Caption = 6 Then
        If lblTärning1.Caption = 2 Or lblTärning2.Caption = 2 Or lblTärning3.Caption = 2 Or lblTärning4.Caption = 2 Or lblTärning5.Caption = 2 Then
            If lblTärning1.Caption = 3 Or lblTärning2.Caption = 3 Or lblTärning3.Caption = 3 Or lblTärning4.Caption = 3 Or lblTärning5.Caption = 3 Then
                If lblTärning1.Caption = 4 Or lblTärning2.Caption = 4 Or lblTärning3.Caption = 4 Or lblTärning4.Caption = 4 Or lblTärning5.Caption = 4 Then
                    If lblTärning1.Caption = 5 Or lblTärning2.Caption = 5 Or lblTärning3.Caption = 5 Or lblTärning4.Caption = 5 Or lblTärning5.Caption = 5 Then
                        cmdStor3.Enabled = False
                    Else
                        Stor3
                        Exit Sub
                    End If
                Else
                    Stor3
                    Exit Sub
                End If
            Else
                Stor3
                Exit Sub
            End If
        Else
            Stor3
            Exit Sub
        End If
    Else
        Stor3
        Exit Sub
    End If
    cmdStor3.Caption = 20
    txtUndre3.Text = txtUndre3.Text + 20
    cmdKasta.Enabled = False
    cmdNästa.Enabled = True
    End If
End Sub

Private Sub cmdStor4_Click()
    If Fusk = 1 Then
    Fusk = 2
    If lblTärning1.Caption = 6 Or lblTärning2.Caption = 6 Or lblTärning3.Caption = 6 Or lblTärning4.Caption = 6 Or lblTärning5.Caption = 6 Then
        If lblTärning1.Caption = 2 Or lblTärning2.Caption = 2 Or lblTärning3.Caption = 2 Or lblTärning4.Caption = 2 Or lblTärning5.Caption = 2 Then
            If lblTärning1.Caption = 3 Or lblTärning2.Caption = 3 Or lblTärning3.Caption = 3 Or lblTärning4.Caption = 3 Or lblTärning5.Caption = 3 Then
                If lblTärning1.Caption = 4 Or lblTärning2.Caption = 4 Or lblTärning3.Caption = 4 Or lblTärning4.Caption = 4 Or lblTärning5.Caption = 4 Then
                    If lblTärning1.Caption = 5 Or lblTärning2.Caption = 5 Or lblTärning3.Caption = 5 Or lblTärning4.Caption = 5 Or lblTärning5.Caption = 5 Then
                        cmdStor4.Enabled = False
                    Else
                        Stor4
                        Exit Sub
                    End If
                Else
                    Stor4
                    Exit Sub
                End If
            Else
                Stor4
                Exit Sub
            End If
        Else
            Stor4
            Exit Sub
        End If
    Else
        Stor4
        Exit Sub
    End If
    cmdStor4.Caption = 20
    txtUndre4.Text = txtUndre4.Text + 20
    cmdKasta.Enabled = False
    cmdNästa.Enabled = True
    End If
End Sub

Private Sub cmdStor5_Click()
    If Fusk = 1 Then
    Fusk = 2
    If lblTärning1.Caption = 6 Or lblTärning2.Caption = 6 Or lblTärning3.Caption = 6 Or lblTärning4.Caption = 6 Or lblTärning5.Caption = 6 Then
        If lblTärning1.Caption = 2 Or lblTärning2.Caption = 2 Or lblTärning3.Caption = 2 Or lblTärning4.Caption = 2 Or lblTärning5.Caption = 2 Then
            If lblTärning1.Caption = 3 Or lblTärning2.Caption = 3 Or lblTärning3.Caption = 3 Or lblTärning4.Caption = 3 Or lblTärning5.Caption = 3 Then
                If lblTärning1.Caption = 4 Or lblTärning2.Caption = 4 Or lblTärning3.Caption = 4 Or lblTärning4.Caption = 4 Or lblTärning5.Caption = 4 Then
                    If lblTärning1.Caption = 5 Or lblTärning2.Caption = 5 Or lblTärning3.Caption = 5 Or lblTärning4.Caption = 5 Or lblTärning5.Caption = 5 Then
                        cmdStor5.Enabled = False
                    Else
                        Stor5
                        Exit Sub
                    End If
                Else
                    Stor5
                    Exit Sub
                End If
            Else
                Stor5
                Exit Sub
            End If
        Else
            Stor5
            Exit Sub
        End If
    Else
        Stor5
        Exit Sub
    End If
    cmdStor5.Caption = 20
    txtUndre5.Text = txtUndre5.Text + 20
    cmdKasta.Enabled = False
    cmdNästa.Enabled = True
    End If
End Sub

Private Sub cmdTreor1_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 3 Then Counter = Counter + 3
    If lblTärning2.Caption = 3 Then Counter = Counter + 3
    If lblTärning3.Caption = 3 Then Counter = Counter + 3
    If lblTärning4.Caption = 3 Then Counter = Counter + 3
    If lblTärning5.Caption = 3 Then Counter = Counter + 3
    If Counter > 0 Then
        cmdTreor1.Caption = Counter
        txtÖvre1.Text = Counter + txtÖvre1.Text
        cmdTreor1.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en trea!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdTreor2_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 3 Then Counter = Counter + 3
    If lblTärning2.Caption = 3 Then Counter = Counter + 3
    If lblTärning3.Caption = 3 Then Counter = Counter + 3
    If lblTärning4.Caption = 3 Then Counter = Counter + 3
    If lblTärning5.Caption = 3 Then Counter = Counter + 3
    If Counter > 0 Then
        cmdTreor2.Caption = Counter
        txtÖvre2.Text = Counter + txtÖvre2.Text
        cmdTreor2.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en trea!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdTreor3_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 3 Then Counter = Counter + 3
    If lblTärning2.Caption = 3 Then Counter = Counter + 3
    If lblTärning3.Caption = 3 Then Counter = Counter + 3
    If lblTärning4.Caption = 3 Then Counter = Counter + 3
    If lblTärning5.Caption = 3 Then Counter = Counter + 3
    If Counter > 0 Then
        cmdTreor3.Caption = Counter
        txtÖvre3.Text = Counter + txtÖvre3.Text
        cmdTreor3.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en trea!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdTreor4_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 3 Then Counter = Counter + 3
    If lblTärning2.Caption = 3 Then Counter = Counter + 3
    If lblTärning3.Caption = 3 Then Counter = Counter + 3
    If lblTärning4.Caption = 3 Then Counter = Counter + 3
    If lblTärning5.Caption = 3 Then Counter = Counter + 3
    If Counter > 0 Then
        cmdTreor4.Caption = Counter
        txtÖvre4.Text = Counter + txtÖvre4.Text
        cmdTreor4.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en trea!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdTreor5_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 3 Then Counter = Counter + 3
    If lblTärning2.Caption = 3 Then Counter = Counter + 3
    If lblTärning3.Caption = 3 Then Counter = Counter + 3
    If lblTärning4.Caption = 3 Then Counter = Counter + 3
    If lblTärning5.Caption = 3 Then Counter = Counter + 3
    If Counter > 0 Then
        cmdTreor5.Caption = Counter
        txtÖvre5.Text = Counter + txtÖvre2.Text
        cmdTreor5.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en trea!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdTretal1_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning = 0         'Tretal för spelare 1
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning3) Then
        Tärning = Tärning1 + Tärning2 + Tärning3
        Tretal
        Exit Sub
    End If
    
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning2 + Tärning4
        Tretal
        Exit Sub
    End If

    If (Tärning1 = Tärning2) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning2 + Tärning5
        Tretal
        Exit Sub
    End If

    If (Tärning1 = Tärning3) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning3 + Tärning4
        Tretal
        Exit Sub
    End If

    If (Tärning1 = Tärning3) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning3 + Tärning5
        Tretal
        Exit Sub
    End If

    If (Tärning1 = Tärning4) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning4 + Tärning5
        Tretal
        Exit Sub
    End If

    If (Tärning3 = Tärning2) And (Tärning3 = Tärning5) Then
        Tärning = Tärning3 + Tärning2 + Tärning5
        Tretal
        Exit Sub
    End If
    
    If (Tärning4 = Tärning2) And (Tärning4 = Tärning5) Then
        Tärning = Tärning4 + Tärning2 + Tärning5
        Tretal
        Exit Sub
    End If
    
    If (Tärning4 = Tärning3) And (Tärning4 = Tärning5) Then
        Tärning = Tärning4 + Tärning3 + Tärning5
        Tretal
        Exit Sub
    End If
    
    If (Tärning2 = Tärning3) And (Tärning2 = Tärning4) Then
        Tärning = Tärning3 + Tärning2 + Tärning4
        Tretal
        Exit Sub
    End If
    Response = MsgBox("Du har inget tretal! Vill du strycka tretal?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmdTretal1.Enabled = False
        cmdTretal1.Caption = "-"
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmdTretal2_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning = 0         'Tretal för spelare 2
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning3) Then
        Tärning = Tärning1 + Tärning2 + Tärning3
        Tretal2
        Exit Sub
    End If
    
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning2 + Tärning4
        Tretal2
        Exit Sub
    End If

    If (Tärning1 = Tärning2) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning2 + Tärning5
        Tretal2
        Exit Sub
    End If

    If (Tärning1 = Tärning3) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning3 + Tärning4
        Tretal2
        Exit Sub
    End If

    If (Tärning1 = Tärning3) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning3 + Tärning5
        Tretal2
        Exit Sub
    End If

    If (Tärning1 = Tärning4) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning4 + Tärning5
        Tretal2
        Exit Sub
    End If

    If (Tärning3 = Tärning2) And (Tärning3 = Tärning5) Then
        Tärning = Tärning3 + Tärning2 + Tärning5
        Tretal2
        Exit Sub
    End If
    
    If (Tärning4 = Tärning2) And (Tärning4 = Tärning5) Then
        Tärning = Tärning4 + Tärning2 + Tärning5
        Tretal2
        Exit Sub
    End If
    
    If (Tärning4 = Tärning3) And (Tärning4 = Tärning5) Then
        Tärning = Tärning4 + Tärning3 + Tärning5
        Tretal2
        Exit Sub
    End If
    
    If (Tärning2 = Tärning3) And (Tärning2 = Tärning4) Then
        Tärning = Tärning3 + Tärning2 + Tärning4
        Tretal2
        Exit Sub
    End If
    Response = MsgBox("Du har inget tretal! Vill du strycka tretal?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmdTretal2.Enabled = False
        cmdTretal2.Caption = "-"
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmdTretal3_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning = 0         'Tretal för spelare 3
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning3) Then
        Tärning = Tärning1 + Tärning2 + Tärning3
        Tretal3
        Exit Sub
    End If
    
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning2 + Tärning4
        Tretal3
        Exit Sub
    End If

    If (Tärning1 = Tärning2) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning2 + Tärning5
        Tretal3
        Exit Sub
    End If

    If (Tärning1 = Tärning3) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning3 + Tärning4
        Tretal3
        Exit Sub
    End If

    If (Tärning1 = Tärning3) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning3 + Tärning5
        Tretal3
        Exit Sub
    End If

    If (Tärning1 = Tärning4) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning4 + Tärning5
        Tretal3
        Exit Sub
    End If

    If (Tärning3 = Tärning2) And (Tärning3 = Tärning5) Then
        Tärning = Tärning3 + Tärning2 + Tärning5
        Tretal3
        Exit Sub
    End If
    
    If (Tärning4 = Tärning2) And (Tärning4 = Tärning5) Then
        Tärning = Tärning4 + Tärning2 + Tärning5
        Tretal3
        Exit Sub
    End If
    
    If (Tärning4 = Tärning3) And (Tärning4 = Tärning5) Then
        Tärning = Tärning4 + Tärning3 + Tärning5
        Tretal3
        Exit Sub
    End If
    
    If (Tärning2 = Tärning3) And (Tärning2 = Tärning4) Then
        Tärning = Tärning3 + Tärning2 + Tärning4
        Tretal3
        Exit Sub
    End If
    Response = MsgBox("Du har inget tretal! Vill du strycka tretal?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmdTretal3.Enabled = False
        cmdTretal3.Caption = "-"
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmdTretal4_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning = 0         'Tretal för spelare 4
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning3) Then
        Tärning = Tärning1 + Tärning2 + Tärning3
        Tretal4
        Exit Sub
    End If
    
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning2 + Tärning4
        Tretal4
        Exit Sub
    End If

    If (Tärning1 = Tärning2) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning2 + Tärning5
        Tretal4
        Exit Sub
    End If

    If (Tärning1 = Tärning3) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning3 + Tärning4
        Tretal4
        Exit Sub
    End If

    If (Tärning1 = Tärning3) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning3 + Tärning5
        Tretal4
        Exit Sub
    End If

    If (Tärning1 = Tärning4) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning4 + Tärning5
        Tretal4
        Exit Sub
    End If

    If (Tärning3 = Tärning2) And (Tärning3 = Tärning5) Then
        Tärning = Tärning3 + Tärning2 + Tärning5
        Tretal4
        Exit Sub
    End If
    
    If (Tärning4 = Tärning2) And (Tärning4 = Tärning5) Then
        Tärning = Tärning4 + Tärning2 + Tärning5
        Tretal4
        Exit Sub
    End If
    
    If (Tärning4 = Tärning3) And (Tärning4 = Tärning5) Then
        Tärning = Tärning4 + Tärning3 + Tärning5
        Tretal4
        Exit Sub
    End If
    
    If (Tärning2 = Tärning3) And (Tärning2 = Tärning4) Then
        Tärning = Tärning3 + Tärning2 + Tärning4
        Tretal4
        Exit Sub
    End If
    Response = MsgBox("Du har inget tretal! Vill du strycka tretal?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmdTretal4.Enabled = False
        cmdTretal4.Caption = "-"
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmdTretal5_Click()
    If Fusk = 1 Then
    Fusk = 2
    Tärning = 0         'Tretal för spelare 5
    Tärning1 = lblTärning1.Caption
    Tärning2 = lblTärning2.Caption
    Tärning3 = lblTärning3.Caption
    Tärning4 = lblTärning4.Caption
    Tärning5 = lblTärning5.Caption
    
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning3) Then
        Tärning = Tärning1 + Tärning2 + Tärning3
        Tretal5
        Exit Sub
    End If
    
    If (Tärning1 = Tärning2) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning2 + Tärning4
        Tretal5
        Exit Sub
    End If

    If (Tärning1 = Tärning2) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning2 + Tärning5
        Tretal5
        Exit Sub
    End If

    If (Tärning1 = Tärning3) And (Tärning1 = Tärning4) Then
        Tärning = Tärning1 + Tärning3 + Tärning4
        Tretal5
        Exit Sub
    End If

    If (Tärning1 = Tärning3) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning3 + Tärning5
        Tretal5
        Exit Sub
    End If

    If (Tärning1 = Tärning4) And (Tärning1 = Tärning5) Then
        Tärning = Tärning1 + Tärning4 + Tärning5
        Tretal5
        Exit Sub
    End If

    If (Tärning3 = Tärning2) And (Tärning3 = Tärning5) Then
        Tärning = Tärning3 + Tärning2 + Tärning5
        Tretal5
        Exit Sub
    End If
    
    If (Tärning4 = Tärning2) And (Tärning4 = Tärning5) Then
        Tärning = Tärning4 + Tärning2 + Tärning5
        Tretal5
        Exit Sub
    End If
    
    If (Tärning4 = Tärning3) And (Tärning4 = Tärning5) Then
        Tärning = Tärning4 + Tärning3 + Tärning5
        Tretal5
        Exit Sub
    End If
    
    If (Tärning2 = Tärning3) And (Tärning2 = Tärning4) Then
        Tärning = Tärning3 + Tärning2 + Tärning4
        Tretal5
        Exit Sub
    End If
    Response = MsgBox("Du har inget tretal! Vill du strycka tretal?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdNästa.Enabled = True
        cmdKasta.Enabled = False
        cmdTretal5.Enabled = False
        cmdTretal5.Caption = "-"
    End If
    If Response = vbNo Then Fusk = 1
    End If
End Sub

Private Sub cmdTvåor1_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 2 Then Counter = Counter + 2
    If lblTärning2.Caption = 2 Then Counter = Counter + 2
    If lblTärning3.Caption = 2 Then Counter = Counter + 2
    If lblTärning4.Caption = 2 Then Counter = Counter + 2
    If lblTärning5.Caption = 2 Then Counter = Counter + 2
    If Counter > 0 Then
        cmdTvåor1.Caption = Counter
        txtÖvre1.Text = Counter + txtÖvre1.Text
        cmdTvåor1.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en tvåa!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdTvåor2_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 2 Then Counter = Counter + 2
    If lblTärning2.Caption = 2 Then Counter = Counter + 2
    If lblTärning3.Caption = 2 Then Counter = Counter + 2
    If lblTärning4.Caption = 2 Then Counter = Counter + 2
    If lblTärning5.Caption = 2 Then Counter = Counter + 2
    If Counter > 0 Then
        cmdTvåor2.Caption = Counter
        txtÖvre2.Text = Counter + txtÖvre2.Text
        cmdTvåor2.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en tvåa!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdTvåor3_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 2 Then Counter = Counter + 2
    If lblTärning2.Caption = 2 Then Counter = Counter + 2
    If lblTärning3.Caption = 2 Then Counter = Counter + 2
    If lblTärning4.Caption = 2 Then Counter = Counter + 2
    If lblTärning5.Caption = 2 Then Counter = Counter + 2
    If Counter > 0 Then
        cmdTvåor3.Caption = Counter
        txtÖvre3.Text = Counter + txtÖvre3.Text
        cmdTvåor3.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en tvåa!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdTvåor4_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 2 Then Counter = Counter + 2
    If lblTärning2.Caption = 2 Then Counter = Counter + 2
    If lblTärning3.Caption = 2 Then Counter = Counter + 2
    If lblTärning4.Caption = 2 Then Counter = Counter + 2
    If lblTärning5.Caption = 2 Then Counter = Counter + 2
    If Counter > 0 Then
        cmdTvåor4.Caption = Counter
        txtÖvre4.Text = Counter + txtÖvre4.Text
        cmdTvåor4.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en tvåa!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdTvåor5_Click()
    If Fusk = 1 Then
    Fusk = 2
    Counter = 0
    If lblTärning1.Caption = 2 Then Counter = Counter + 2
    If lblTärning2.Caption = 2 Then Counter = Counter + 2
    If lblTärning3.Caption = 2 Then Counter = Counter + 2
    If lblTärning4.Caption = 2 Then Counter = Counter + 2
    If lblTärning5.Caption = 2 Then Counter = Counter + 2
    If Counter > 0 Then
        cmdTvåor5.Caption = Counter
        txtÖvre5.Text = Counter + txtÖvre5.Text
        cmdTvåor5.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Counter = 0 Then
        MsgBox "Du måste åtminstånde ha en tvåa!", vbOKOnly, "Medelande"
        Fusk = 1
    End If
    End If
End Sub

Private Sub cmdYatzy1_Click()
    If Fusk = 1 Then
    Fusk = 2
    If lblTärning1.Caption = lblTärning2.Caption And lblTärning1.Caption = lblTärning3.Caption And lblTärning1.Caption = lblTärning4.Caption And lblTärning1.Caption = lblTärning5.Caption Then
        cmdYatzy1.Caption = 50
        txtUndre1.Text = txtUndre1.Text + 50
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
        cmdYatzy1.Enabled = False
    Else
        Response = MsgBox("Du har inte YATZY! Vill du strycka YATZY?", vbYesNo, "Medelande")
            If Response = vbYes Then
                cmdYatzy1.Caption = "-"
                cmdKasta.Enabled = False
                cmdNästa.Enabled = True
                cmdYatzy1.Enabled = False
            End If
            If Response = vbNo Then Fusk = 1
    End If
    End If
End Sub

Private Sub cmdYatzy2_Click()
    If Fusk = 1 Then
    Fusk = 2
    If lblTärning1.Caption = lblTärning2.Caption And lblTärning1.Caption = lblTärning3.Caption And lblTärning1.Caption = lblTärning4.Caption And lblTärning1.Caption = lblTärning5.Caption Then
        cmdYatzy2.Caption = 50
        txtUndre2.Text = txtUndre2.Text + 50
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
        cmdYatzy2.Enabled = False
    Else
        Response = MsgBox("Du har inte YATZY! Vill du strycka YATZY?", vbYesNo, "Medelande")
            If Response = vbYes Then
                cmdYatzy2.Caption = "-"
                cmdKasta.Enabled = False
                cmdNästa.Enabled = True
                cmdYatzy2.Enabled = False
            End If
            If Response = vbNo Then Fusk = 1
    End If
    End If
End Sub

Private Sub cmdYatzy3_Click()
    If Fusk = 1 Then
    Fusk = 2
    If lblTärning1.Caption = lblTärning2.Caption And lblTärning1.Caption = lblTärning3.Caption And lblTärning1.Caption = lblTärning4.Caption And lblTärning1.Caption = lblTärning5.Caption Then
        cmdYatzy3.Caption = 50
        txtUndre3.Text = txtUndre3.Text + 50
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
        cmdYatzy3.Enabled = False
    Else
        Response = MsgBox("Du har inte YATZY! Vill du strycka YATZY?", vbYesNo, "Medelande")
            If Response = vbYes Then
                cmdYatzy3.Caption = "-"
                cmdKasta.Enabled = False
                cmdNästa.Enabled = True
                cmdYatzy3.Enabled = False
            End If
            If Response = vbNo Then Fusk = 1
    End If
    End If
End Sub

Private Sub cmdYatzy4_Click()
    If Fusk = 1 Then
    Fusk = 2
    If lblTärning1.Caption = lblTärning2.Caption And lblTärning1.Caption = lblTärning3.Caption And lblTärning1.Caption = lblTärning4.Caption And lblTärning1.Caption = lblTärning5.Caption Then
        cmdYatzy4.Caption = 50
        txtUndre4.Text = txtUndre4.Text + 50
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
        cmdYatzy4.Enabled = False
    Else
        Response = MsgBox("Du har inte YATZY! Vill du strycka YATZY?", vbYesNo, "Medelande")
            If Response = vbYes Then
                cmdYatzy4.Caption = "-"
                cmdKasta.Enabled = False
                cmdNästa.Enabled = True
                cmdYatzy4.Enabled = False
            End If
            If Response = vbNo Then Fusk = 1
    End If
    End If
End Sub

Private Sub cmdYatzy5_Click()
    If Fusk = 1 Then
    Fusk = 2
    If lblTärning1.Caption = lblTärning2.Caption And lblTärning1.Caption = lblTärning3.Caption And lblTärning1.Caption = lblTärning4.Caption And lblTärning1.Caption = lblTärning5.Caption Then
        cmdYatzy5.Caption = 50
        txtUndre5.Text = txtUndre5.Text + 50
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
        cmdYatzy5.Enabled = False
    Else
        Response = MsgBox("Du har inte YATZY! Vill du strycka YATZY?", vbYesNo, "Medelande")
            If Response = vbYes Then
                cmdYatzy5.Caption = "-"
                cmdKasta.Enabled = False
                cmdNästa.Enabled = True
                cmdYatzy5.Enabled = False
            End If
            If Response = vbNo Then Fusk = 1
    End If
    End If
End Sub


Private Sub Form_Load()
   
    Spelare = 1
    SpelareNr = 1
    mnuEn.Checked = True
    Animering = 60
    cmdNästa.Enabled = False
    cmdKasta.Enabled = False
    Kast = 0
    txtÖvre1.Text = 0
    txtÖvre2.Text = 0
    txtÖvre3.Text = 0
    txtÖvre4.Text = 0
    txtÖvre5.Text = 0
    txtUndre1.Text = 0
    txtUndre2.Text = 0
    txtUndre3.Text = 0
    txtUndre4.Text = 0
    txtUndre5.Text = 0

    gRecordLen = Len(gScore)
    gFileNum = FreeFile
    
    Open "SCORE.DAT" For Random As gFileNum Len = gRecordLen
    gCurrentRecord = 1
    ShowScore
End Sub

Private Sub mnuAvsluta_Click()
    SaveScore
    Close #gFileNum
    End 'Avslutar programet
End Sub

Public Sub Reset() 'Reset, nytt spel
    
    cmdEttor1.Enabled = True
    cmdTvåor1.Enabled = True
    cmdTreor1.Enabled = True
    cmdFyror1.Enabled = True
    cmdFemmor1.Enabled = True
    cmdSexor1.Enabled = True
    cmd1Par1.Enabled = True
    cmd2Par1.Enabled = True
    cmdTretal1.Enabled = True
    cmdFyrtal1.Enabled = True
    cmdLiten1.Enabled = True
    cmdStor1.Enabled = True
    cmdKåk1.Enabled = True
    cmdChans1.Enabled = True
    cmdYatzy1.Enabled = True
    
    cmdEttor2.Enabled = True
    cmdTvåor2.Enabled = True
    cmdTreor2.Enabled = True
    cmdFyror2.Enabled = True
    cmdFemmor2.Enabled = True
    cmdSexor2.Enabled = True
    cmd1Par2.Enabled = True
    cmd2Par2.Enabled = True
    cmdTretal2.Enabled = True
    cmdFyrtal2.Enabled = True
    cmdLiten2.Enabled = True
    cmdStor2.Enabled = True
    cmdKåk2.Enabled = True
    cmdChans2.Enabled = True
    cmdYatzy2.Enabled = True
    
    cmdEttor3.Enabled = True
    cmdTvåor3.Enabled = True
    cmdTreor3.Enabled = True
    cmdFyror3.Enabled = True
    cmdFemmor3.Enabled = True
    cmdSexor3.Enabled = True
    cmd1Par3.Enabled = True
    cmd2Par3.Enabled = True
    cmdTretal3.Enabled = True
    cmdFyrtal3.Enabled = True
    cmdLiten3.Enabled = True
    cmdStor3.Enabled = True
    cmdKåk3.Enabled = True
    cmdChans3.Enabled = True
    cmdYatzy3.Enabled = True
    
    cmdEttor4.Enabled = True
    cmdTvåor4.Enabled = True
    cmdTreor4.Enabled = True
    cmdFyror4.Enabled = True
    cmdFemmor4.Enabled = True
    cmdSexor4.Enabled = True
    cmd1Par4.Enabled = True
    cmd2Par4.Enabled = True
    cmdTretal4.Enabled = True
    cmdFyrtal4.Enabled = True
    cmdLiten4.Enabled = True
    cmdStor4.Enabled = True
    cmdKåk4.Enabled = True
    cmdChans4.Enabled = True
    cmdYatzy4.Enabled = True
    
    cmdEttor5.Enabled = True
    cmdTvåor5.Enabled = True
    cmdTreor5.Enabled = True
    cmdFyror5.Enabled = True
    cmdFemmor5.Enabled = True
    cmdSexor5.Enabled = True
    cmd1Par5.Enabled = True
    cmd2Par5.Enabled = True
    cmdTretal5.Enabled = True
    cmdFyrtal5.Enabled = True
    cmdLiten5.Enabled = True
    cmdStor5.Enabled = True
    cmdKåk5.Enabled = True
    cmdChans5.Enabled = True
    cmdYatzy5.Enabled = True
    
    
    cmdEttor1.Caption = ""
    cmdTvåor1.Caption = ""
    cmdTreor1.Caption = ""
    cmdFyror1.Caption = ""
    cmdFemmor1.Caption = ""
    cmdSexor1.Caption = ""
    cmdSumma1.Caption = ""
    cmdBonus1.Caption = ""
    cmd1Par1.Caption = ""
    cmd2Par1.Caption = ""
    cmdTretal1.Caption = ""
    cmdFyrtal1.Caption = ""
    cmdLiten1.Caption = ""
    cmdStor1.Caption = ""
    cmdKåk1.Caption = ""
    cmdChans1.Caption = ""
    cmdYatzy1.Caption = ""
    cmd2Summa1.Caption = ""
    cmdEttor2.Caption = ""
    cmdTvåor2.Caption = ""
    cmdTreor2.Caption = ""
    cmdFyror2.Caption = ""
    cmdFemmor2.Caption = ""
    cmdSexor2.Caption = ""
    cmdSumma2.Caption = ""
    cmdBonus2.Caption = ""
    cmd1Par2.Caption = ""
    cmd2Par2.Caption = ""
    cmdTretal2.Caption = ""
    cmdFyrtal2.Caption = ""
    cmdLiten2.Caption = ""
    cmdStor2.Caption = ""
    cmdKåk2.Caption = ""
    cmdChans2.Caption = ""
    cmdYatzy2.Caption = ""
    cmd2Summa2.Caption = ""
    cmdEttor3.Caption = ""
    cmdTvåor3.Caption = ""
    cmdTreor3.Caption = ""
    cmdFyror3.Caption = ""
    cmdFemmor3.Caption = ""
    cmdSexor3.Caption = ""
    cmdSumma3.Caption = ""
    cmdBonus3.Caption = ""
    cmd1Par3.Caption = ""
    cmd2Par3.Caption = ""
    cmdTretal3.Caption = ""
    cmdFyrtal3.Caption = ""
    cmdLiten3.Caption = ""
    cmdStor3.Caption = ""
    cmdKåk3.Caption = ""
    cmdChans3.Caption = ""
    cmdYatzy3.Caption = ""
    cmd2Summa3.Caption = ""
    cmdEttor4.Caption = ""
    cmdTvåor4.Caption = ""
    cmdTreor4.Caption = ""
    cmdFyror4.Caption = ""
    cmdFemmor4.Caption = ""
    cmdSexor4.Caption = ""
    cmdSumma4.Caption = ""
    cmdBonus4.Caption = ""
    cmd1Par4.Caption = ""
    cmd2Par4.Caption = ""
    cmdTretal4.Caption = ""
    cmdFyrtal4.Caption = ""
    cmdLiten4.Caption = ""
    cmdStor4.Caption = ""
    cmdKåk4.Caption = ""
    cmdChans4.Caption = ""
    cmdYatzy4.Caption = ""
    cmd2Summa4.Caption = ""
    cmdEttor5.Caption = ""
    cmdTvåor5.Caption = ""
    cmdTreor5.Caption = ""
    cmdFyror5.Caption = ""
    cmdFemmor5.Caption = ""
    cmdSexor5.Caption = ""
    cmdSumma5.Caption = ""
    cmdBonus5.Caption = ""
    cmd1Par5.Caption = ""
    cmd2Par5.Caption = ""
    cmdTretal5.Caption = ""
    cmdFyrtal5.Caption = ""
    cmdLiten5.Caption = ""
    cmdStor5.Caption = ""
    cmdKåk5.Caption = ""
    cmdChans5.Caption = ""
    cmdYatzy5.Caption = ""
    cmd2Summa5.Caption = ""
    cmdNästa.Enabled = False
    cmdKasta.Enabled = True
    cmdKlara.Enabled = True
    Fusk = 1
    chk1.Value = 0
    chk2.Value = 0
    chk3.Value = 0
    chk4.Value = 0
    chk5.Value = 0
    imgVisa(0).Visible = False
    imgVisa(1).Visible = False
    imgVisa(2).Visible = False
    imgVisa(3).Visible = False
    imgVisa(4).Visible = False
    imgVisa2(0).Visible = False
    imgVisa2(1).Visible = False
    imgVisa2(2).Visible = False
    imgVisa2(3).Visible = False
    imgVisa2(4).Visible = False
    Kast = 0
    txtÖvre1.Text = "0"
    txtÖvre2.Text = "0"
    txtÖvre3.Text = "0"
    txtÖvre4.Text = "0"
    txtÖvre5.Text = "0"
    txtUndre1.Text = "0"
    txtUndre2.Text = "0"
    txtUndre3.Text = "0"
    txtUndre4.Text = "0"
    txtUndre5.Text = "0"
End Sub

Private Sub mnuEn_Click()
    Reset
    Spelare = 1
    mnuEn.Checked = True
    mnutvå.Checked = False
    mnuTre.Checked = False
    mnuFyra.Checked = False
    mnuFem.Checked = False
    frmNamn.opt1.Value = True
    Namn
End Sub

Private Sub mnuFem_Click()
    Reset
    Spelare = 5
    mnuEn.Checked = False
    mnutvå.Checked = False
    mnuTre.Checked = False
    mnuFyra.Checked = False
    mnuFem.Checked = True
    frmNamn.opt5.Value = True
    Namn
End Sub

Private Sub mnuFyra_Click()
    Reset
    Spelare = 4
    mnuEn.Checked = False
    mnutvå.Checked = False
    mnuTre.Checked = False
    mnuFyra.Checked = True
    mnuFem.Checked = False
    frmNamn.opt4.Value = True
    Namn
End Sub

Private Sub mnuNytt_Click()
    Reset
    Namn
End Sub

Private Sub mnuOm_Click()
    frmAbout.Show
End Sub

Private Sub mnuPrint_Click()
    'Print Results
    Printer.Print
    Printer.Print "Super YATZY '98 - Resultat. " + Time$; ", " + Date$
    Printer.Print
    Printer.Print lblSpelare1.Caption; "                            "; lblSpelare2.Caption; "                      " + lblSpelare3.Caption
    Printer.Print
    Printer.Print "Ettor:       " + cmdEttor1.Caption + "                  " + cmdEttor2.Caption + "                        " + cmdEttor3.Caption
    Printer.Print "Tvåor:       " + cmdTvåor1.Caption + "                  " + cmdTvåor2.Caption + "                        " + cmdTvåor3.Caption
    Printer.Print "Treor:       " + cmdTreor1.Caption + "                  " + cmdTreor2.Caption + "                        " + cmdTreor3.Caption
    Printer.Print "Fyror:       " + cmdFyror1.Caption + "                 " + cmdFyror2.Caption + "                       " + cmdFyror3.Caption
    Printer.Print "Femmor:      " + cmdFemmor1.Caption + "                 " + cmdFemmor2.Caption + "                       " + cmdFemmor3.Caption
    Printer.Print "Sexor:       " + cmdSexor1.Caption + "                 " + cmdSexor2.Caption + "                       " + cmdSexor3.Caption
    Printer.Print "Summa:       " + cmdSumma1.Caption + "                 " + cmdSumma2.Caption + "                       " + cmdSumma3.Caption
    Printer.Print "Bonus:       " + cmdBonus1.Caption + "                 " + cmdBonus2.Caption + "                       " + cmdBonus3.Caption
    Printer.Print "1 Par:       " + cmd1Par1.Caption + "                 " + cmd1Par2.Caption + "                       " + cmd1Par3.Caption
    Printer.Print "2 Par:       " + cmd2Par1.Caption + "                 " + cmd2Par2.Caption + "                       " + cmd2Par3.Caption
    Printer.Print "Tretal:      " + cmdTretal1.Caption + "                 " + cmdTretal2.Caption + "                       " + cmdTretal3.Caption
    Printer.Print "Fyrtal:      " + cmdFyrtal1.Caption + "                 " + cmdFyrtal2.Caption + "                       " + cmdFyrtal3.Caption
    Printer.Print "Liten stege: " + cmdLiten1.Caption + "                 " + cmdLiten2.Caption + "                       " + cmdLiten3.Caption
    Printer.Print "Stor stege:  " + cmdStor1.Caption + "                 " + cmdStor2.Caption + "                       " + cmdStor3.Caption
    Printer.Print "Kåk:         " + cmdKåk1.Caption + "                 " + cmdKåk2.Caption + "                       " + cmdKåk3.Caption
    Printer.Print "Chans:       " + cmdChans1.Caption + "                 " + cmdChans2.Caption + "                       " + cmdChans3.Caption
    Printer.Print "YATZY:       " + cmdYatzy1.Caption + "                 " + cmdYatzy2.Caption + "                       " + cmdYatzy3.Caption
    Printer.Print "Summa:       " + cmd2Summa1.Caption + "                " + cmd2Summa2.Caption + "                      " + cmd2Summa3.Caption
    
    Printer.Print ""
    
    Printer.Print lblSpelare4.Caption; "                            " + lblSpelare5.Caption
    Printer.Print
    Printer.Print "Ettor:       " + cmdEttor4.Caption + "                  " + cmdEttor5.Caption
    Printer.Print "Tvåor:       " + cmdTvåor4.Caption + "                  " + cmdTvåor5.Caption
    Printer.Print "Treor:       " + cmdTreor4.Caption + "                  " + cmdTreor5.Caption
    Printer.Print "Fyror:       " + cmdFyror4.Caption + "                 " + cmdFyror5.Caption
    Printer.Print "Femmor:      " + cmdFemmor4.Caption + "                 " + cmdFemmor5.Caption
    Printer.Print "Sexor:       " + cmdSexor4.Caption + "                 " + cmdSexor5.Caption
    Printer.Print "Summa:       " + cmdSumma4.Caption + "                 " + cmdSumma5.Caption
    Printer.Print "Bonus:       " + cmdBonus4.Caption + "                 " + cmdBonus5.Caption
    Printer.Print "1 Par:       " + cmd1Par4.Caption + "                 " + cmd1Par5.Caption
    Printer.Print "2 Par:       " + cmd2Par4.Caption + "                 " + cmd2Par5.Caption
    Printer.Print "Tretal:      " + cmdTretal4.Caption + "                 " + cmdTretal5.Caption
    Printer.Print "Fyrtal:      " + cmdFyrtal4.Caption + "                 " + cmdFyrtal5.Caption
    Printer.Print "Liten stege: " + cmdLiten4.Caption + "                 " + cmdLiten5.Caption
    Printer.Print "Stor stege:  " + cmdStor4.Caption + "                 " + cmdStor5.Caption
    Printer.Print "Kåk:         " + cmdKåk4.Caption + "                 " + cmdKåk5.Caption
    Printer.Print "Chans:       " + cmdChans4.Caption + "                 " + cmdChans5.Caption
    Printer.Print "YATZY:       " + cmdYatzy4.Caption + "                 " + cmdYatzy5.Caption
    Printer.Print "Summa:       " + cmd2Summa4.Caption + "                " + cmd2Summa5.Caption

    Printer.EndDoc
End Sub

Private Sub mnuScore_Click()
    frmScore.txtScore1.Text = frmYatzy.txtScore1.Text
    frmScore.txtName1.Text = frmYatzy.txtName1.Text
    frmScore.txtDate1.Text = frmYatzy.txtDate1.Text
    frmScore.txtScore2.Text = frmYatzy.txtScore2.Text
    frmScore.txtName2.Text = frmYatzy.txtName2.Text
    frmScore.txtDate2.Text = frmYatzy.txtDate2.Text
    frmScore.txtScore3.Text = frmYatzy.txtScore3.Text
    frmScore.txtName3.Text = frmYatzy.txtName3.Text
    frmScore.txtDate3.Text = frmYatzy.txtDate3.Text
    frmScore.Show
End Sub

Private Sub mnuTre_Click()
    Reset
    Spelare = 3
    mnuEn.Checked = False
    mnutvå.Checked = False
    mnuTre.Checked = True
    mnuFyra.Checked = False
    mnuFem.Checked = False
    frmNamn.opt3.Value = True
    Namn
End Sub

Private Sub mnutvå_Click()
    Reset
    Spelare = 2
    mnuEn.Checked = False
    mnutvå.Checked = True
    mnuTre.Checked = False
    mnuFyra.Checked = False
    mnuFem.Checked = False
    frmNamn.opt2.Value = True
    Namn
End Sub

Private Sub trmTimer_Timer()
    StatusBar1.SimpleText = "Tid: " + Time$ + ", Datum: " + Date$ + ", Spelare: Viktor"
    If Animering < 51 Then
       Animering = Animering + 1
       imgVisa(Visa).Picture = imgTär(Current).Picture
       Current = Current + 1
       Visa = Visa + 1
       If Current > 5 Then Current = 0
       If Visa > 4 Then Visa = 0
    Else
      Exit Sub
    End If
    If Animering = 51 Then
        
        imgVisa2(0).Visible = True
        imgVisa2(1).Visible = True
        imgVisa2(2).Visible = True
        imgVisa2(3).Visible = True
        imgVisa2(4).Visible = True
        
        If lblTärning1.Caption = 1 Then imgVisa2(3).Picture = imgTär(0).Picture
        If lblTärning1.Caption = 2 Then imgVisa2(3).Picture = imgTär(1).Picture
        If lblTärning1.Caption = 3 Then imgVisa2(3).Picture = imgTär(3).Picture
        If lblTärning1.Caption = 4 Then imgVisa2(3).Picture = imgTär(5).Picture
        If lblTärning1.Caption = 5 Then imgVisa2(3).Picture = imgTär(2).Picture
        If lblTärning1.Caption = 6 Then imgVisa2(3).Picture = imgTär(4).Picture
      
        If lblTärning2.Caption = 1 Then imgVisa2(2).Picture = imgTär(0).Picture
        If lblTärning2.Caption = 2 Then imgVisa2(2).Picture = imgTär(1).Picture
        If lblTärning2.Caption = 3 Then imgVisa2(2).Picture = imgTär(3).Picture
        If lblTärning2.Caption = 4 Then imgVisa2(2).Picture = imgTär(5).Picture
        If lblTärning2.Caption = 5 Then imgVisa2(2).Picture = imgTär(2).Picture
        If lblTärning2.Caption = 6 Then imgVisa2(2).Picture = imgTär(4).Picture
      
        If lblTärning3.Caption = 1 Then imgVisa2(4).Picture = imgTär(0).Picture
        If lblTärning3.Caption = 2 Then imgVisa2(4).Picture = imgTär(1).Picture
        If lblTärning3.Caption = 3 Then imgVisa2(4).Picture = imgTär(3).Picture
        If lblTärning3.Caption = 4 Then imgVisa2(4).Picture = imgTär(5).Picture
        If lblTärning3.Caption = 5 Then imgVisa2(4).Picture = imgTär(2).Picture
        If lblTärning3.Caption = 6 Then imgVisa2(4).Picture = imgTär(4).Picture
      
        If lblTärning4.Caption = 1 Then imgVisa2(0).Picture = imgTär(0).Picture
        If lblTärning4.Caption = 2 Then imgVisa2(0).Picture = imgTär(1).Picture
        If lblTärning4.Caption = 3 Then imgVisa2(0).Picture = imgTär(3).Picture
        If lblTärning4.Caption = 4 Then imgVisa2(0).Picture = imgTär(5).Picture
        If lblTärning4.Caption = 5 Then imgVisa2(0).Picture = imgTär(2).Picture
        If lblTärning4.Caption = 6 Then imgVisa2(0).Picture = imgTär(4).Picture
      
        If lblTärning5.Caption = 1 Then imgVisa2(1).Picture = imgTär(0).Picture
        If lblTärning5.Caption = 2 Then imgVisa2(1).Picture = imgTär(1).Picture
        If lblTärning5.Caption = 3 Then imgVisa2(1).Picture = imgTär(3).Picture
        If lblTärning5.Caption = 4 Then imgVisa2(1).Picture = imgTär(5).Picture
        If lblTärning5.Caption = 5 Then imgVisa2(1).Picture = imgTär(2).Picture
        If lblTärning5.Caption = 6 Then imgVisa2(1).Picture = imgTär(4).Picture
        Animering = 71
    End If
End Sub

Public Sub ParSpelare1()
    txtUndre1.Text = txtUndre1.Text + Tärning
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
    cmd2Par1.Enabled = False
    cmd2Par1.Caption = Tärning
End Sub

Public Sub Tretal()
    txtUndre1.Text = txtUndre1.Text + Tärning
    cmdTretal1.Caption = Tärning
    cmdTretal1.Enabled = False
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
End Sub

Public Sub Fyrtal()
    txtUndre1.Text = txtUndre1.Text + Tärning
    cmdFyrtal1.Caption = Tärning
    cmdFyrtal1.Enabled = False
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
End Sub

Public Sub Kåk()
    txtUndre1.Text = txtUndre1.Text + Tärning
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
    cmdKåk1.Enabled = False
    cmdKåk1.Caption = Tärning
End Sub

Public Sub Liten()
    Response = MsgBox("Du har inte en liten stege! Vill du strycka Liten stege?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdLiten1.Caption = "-"
        cmdLiten1.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Response = vbNo Then Fusk = 1
    Exit Sub
End Sub

Public Sub Stor()
    Response = MsgBox("Du har inte en Stor Stege! Vill du strycka Stor Stege?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdStor1.Caption = "-"
        cmdStor1.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Response = vbNo Then Fusk = 1
    Exit Sub
End Sub

Public Sub ParSpelare2()
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
    cmd2Par2.Enabled = False
    cmd2Par2.Caption = Tärning
    txtUndre2.Text = txtUndre2.Text + Tärning
End Sub

Public Sub ParSpelare3()
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
    cmd2Par3.Enabled = False
    cmd2Par3.Caption = Tärning
    txtUndre3.Text = txtUndre3.Text + Tärning
End Sub

Public Sub ParSpelare4()
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
    cmd2Par4.Enabled = False
    cmd2Par4.Caption = Tärning
    txtUndre4.Text = txtUndre4.Text + Tärning
End Sub

Public Sub ParSpelare5()
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
    cmd2Par5.Enabled = False
    cmd2Par5.Caption = Tärning
    txtUndre5.Text = txtUndre5.Text + Tärning
End Sub

Public Sub Tretal2()
    txtUndre2.Text = txtUndre2.Text + Tärning
    cmdTretal2.Caption = Tärning
    cmdTretal2.Enabled = False
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
End Sub

Public Sub Tretal3()
    txtUndre3.Text = txtUndre3.Text + Tärning
    cmdTretal3.Caption = Tärning
    cmdTretal3.Enabled = False
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
End Sub

Public Sub Tretal4()
    txtUndre4.Text = txtUndre4.Text + Tärning
    cmdTretal4.Caption = Tärning
    cmdTretal4.Enabled = False
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
End Sub

Public Sub Tretal5()
    txtUndre5.Text = txtUndre5.Text + Tärning
    cmdTretal5.Caption = Tärning
    cmdTretal5.Enabled = False
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
End Sub

Public Sub Fyrtal2()
    txtUndre2.Text = txtUndre2.Text + Tärning
    cmdFyrtal2.Caption = Tärning
    cmdFyrtal2.Enabled = False
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
End Sub

Public Sub Fyrtal3()
    txtUndre3.Text = txtUndre3.Text + Tärning
    cmdFyrtal3.Caption = Tärning
    cmdFyrtal3.Enabled = False
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
End Sub

Public Sub Fyrtal5()
    txtUndre5.Text = txtUndre5.Text + Tärning
    cmdFyrtal5.Caption = Tärning
    cmdFyrtal5.Enabled = False
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
End Sub

Public Sub Fyrtal4()
    txtUndre4.Text = txtUndre4.Text + Tärning
    cmdFyrtal4.Caption = Tärning
    cmdFyrtal4.Enabled = False
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
End Sub

Public Sub Liten2()
    Response = MsgBox("Du har inte en liten stege! Vill du strycka Liten stege?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdLiten2.Caption = "-"
        cmdLiten2.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Response = vbNo Then Fusk = 1
    Exit Sub
End Sub

Public Sub Liten3()
    Response = MsgBox("Du har inte en liten stege! Vill du strycka Liten stege?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdLiten3.Caption = "-"
        cmdLiten3.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Response = vbNo Then Fusk = 1
    Exit Sub
End Sub

Public Sub Liten4()
    Response = MsgBox("Du har inte en liten stege! Vill du strycka Liten stege?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdLiten4.Caption = "-"
        cmdLiten4.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Response = vbNo Then Fusk = 1
    Exit Sub
End Sub

Public Sub Liten5()
    Response = MsgBox("Du har inte en liten stege! Vill du strycka Liten stege?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdLiten5.Caption = "-"
        cmdLiten5.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Response = vbNo Then Fusk = 1
    Exit Sub
End Sub

Public Sub Kåk2()
    txtUndre2.Text = txtUndre2.Text + Tärning
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
    cmdKåk2.Enabled = False
    cmdKåk2.Caption = Tärning
End Sub
Public Sub Kåk3()
    txtUndre3.Text = txtUndre3.Text + Tärning
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
    cmdKåk3.Enabled = False
    cmdKåk3.Caption = Tärning
End Sub
Public Sub Kåk4()
    txtUndre4.Text = txtUndre4.Text + Tärning
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
    cmdKåk4.Enabled = False
    cmdKåk4.Caption = Tärning
End Sub
Public Sub Kåk5()
    txtUndre5.Text = txtUndre5.Text + Tärning
    cmdNästa.Enabled = True
    cmdKasta.Enabled = False
    cmdKåk5.Enabled = False
    cmdKåk5.Caption = Tärning
End Sub

Public Sub Stor2()
    Response = MsgBox("Du har inte en Stor Stege! Vill du strycka Stor Stege?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdStor2.Caption = "-"
        cmdStor2.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Response = vbNo Then Fusk = 1
    Exit Sub
End Sub
Public Sub Stor3()
    Response = MsgBox("Du har inte en Stor Stege! Vill du strycka Stor Stege?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdStor3.Caption = "-"
        cmdStor3.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Response = vbNo Then Fusk = 1
    Exit Sub
End Sub
Public Sub Stor4()
    Response = MsgBox("Du har inte en Stor Stege! Vill du strycka Stor Stege?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdStor4.Caption = "-"
        cmdStor4.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Response = vbNo Then Fusk = 1
    Exit Sub
End Sub
Public Sub Stor5()
    Response = MsgBox("Du har inte en Stor Stege! Vill du strycka Stor Stege?", vbYesNo, "Medelande")
    If Response = vbYes Then
        cmdStor5.Caption = "-"
        cmdStor5.Enabled = False
        cmdKasta.Enabled = False
        cmdNästa.Enabled = True
    End If
    If Response = vbNo Then Fusk = 1
    Exit Sub
End Sub

Public Sub Namn()
    If mnuEn.Checked = True Then
        frmNamn.txtEn.Enabled = True
        frmNamn.txtTvå.Enabled = False
        frmNamn.txtTre.Enabled = False
        frmNamn.txtFyra.Enabled = False
        frmNamn.txtFem.Enabled = False
    End If
    If mnutvå.Checked = True Then
        frmNamn.txtEn.Enabled = True
        frmNamn.txtTvå.Enabled = True
        frmNamn.txtTre.Enabled = False
        frmNamn.txtFyra.Enabled = False
        frmNamn.txtFem.Enabled = False
    End If
    If mnuTre.Checked = True Then
        frmNamn.txtEn.Enabled = True
        frmNamn.txtTvå.Enabled = True
        frmNamn.txtTre.Enabled = True
        frmNamn.txtFyra.Enabled = False
        frmNamn.txtFem.Enabled = False
    End If
    If mnuFyra.Checked = True Then
        frmNamn.txtEn.Enabled = True
        frmNamn.txtTvå.Enabled = True
        frmNamn.txtTre.Enabled = True
        frmNamn.txtFyra.Enabled = True
        frmNamn.txtFem.Enabled = False
    End If
    If mnuFem.Checked = True Then
        frmNamn.txtEn.Enabled = True
        frmNamn.txtTvå.Enabled = True
        frmNamn.txtTre.Enabled = True
        frmNamn.txtFyra.Enabled = True
        frmNamn.txtFem.Enabled = True
    End If
    frmNamn.Show
End Sub

Public Sub SaveScore()
    gCurrentRecord = 1
    gScore.Score1 = txtScore1.Text
    gScore.Name1 = txtName1.Text
    gScore.Date1 = txtDate1.Text
    gScore.Score2 = txtScore2.Text
    gScore.Name2 = txtName2.Text
    gScore.Date2 = txtDate2.Text
    gScore.Score3 = txtScore3.Text
    gScore.Name3 = txtName3.Text
    gScore.Date3 = txtDate3.Text
    
    Put #gFileNum, gCurrentRecord, gScore
End Sub

Public Sub ShowScore()
    gCurrentRecord = 1
    Get #gFileNum, gCurrentRecord, gScore
    
    txtScore1.Text = Trim(gScore.Score1)
    txtName1.Text = Trim(gScore.Name1)
    txtDate1.Text = Trim(gScore.Date1)
    txtScore2.Text = Trim(gScore.Score2)
    txtName2.Text = Trim(gScore.Name2)
    txtDate2.Text = Trim(gScore.Date2)
    txtScore3.Text = Trim(gScore.Score3)
    txtName3.Text = Trim(gScore.Name3)
    txtDate3.Text = Trim(gScore.Date3)
End Sub

