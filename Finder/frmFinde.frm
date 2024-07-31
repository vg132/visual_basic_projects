VERSION 5.00
Begin VB.Form frmFinde 
   Caption         =   "Search"
   ClientHeight    =   3600
   ClientLeft      =   3765
   ClientTop       =   3510
   ClientWidth     =   3900
   Icon            =   "frmFinde.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   3900
   Begin VB.Frame Frame 
      Caption         =   "Options"
      Height          =   615
      Left            =   0
      TabIndex        =   93
      Top             =   0
      Width           =   3855
      Begin VB.OptionButton optSamling 
         Caption         =   "Collection CD's"
         Height          =   255
         Left            =   2280
         TabIndex        =   96
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optEnsam 
         Caption         =   "Ordinary CD"
         Height          =   255
         Left            =   1080
         TabIndex        =   95
         Top             =   240
         Width           =   1155
      End
      Begin VB.OptionButton optBåda 
         Caption         =   "Both"
         Height          =   255
         Left            =   240
         TabIndex        =   94
         Top             =   240
         Value           =   -1  'True
         Width           =   700
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "sammling"
      Top             =   5880
      Width           =   1140
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CD Title"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   61
      Left            =   960
      MaxLength       =   50
      TabIndex        =   92
      Top             =   5880
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 1"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   60
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   91
      Top             =   5760
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 1"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   59
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   90
      Top             =   5760
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 2"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   58
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   89
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 2"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   57
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   88
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 3"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   56
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   87
      Top             =   5760
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 3"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   55
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   86
      Top             =   5760
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 4"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   54
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   85
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 4"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   53
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   84
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 5"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   52
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   83
      Top             =   6000
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 5"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   51
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   82
      Top             =   6000
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 6"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   50
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   81
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 6"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   49
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   80
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 7"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   48
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   79
      Top             =   5760
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 7"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   47
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   78
      Top             =   5760
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 8"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   46
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   77
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 8"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   45
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   76
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 9"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   44
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   75
      Top             =   5760
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 9"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   43
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   74
      Top             =   5760
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 10"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   42
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   73
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 10"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   41
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   72
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 11"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   40
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   71
      Top             =   6000
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 11"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   39
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   70
      Top             =   6000
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 12"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   38
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   69
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 12"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   37
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   68
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 13"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   36
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   67
      Top             =   6000
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 13"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   35
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   66
      Top             =   6000
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 14"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   960
      MaxLength       =   50
      TabIndex        =   65
      Top             =   5760
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 14"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   29
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   64
      Top             =   5760
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 15"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   30
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   63
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 15"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   31
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   62
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 16"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   32
      Left            =   960
      MaxLength       =   50
      TabIndex        =   61
      Top             =   5760
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 16"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   33
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   60
      Top             =   5760
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 17"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   34
      Left            =   960
      MaxLength       =   50
      TabIndex        =   59
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "track 17"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   62
      Left            =   1080
      TabIndex        =   58
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 18"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   72
      Left            =   1080
      TabIndex        =   57
      Top             =   6000
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 18"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   63
      Left            =   1200
      TabIndex        =   56
      Top             =   6000
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "artist 19"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   73
      Left            =   960
      TabIndex        =   55
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 19"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   64
      Left            =   1080
      TabIndex        =   54
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 20"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   74
      Left            =   1080
      TabIndex        =   53
      Top             =   5760
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 20"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   65
      Left            =   1200
      TabIndex        =   52
      Top             =   5760
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 21"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   75
      Left            =   1200
      TabIndex        =   51
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 21"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   66
      Left            =   1320
      TabIndex        =   50
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 22"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   76
      Left            =   1080
      TabIndex        =   49
      Top             =   5760
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 22"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   67
      Left            =   1200
      TabIndex        =   48
      Top             =   5760
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 23"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   77
      Left            =   1080
      TabIndex        =   47
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 23"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   68
      Left            =   1200
      TabIndex        =   46
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 24"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   78
      Left            =   1200
      TabIndex        =   45
      Top             =   6000
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 24"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   69
      Left            =   1320
      TabIndex        =   44
      Top             =   6000
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 25"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   79
      Left            =   1080
      TabIndex        =   43
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 25"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   70
      Left            =   1200
      TabIndex        =   42
      Top             =   5880
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 26"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   80
      Left            =   1200
      TabIndex        =   41
      Top             =   6000
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 26"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   71
      Left            =   1320
      TabIndex        =   40
      Top             =   6000
      Width           =   2900
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Command10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   1920
      TabIndex        =   39
      Top             =   3000
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Command9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   720
      TabIndex        =   38
      Top             =   3000
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Command8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   2520
      TabIndex        =   37
      Top             =   2520
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Command7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   1320
      TabIndex        =   36
      Top             =   2520
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Command6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   35
      Top             =   2520
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Command5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2520
      TabIndex        =   34
      Top             =   2040
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Command4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1320
      TabIndex        =   33
      Top             =   2040
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Command3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   32
      Top             =   2040
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Command2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   31
      Top             =   1560
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   30
      Top             =   1560
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   0
      MaxLength       =   50
      TabIndex        =   29
      Top             =   720
      Width           =   3855
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   28
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "samma artist"
      Top             =   6600
      Width           =   1140
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CD Title"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   960
      MaxLength       =   50
      TabIndex        =   27
      Top             =   6480
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   960
      MaxLength       =   50
      TabIndex        =   26
      Top             =   6360
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 1"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   960
      MaxLength       =   50
      TabIndex        =   25
      Top             =   6765
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 2"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   960
      MaxLength       =   50
      TabIndex        =   24
      Top             =   6840
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 3"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   960
      MaxLength       =   50
      TabIndex        =   23
      Top             =   6795
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 4"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   960
      MaxLength       =   50
      TabIndex        =   22
      Top             =   6885
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 5"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   960
      MaxLength       =   50
      TabIndex        =   21
      Top             =   6600
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 6"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   960
      MaxLength       =   50
      TabIndex        =   20
      Top             =   6915
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 7"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   840
      MaxLength       =   50
      TabIndex        =   19
      Top             =   6405
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 8"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   840
      MaxLength       =   50
      TabIndex        =   18
      Top             =   6720
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 9"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   840
      MaxLength       =   50
      TabIndex        =   17
      Top             =   6795
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 10"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   840
      MaxLength       =   50
      TabIndex        =   16
      Top             =   6765
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 11"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   840
      MaxLength       =   50
      TabIndex        =   15
      Top             =   6840
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 12"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   840
      MaxLength       =   50
      TabIndex        =   14
      Top             =   6555
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 13"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   840
      MaxLength       =   50
      TabIndex        =   13
      Top             =   6885
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 14"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   960
      MaxLength       =   50
      TabIndex        =   12
      Top             =   6720
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 15"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   960
      MaxLength       =   50
      TabIndex        =   11
      Top             =   6435
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 16"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   18
      Left            =   960
      MaxLength       =   50
      TabIndex        =   10
      Top             =   6765
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 17"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   19
      Left            =   960
      MaxLength       =   50
      TabIndex        =   9
      Top             =   6840
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 18"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   20
      Left            =   960
      MaxLength       =   50
      TabIndex        =   8
      Top             =   6555
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 19"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   21
      Left            =   960
      MaxLength       =   50
      TabIndex        =   7
      Top             =   6645
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 20"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   22
      Left            =   840
      MaxLength       =   50
      TabIndex        =   6
      Top             =   6360
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 21"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   23
      Left            =   840
      MaxLength       =   50
      TabIndex        =   5
      Top             =   6675
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 22"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   24
      Left            =   840
      MaxLength       =   50
      TabIndex        =   4
      Top             =   6765
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 23"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   25
      Left            =   840
      MaxLength       =   50
      TabIndex        =   3
      Top             =   6720
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 24"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   26
      Left            =   840
      MaxLength       =   50
      TabIndex        =   2
      Top             =   6795
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 25"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   27
      Left            =   840
      MaxLength       =   50
      TabIndex        =   1
      Top             =   6525
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 26"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   28
      Left            =   840
      MaxLength       =   50
      TabIndex        =   0
      Top             =   6600
      Width           =   3375
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmFinde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Finde As String
Dim X As Integer
Dim Y As Integer
Dim Hittat As Integer
Dim Found As String
Dim Counter As Integer
Dim Opt As Integer
Dim myDB As Database
Dim myRS As Recordset
'Denna del av programet är såvitt jag vet klar och fungerar, ej

Private Sub cmdFind_Click()
    On Error GoTo errorstopp
    Hittat = 0
    Do Until Hittat = 10
        cmdGoto(Hittat).Visible = False
        Hittat = Hittat + 1
    Loop
    Hittat = 0
    Finde = txtFind.Text
    Found = String(19, " ")
    Found = ""
    Y = 1
    X = 0
    'Sök skivor med samma artist, ändast om ensama eller båda är valda
    If (optBåda.Value = True) Or (optEnsam.Value = True) Then
    Data1.Recordset.MoveFirst
    Do Until Y > Data1.Recordset.RecordCount
        Do Until X = 28
        X = X + 1
        If UCase(txtFields(X).Text) = UCase(Finde) Then
            Found = txtFields(1).Text
            cmdGoto(Hittat).Visible = True
            cmdGoto(Hittat).Caption = Found
            cmdGoto(Hittat).Tag = Y - 1
            cmdGoto(Hittat).ToolTipText = "Ordinary CD"
            Found = ""
            Hittat = Hittat + 1
        End If
        Loop
        Data1.Recordset.MoveNext
        Y = Y + 1
        X = 0
    Loop
    End If
    Y = 1
    X = 27
    'Sök sammlings skivorna, endast om Båda eller Sammling är valda
    If (optSamling.Value = True) Or (optBåda.Value = True) Then
    Data2.Recordset.MoveFirst
    Do Until Y > Data2.Recordset.RecordCount
        Do Until X = 80
        X = X + 1
        If UCase(txtFields(X).Text) = UCase(Finde) Then
            Found = txtFields(61).Text
            cmdGoto(Hittat).Visible = True
            cmdGoto(Hittat).Caption = Found
            cmdGoto(Hittat).Tag = Y - 1
            cmdGoto(Hittat).ToolTipText = "Collection CD"
            Found = ""
            Hittat = Hittat + 1
        End If
        Loop
        Data2.Recordset.MoveNext
        Y = Y + 1
        X = 27
    Loop
    End If
    'Om programet inte hittar några records så visas detta medelande
    If Hittat = 0 Then MsgBox "No song/record found with that title.", vbInformation, "Finder"
    Exit Sub

'Error fälla, om det blir fel visas detta medelande.
errorstopp:
Select Case Err.Number
        Case 340
            MsgBox "The program found more then 10 matchas, this program can only handel 10 matches but loock for an updat on my site.", vbCritical, "Finder"
        Case Else
            MsgBox "It accrued an error when the program searched the database. Sorry about this.", vbCritical, "Finder"
    End Select
End Sub

Private Sub cmdGoto_Click(Index As Integer)
    If cmdGoto(Index).ToolTipText = "Ordinary CD" Then
        frmsammaartist.Show
        frmsammaartist.Data1.Recordset.MoveFirst
        Do Until frmsammaartist.txtFields(1).Text = frmFinde.cmdGoto(Index).Caption
            frmsammaartist.Data1.Recordset.MoveNext
        Loop
    End If
    If cmdGoto(Index).ToolTipText = "Collection CD" Then
        frmsammling.Show
        frmsammling.Data1.Recordset.MoveFirst
        Do Until frmsammling.txtFields(1).Text = frmFinde.cmdGoto(Index).Caption
            frmsammling.Data1.Recordset.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Load()
    'Räkna antalet records i databasen, om mindre än 1 så går det inte att
    'söka i den databasen
    Set myDB = OpenDatabase("data.mdb")
    Set myRS = myDB.OpenRecordset("sammling", dbOpenTable)
    If myRS.RecordCount < 1 Then
        optBåda.Enabled = False
        optSamling.Enabled = False
        optEnsam.Value = True
    End If
    Set myDB = OpenDatabase("data.mdb")
    Set myRS = myDB.OpenRecordset("samma artist", dbOpenTable)
    If myRS.RecordCount < 1 Then
        optBåda.Enabled = False
        optEnsam.Enabled = False
        optSamling.Value = True
    End If
End Sub

Private Sub txtFind_Change()
    If txtFind <> "" Then cmdFind.Enabled = True
    If txtFind = "" Then cmdFind.Enabled = False
End Sub
