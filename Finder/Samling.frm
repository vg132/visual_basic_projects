VERSION 5.00
Begin VB.Form frmsammling 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Collection Records"
   ClientHeight    =   6615
   ClientLeft      =   2970
   ClientTop       =   2085
   ClientWidth     =   11925
   Icon            =   "Samling.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   Tag             =   "sammling"
   Begin VB.OptionButton opt2 
      Caption         =   "Artist 14-26"
      Height          =   255
      Left            =   4680
      TabIndex        =   67
      Top             =   360
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.OptionButton opt1 
      Caption         =   "Artist 1-13"
      Height          =   255
      Left            =   4680
      TabIndex        =   66
      Top             =   0
      Value           =   -1  'True
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   315
      Left            =   6720
      TabIndex        =   65
      Tag             =   "&Refresh"
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   315
      Left            =   5640
      TabIndex        =   64
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   4560
      TabIndex        =   63
      Top             =   5880
      Width           =   975
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "sammling"
      Top             =   6270
      Width           =   11925
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
      Index           =   44
      Left            =   9000
      TabIndex        =   53
      Top             =   5520
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 26"
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
      Index           =   46
      Left            =   6000
      TabIndex        =   52
      Top             =   5520
      Width           =   2900
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
      Index           =   43
      Left            =   9000
      TabIndex        =   51
      Top             =   5160
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 25"
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
      Index           =   47
      Left            =   6000
      TabIndex        =   50
      Top             =   5160
      Width           =   2900
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
      Index           =   45
      Left            =   9000
      TabIndex        =   49
      Top             =   4800
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 24"
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
      Index           =   48
      Left            =   6000
      TabIndex        =   48
      Top             =   4800
      Width           =   2900
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
      Index           =   42
      Left            =   9000
      TabIndex        =   47
      Top             =   4440
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 23"
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
      Index           =   49
      Left            =   6000
      TabIndex        =   46
      Top             =   4440
      Width           =   2900
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
      Index           =   41
      Left            =   9000
      TabIndex        =   45
      Top             =   4080
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 22"
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
      Index           =   50
      Left            =   6000
      TabIndex        =   44
      Top             =   4080
      Width           =   2900
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
      Index           =   40
      Left            =   9000
      TabIndex        =   43
      Top             =   3720
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 21"
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
      Index           =   51
      Left            =   6000
      TabIndex        =   42
      Top             =   3720
      Width           =   2900
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
      Index           =   39
      Left            =   9000
      TabIndex        =   41
      Top             =   3360
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 20"
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
      Index           =   52
      Left            =   6000
      TabIndex        =   40
      Top             =   3360
      Width           =   2900
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
      Index           =   38
      Left            =   9000
      TabIndex        =   39
      Top             =   3000
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "artist 19"
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
      Index           =   53
      Left            =   6000
      TabIndex        =   38
      Top             =   3000
      Width           =   2900
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
      Index           =   37
      Left            =   9000
      TabIndex        =   37
      Top             =   2640
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 18"
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
      Index           =   54
      Left            =   6000
      TabIndex        =   36
      Top             =   2640
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "track 17"
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
      Index           =   0
      Left            =   9000
      TabIndex        =   35
      Top             =   2280
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 18"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   36
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   57
      Top             =   11560
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Track 17"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   35
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   55
      Top             =   11240
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 17"
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
      Index           =   34
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   34
      Top             =   2280
      Width           =   2900
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
      Index           =   33
      Left            =   9000
      MaxLength       =   50
      TabIndex        =   33
      Top             =   1920
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 16"
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
      Index           =   32
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   32
      Top             =   1920
      Width           =   2900
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
      Index           =   31
      Left            =   9000
      MaxLength       =   50
      TabIndex        =   31
      Top             =   1560
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 15"
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
      Index           =   30
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   30
      Top             =   1560
      Width           =   2900
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
      Index           =   29
      Left            =   9000
      MaxLength       =   50
      TabIndex        =   29
      Top             =   1200
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 14"
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
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   28
      Top             =   1200
      Width           =   2900
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
      Index           =   27
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   27
      Top             =   5520
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 13"
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
      Left            =   0
      MaxLength       =   50
      TabIndex        =   26
      Top             =   5520
      Width           =   2900
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
      Index           =   25
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   25
      Top             =   5160
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 12"
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
      Left            =   0
      MaxLength       =   50
      TabIndex        =   24
      Top             =   5160
      Width           =   2900
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
      Index           =   23
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   23
      Top             =   4800
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 11"
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
      Left            =   0
      MaxLength       =   50
      TabIndex        =   22
      Top             =   4800
      Width           =   2900
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
      Index           =   21
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   21
      Top             =   4440
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 10"
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
      Left            =   0
      MaxLength       =   50
      TabIndex        =   20
      Top             =   4440
      Width           =   2900
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
      Index           =   19
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   19
      Top             =   4080
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 9"
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
      Left            =   0
      MaxLength       =   50
      TabIndex        =   18
      Top             =   4080
      Width           =   2900
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
      Index           =   17
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   17
      Top             =   3720
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 8"
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
      Left            =   0
      MaxLength       =   50
      TabIndex        =   16
      Top             =   3720
      Width           =   2900
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
      Index           =   15
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   15
      Top             =   3360
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 7"
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
      Left            =   0
      MaxLength       =   50
      TabIndex        =   14
      Top             =   3360
      Width           =   2900
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
      Index           =   13
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   13
      Top             =   3000
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 6"
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
      Left            =   0
      MaxLength       =   50
      TabIndex        =   12
      Top             =   3000
      Width           =   2900
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
      Index           =   11
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   11
      Top             =   2640
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 5"
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
      Left            =   0
      MaxLength       =   50
      TabIndex        =   10
      Top             =   2640
      Width           =   2900
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
      Index           =   9
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   9
      Top             =   2280
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 4"
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
      Left            =   0
      MaxLength       =   50
      TabIndex        =   8
      Top             =   2280
      Width           =   2900
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
      Index           =   7
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1920
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 3"
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
      Left            =   0
      MaxLength       =   50
      TabIndex        =   6
      Top             =   1920
      Width           =   2900
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
      Index           =   5
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1560
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 2"
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
      Left            =   0
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1560
      Width           =   2900
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
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1200
      Width           =   2900
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist 1"
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
      Left            =   0
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1200
      Width           =   2900
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
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Track 14-26"
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
      Left            =   9000
      TabIndex        =   62
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Artist 14-26"
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
      Left            =   6000
      TabIndex        =   61
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Track 1-13"
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
      Left            =   3000
      TabIndex        =   60
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Artist 1-13"
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
      Left            =   0
      TabIndex        =   59
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 18:"
      Height          =   255
      Index           =   37
      Left            =   120
      TabIndex        =   58
      Tag             =   "Track 18:"
      Top             =   11900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Artist 18:"
      Height          =   255
      Index           =   36
      Left            =   120
      TabIndex        =   56
      Tag             =   "Artist 18:"
      Top             =   11580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 17:"
      Height          =   255
      Index           =   35
      Left            =   120
      TabIndex        =   54
      Tag             =   "Track 17:"
      Top             =   11260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CD Title:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Tag             =   "CD Title:"
      Top             =   375
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuGoto 
         Caption         =   "Goto"
         Shortcut        =   ^G
      End
   End
End
Attribute VB_Name = "frmsammling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ScreenSize As Integer
Dim X As Integer
Dim Start As Boolean
Dim Respons As Integer
Dim Temp As String

Private Sub cmdAdd_Click()
    'Lägg till ett record
    On Error GoTo erroradd
    Data1.Recordset.AddNew
    txtFields(1).SetFocus
erroradd:
Exit Sub
End Sub

Private Sub cmdDelete_Click()
    'fråga om man vill ta bort dett record
    Respons = MsgBox("Do you want to delete this record?", vbYesNo, "Finder - Delete Record")
    If Respons = vbNo Then Exit Sub
    On Error GoTo deleteerror
    With Data1.Recordset
        .Delete
        .MoveNext
        If .EOF Then .MoveLast
    End With
deleteerror:
    Exit Sub
End Sub

Private Sub cmdRefresh_Click()
    'Uppdatera databasen
    Data1.Refresh
End Sub

Private Sub Data1_Reposition()
    Screen.MousePointer = vbDefault
    On Error Resume Next
    Data1.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1)
End Sub

Private Sub Form_Load()
    'Ändra utseendet om man har 640*480 skärm, annars görs inget
    Start = False
    ScreenSize = GetSetting("Finder", "Config", "Screen Size")
    If ScreenSize = 3 Then Exit Sub
    If ScreenSize = 2 Then Exit Sub
    If ScreenSize = 1 Then
        cmdDelete.Left = cmdDelete.Left - 3120
        cmdRefresh.Left = cmdRefresh.Left - 3120
        cmdAdd.Left = cmdAdd.Left - 3120
        frmsammling.Width = 6000
        opt1.Visible = True
        opt2.Visible = True
        ScreenChange
    End If
End Sub

Public Sub ScreenChange()
    'Ändra skärmen
    If opt2.Value = True Then
        X = 2
        Do Until X = 28
            txtFields(X).Visible = False
            X = X + 1
        Loop
        Do Until X = 55
            txtFields(X).Visible = True
            If Start = False Then txtFields(X).Left = txtFields(X).Left - 6000
            X = X + 1
        Loop
        Label1.Visible = False
        Label2.Visible = False
        Label3.Visible = True
        Label4.Visible = True
        txtFields(0).Visible = True
        If Start = False Then
            Label3.Left = Label3.Left - 6000
            Label4.Left = Label4.Left - 6000
            txtFields(0).Left = txtFields(0).Left - 6000
        End If
        Start = True
    End If
    If opt1.Value = True Then
        X = 2
        Do Until X = 28
            txtFields(X).Visible = True
            X = X + 1
        Loop
        Do Until X = 55
            txtFields(X).Visible = False
            X = X + 1
        Loop
        txtFields(0).Visible = False
        Label1.Visible = True
        Label2.Visible = True
        Label3.Visible = False
        Label4.Visible = False
    End If
End Sub

Private Sub mnuGoto_Click()
    Temp = InputBox("CD Title:", "Goto")
    If Temp <> "" Then
        Data1.Recordset.MoveFirst
        Do Until UCase(txtFields(1).Text) = UCase(Temp)
            If (Data1.Recordset.EOF) And (UCase(txtFields(1).Text) <> UCase(Temp)) Then
                MsgBox "Finder could not find a record with that name.", vbInformation, "Goto"
                Exit Sub
            End If
            Data1.Recordset.MoveNext
        Loop
    End If
End Sub

Private Sub opt1_Click()
    ScreenChange
End Sub

Private Sub opt2_Click()
    ScreenChange
End Sub
