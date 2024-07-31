VERSION 5.00
Begin VB.Form frmsammaartist 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5745
   ClientLeft      =   3555
   ClientTop       =   2280
   ClientWidth     =   9570
   Icon            =   "Samma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   Tag             =   "samma artist"
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   300
      Left            =   5520
      TabIndex        =   58
      Tag             =   "&Refresh"
      Top             =   5025
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   300
      Left            =   4440
      TabIndex        =   57
      Tag             =   "&Delete"
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   300
      Left            =   3360
      TabIndex        =   56
      Tag             =   "&Add"
      Top             =   5025
      Width           =   975
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
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
      RecordSource    =   "samma artist"
      Top             =   5400
      Width           =   9570
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
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   55
      Top             =   4560
      Width           =   3375
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
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   53
      Top             =   4245
      Width           =   3375
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
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   51
      Top             =   3915
      Width           =   3375
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
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   49
      Top             =   3600
      Width           =   3375
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
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   47
      Top             =   3285
      Width           =   3375
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
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   45
      Top             =   2955
      Width           =   3375
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
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   43
      Top             =   2640
      Width           =   3375
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
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   41
      Top             =   2325
      Width           =   3375
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
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   39
      Top             =   1995
      Width           =   3375
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
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   37
      Top             =   1680
      Width           =   3375
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
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   35
      Top             =   1365
      Width           =   3375
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
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   33
      Top             =   1035
      Width           =   3375
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
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   31
      Top             =   720
      Width           =   3375
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
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   29
      Top             =   4485
      Width           =   3375
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
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   27
      Top             =   4155
      Width           =   3375
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
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   25
      Top             =   3840
      Width           =   3375
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
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   23
      Top             =   3525
      Width           =   3375
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
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   21
      Top             =   3195
      Width           =   3375
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
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   19
      Top             =   2880
      Width           =   3375
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
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   17
      Top             =   2565
      Width           =   3375
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
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   15
      Top             =   2235
      Width           =   3375
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
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   13
      Top             =   1920
      Width           =   3375
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
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   11
      Top             =   1605
      Width           =   3375
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
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   9
      Top             =   1275
      Width           =   3375
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
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   7
      Top             =   960
      Width           =   3375
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
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   5
      Top             =   645
      Width           =   3375
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
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   3
      Top             =   240
      Width           =   3375
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
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 26:"
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
      Index           =   28
      Left            =   4920
      TabIndex        =   54
      Tag             =   "Track 26:"
      Top             =   4575
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 25:"
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
      Index           =   27
      Left            =   4920
      TabIndex        =   52
      Tag             =   "Track 25:"
      Top             =   4260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 24:"
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
      Index           =   26
      Left            =   4920
      TabIndex        =   50
      Tag             =   "Track 24:"
      Top             =   3945
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 23:"
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
      Index           =   25
      Left            =   4920
      TabIndex        =   48
      Tag             =   "Track 23:"
      Top             =   3615
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 22:"
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
      Index           =   24
      Left            =   4920
      TabIndex        =   46
      Tag             =   "Track 22:"
      Top             =   3300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 21:"
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
      Index           =   23
      Left            =   4920
      TabIndex        =   44
      Tag             =   "Track 21:"
      Top             =   2985
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 20:"
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
      Index           =   22
      Left            =   4920
      TabIndex        =   42
      Tag             =   "Track 20:"
      Top             =   2655
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 19:"
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
      Index           =   21
      Left            =   4920
      TabIndex        =   40
      Tag             =   "Track 19:"
      Top             =   2340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 18:"
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
      Index           =   20
      Left            =   4920
      TabIndex        =   38
      Tag             =   "Track 18:"
      Top             =   2025
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 17:"
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
      Index           =   19
      Left            =   4920
      TabIndex        =   36
      Tag             =   "Track 17:"
      Top             =   1695
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 16:"
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
      Index           =   18
      Left            =   4920
      TabIndex        =   34
      Tag             =   "Track 16:"
      Top             =   1380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 15:"
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
      Index           =   17
      Left            =   4920
      TabIndex        =   32
      Tag             =   "Track 15:"
      Top             =   1065
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 14:"
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
      Index           =   16
      Left            =   4920
      TabIndex        =   30
      Tag             =   "Track 14:"
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 13:"
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
      Index           =   15
      Left            =   120
      TabIndex        =   28
      Tag             =   "Track 13:"
      Top             =   4500
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 12:"
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
      Index           =   14
      Left            =   120
      TabIndex        =   26
      Tag             =   "Track 12:"
      Top             =   4185
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 11:"
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
      Index           =   13
      Left            =   120
      TabIndex        =   24
      Tag             =   "Track 11:"
      Top             =   3855
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 10:"
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
      Index           =   12
      Left            =   120
      TabIndex        =   22
      Tag             =   "Track 10:"
      Top             =   3540
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 9:"
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
      Index           =   11
      Left            =   120
      TabIndex        =   20
      Tag             =   "Track 9:"
      Top             =   3225
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 8:"
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
      Index           =   10
      Left            =   120
      TabIndex        =   18
      Tag             =   "Track 8:"
      Top             =   2895
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 7:"
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
      Index           =   9
      Left            =   120
      TabIndex        =   16
      Tag             =   "Track 7:"
      Top             =   2580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 6:"
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
      Index           =   8
      Left            =   120
      TabIndex        =   14
      Tag             =   "Track 6:"
      Top             =   2265
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 5:"
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
      Index           =   7
      Left            =   120
      TabIndex        =   12
      Tag             =   "Track 5:"
      Top             =   1935
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 4:"
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
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Tag             =   "Track 4:"
      Top             =   1620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 3:"
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
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Tag             =   "Track 3:"
      Top             =   1305
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 2:"
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
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Tag             =   "Track 2:"
      Top             =   975
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Track 1:"
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
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Tag             =   "Track 1:"
      Top             =   660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Artist:"
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
      Index           =   2
      Left            =   4920
      TabIndex        =   2
      Tag             =   "Artist:"
      Top             =   240
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
      Top             =   240
      Width           =   1815
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "File"
      Begin VB.Menu mnuGoto 
         Caption         =   "Goto"
         Shortcut        =   ^G
      End
   End
End
Attribute VB_Name = "frmsammaartist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Respons As Integer
Dim Temp As String

Private Sub cmdAdd_Click()
    'Lägg till en record
    On Error GoTo erroradd
    Data1.Recordset.AddNew
    txtFields(1).SetFocus
erroradd:
    Exit Sub
End Sub


Private Sub cmdDelete_Click()
    'Delete en record
    Respons = MsgBox("Do you want to delete this record?", vbYesNo, "Finder - Delete Record")
    If Respons = vbNo Then Exit Sub
    'this may produce an error if you delete the last
    'record or the only record in the recordset
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
    'Uppdatera programet
    Data1.Refresh
End Sub

Private Sub Data1_Error(DataErr As Integer, Response As Integer)
    'Avsluta denna del av programet
    Exit Sub
End Sub

Private Sub Data1_Reposition()
    Screen.MousePointer = vbDefault
    On Error GoTo errortest
    'This will display the current record position
    'for dynasets and snapshots
    Data1.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1)
    'for the table object you must set the index property when
    'the recordset gets created and use the following line
    'Data1.Caption = "Record: " & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
errortest:
    Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Vet ej vad detta är!
    Screen.MousePointer = vbDefault
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
