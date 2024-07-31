VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Super Formula One"
   ClientHeight    =   8940
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9045
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTid 
      Caption         =   "Ren Tid"
      Height          =   375
      Left            =   5040
      TabIndex        =   95
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton cmd2HTML 
      Caption         =   "Write HTML"
      Height          =   375
      Left            =   480
      TabIndex        =   94
      Top             =   8520
      Width           =   1215
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   0
      Left            =   6840
      TabIndex        =   93
      Top             =   8520
      Width           =   255
      Visible         =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ren Diff"
      Height          =   375
      Left            =   7320
      TabIndex        =   92
      Top             =   8520
      Width           =   1215
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   22
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   90
      Top             =   8160
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   21
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   89
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   20
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   88
      Top             =   7440
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   19
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   87
      Top             =   7080
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   18
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   86
      Top             =   6720
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   17
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   85
      Top             =   6360
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   16
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   84
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   15
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   83
      Top             =   5640
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   14
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   82
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   13
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   81
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   12
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   80
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   11
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   79
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   10
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   78
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   9
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   77
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   8
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   76
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   7
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   75
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   6
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   74
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   5
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   73
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   4
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   72
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   1
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   71
      Text            =   "+0.000"
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   2
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   70
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtDiff 
      Height          =   285
      Index           =   3
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   69
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtTime2 
      Height          =   285
      Left            =   4680
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtTime22 
      Height          =   285
      Left            =   4680
      TabIndex        =   43
      Top             =   8160
      Width           =   1935
   End
   Begin VB.TextBox txtTime21 
      Height          =   285
      Left            =   4680
      TabIndex        =   41
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox txtTime20 
      Height          =   285
      Left            =   4680
      TabIndex        =   39
      Top             =   7440
      Width           =   1935
   End
   Begin VB.TextBox txtTime19 
      Height          =   285
      Left            =   4680
      TabIndex        =   37
      Top             =   7080
      Width           =   1935
   End
   Begin VB.TextBox txtTime18 
      Height          =   285
      Left            =   4680
      TabIndex        =   35
      Top             =   6720
      Width           =   1935
   End
   Begin VB.TextBox txtTime17 
      Height          =   285
      Left            =   4680
      TabIndex        =   33
      Top             =   6360
      Width           =   1935
   End
   Begin VB.TextBox txtTime16 
      Height          =   285
      Left            =   4680
      TabIndex        =   31
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox txtTime15 
      Height          =   285
      Left            =   4680
      TabIndex        =   29
      Top             =   5640
      Width           =   1935
   End
   Begin VB.TextBox txtTime14 
      Height          =   285
      Left            =   4680
      TabIndex        =   27
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox txtTime13 
      Height          =   285
      Left            =   4680
      TabIndex        =   25
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox txtTime12 
      Height          =   285
      Left            =   4680
      TabIndex        =   23
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox txtTime11 
      Height          =   285
      Left            =   4680
      TabIndex        =   21
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox txtTime10 
      Height          =   285
      Left            =   4680
      TabIndex        =   19
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txtTime9 
      Height          =   285
      Left            =   4680
      TabIndex        =   17
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox txtTime8 
      Height          =   285
      Left            =   4680
      TabIndex        =   15
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtTime7 
      Height          =   285
      Left            =   4680
      TabIndex        =   13
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtTime6 
      Height          =   285
      Left            =   4680
      TabIndex        =   11
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtTime5 
      Height          =   285
      Left            =   4680
      TabIndex        =   9
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtTime4 
      Height          =   285
      Left            =   4680
      TabIndex        =   7
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtTime3 
      Height          =   285
      Left            =   4680
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtTime1 
      Height          =   285
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtDriver22 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   66
      Top             =   8160
      Width           =   1935
   End
   Begin VB.TextBox txtDriver21 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   65
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox txtDriver20 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   7440
      Width           =   1935
   End
   Begin VB.TextBox txtDriver19 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   7080
      Width           =   1935
   End
   Begin VB.TextBox txtDriver18 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   6720
      Width           =   1935
   End
   Begin VB.TextBox txtDriver17 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   6360
      Width           =   1935
   End
   Begin VB.TextBox txtDriver16 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox txtDriver15 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   5640
      Width           =   1935
   End
   Begin VB.TextBox txtDriver14 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox txtDriver13 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox txtDriver12 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox txtDriver11 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox txtDriver10 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txtDriver9 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox txtDriver8 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtDriver7 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtDriver6 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtDriver5 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtDriver4 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtDriver3 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtDriver2 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtDriver1 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtDriver22 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   42
      Top             =   8160
      Width           =   1935
   End
   Begin VB.TextBox txtDriver21 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   40
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox txtDriver20 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   38
      Top             =   7440
      Width           =   1935
   End
   Begin VB.TextBox txtDriver19 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   36
      Top             =   7080
      Width           =   1935
   End
   Begin VB.TextBox txtDriver18 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   34
      Top             =   6720
      Width           =   1935
   End
   Begin VB.TextBox txtDriver17 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   32
      Top             =   6360
      Width           =   1935
   End
   Begin VB.TextBox txtDriver16 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   30
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox txtDriver15 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   28
      Top             =   5640
      Width           =   1935
   End
   Begin VB.TextBox txtDriver14 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox txtDriver13 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox txtDriver12 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox txtDriver11 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox txtDriver10 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txtDriver9 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox txtDriver8 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtDriver7 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtDriver6 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtDriver5 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtDriver4 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtDriver3 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtDriver2 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtDriver1 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Diff"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   91
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   68
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Team"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   67
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Driver"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   44
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu mnuArkiv 
      Caption         =   "Arkiv"
      Begin VB.Menu mnuDiff 
         Caption         =   "Ren Diff"
      End
      Begin VB.Menu mnuRenTid 
         Caption         =   "Ren Tid"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHTML 
         Caption         =   "Skriv HTML"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Avsluta"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DName As String 'Driver name
Dim Team2 As String 'Team Name
Dim BTime As Long 'Best time
Dim DTime As Long 'Time to conper
Dim Diff As Long 'Diff
Dim X As Integer
Dim DiffFul As String 'Som Diff ser ut när den är oformaterad
Dim DiffRen As String 'Diff när den är formaterad
Dim CharLoop As Integer
Dim FulLength As Integer
Dim FileNum As Integer
Dim Tryck As Boolean

Public Sub Team()
    If DName = "schumacher" Then
        Team2 = "Ferrari"
        DName = "Michael Schumacher"
    End If
    If DName = "irvine" Then
        Team2 = "Ferrari"
        DName = "Eddie Irvine"
    End If
    If DName = "hakkinen" Then
        Team2 = "McLaren-Mercedes"
        DName = "Mika Hakkinen"
    End If
    If DName = "coulthard" Then
        Team2 = "McLaren-Mercedes"
        DName = "David Coulthard"
    End If
    If DName = "villeneuve" Then
        Team2 = "Williams-Mecachrome"
        DName = "Jacques Villeneuve"
    End If
    If DName = "frentzen" Then
        Team2 = "Williams-Mecachrome"
        DName = "Heinz-Harald Frentzen"
    End If
    If DName = "wurz" Then
        Team2 = "Benetton-Mecachrome"
        DName = "Alexander Wurz"
    End If
    If DName = "fisichella" Then
        Team2 = "Benetton-Mecachrome"
        DName = "Giancarlo Fisichella"
    End If
    If DName = "ralf" Then
        Team2 = "Jordan-Mugen/Honda"
        DName = "Ralf Schumacher"
    End If
    If DName = "hill" Then
        Team2 = "Jordan-Mugen/Honda"
        DName = "Damon Hill"
    End If
    If DName = "trulli" Then
        Team2 = "Prost-Peugeot"
        DName = "Jarno Trulli"
    End If
    If DName = "panis" Then
        Team2 = "Prost-Peugeot"
        DName = "Olivier Panis"
    End If
    If DName = "nakano" Then
        Team2 = "Minardi-Ford"
        DName = "Shinji Nakano"
    End If
    If DName = "tuero" Then
        Team2 = "Minardi-Ford"
        DName = "Esteban Tuero"
    End If
    If DName = "rosset" Then
        Team2 = "Tyrrell-Ford"
        DName = "Ricardo Rosset"
    End If
    If DName = "takagi" Then
        Team2 = "Tyrrell-Ford"
        DName = "Toranosuke Takagi"
    End If
    If DName = "salo" Then
        Team2 = "Arrows"
        DName = "Mika Salo"
    End If
    If DName = "diniz" Then
        Team2 = "Arrows"
        DName = "Pedro Diniz"
    End If
    If DName = "barrichello" Then
        Team2 = "Stewart-Ford"
        DName = "Rubens Barrichello"
    End If
    If DName = "magnussen" Then
        Team2 = "Stewart-Ford"
        DName = "Jan Magnussen"
    End If
    If DName = "herbert" Then
        Team2 = "Sauber-Petronas"
        DName = "Johnny Herbert"
    End If
    If DName = "alesi" Then
        Team2 = "Sauber-Petronas"
        DName = "Jean Alesi"
    End If
End Sub

Private Sub cmd2HTML_Click()
    WriteHTML
End Sub

Private Sub cmdTid_Click()
    Tryck = True
    RenTid
End Sub

Private Sub Command1_Click()
    RenDiff
End Sub

Private Sub Form_Load()
    Tryck = False
End Sub

Private Sub mnuDiff_Click()
    RenDiff
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuHTML_Click()
    WriteHTML
    Beep
    MsgBox "HTML texten är färdig, filen ligger i katalogen g:\internet\f1\ och filen heter superf1.htm", vbInformation, "HTML"
End Sub

Private Sub mnuRenTid_Click()
    Tryck = True
    RenTid
End Sub

Private Sub txtDriver1_Change(Index As Integer)
    DName = txtDriver1(0).Text
    Team
    txtDriver1(1).Text = Team2
    txtDriver1(0).Text = DName
End Sub

Private Sub txtDriver2_Change(Index As Integer)
    DName = txtDriver2(0).Text
    Team
    txtDriver2(1).Text = Team2
    txtDriver2(0).Text = DName
End Sub
Private Sub txtDriver3_Change(Index As Integer)
    DName = txtDriver3(0).Text
    Team
    txtDriver3(1).Text = Team2
    txtDriver3(0).Text = DName
End Sub
Private Sub txtDriver4_Change(Index As Integer)
    DName = txtDriver4(0).Text
    Team
    txtDriver4(1).Text = Team2
    txtDriver4(0).Text = DName
End Sub
Private Sub txtDriver5_Change(Index As Integer)
    DName = txtDriver5(0).Text
    Team
    txtDriver5(1).Text = Team2
    txtDriver5(0).Text = DName
End Sub
Private Sub txtDriver6_Change(Index As Integer)
    DName = txtDriver6(0).Text
    Team
    txtDriver6(1).Text = Team2
    txtDriver6(0).Text = DName
End Sub
Private Sub txtDriver7_Change(Index As Integer)
    DName = txtDriver7(0).Text
    Team
    txtDriver7(1).Text = Team2
    txtDriver7(0).Text = DName
End Sub
Private Sub txtDriver8_Change(Index As Integer)
    DName = txtDriver8(0).Text
    Team
    txtDriver8(1).Text = Team2
    txtDriver8(0).Text = DName
End Sub
Private Sub txtDriver9_Change(Index As Integer)
    DName = txtDriver9(0).Text
    Team
    txtDriver9(1).Text = Team2
    txtDriver9(0).Text = DName
End Sub
Private Sub txtDriver10_Change(Index As Integer)
    DName = txtDriver10(0).Text
    Team
    txtDriver10(1).Text = Team2
    txtDriver10(0).Text = DName
End Sub
Private Sub txtDriver11_Change(Index As Integer)
    DName = txtDriver11(0).Text
    Team
    txtDriver11(1).Text = Team2
    txtDriver11(0).Text = DName
End Sub
Private Sub txtDriver12_Change(Index As Integer)
    DName = txtDriver12(0).Text
    Team
    txtDriver12(1).Text = Team2
    txtDriver12(0).Text = DName
End Sub
Private Sub txtDriver13_Change(Index As Integer)
    DName = txtDriver13(0).Text
    Team
    txtDriver13(1).Text = Team2
    txtDriver13(0).Text = DName
End Sub
Private Sub txtDriver14_Change(Index As Integer)
    DName = txtDriver14(0).Text
    Team
    txtDriver14(1).Text = Team2
    txtDriver14(0).Text = DName
End Sub
Private Sub txtDriver15_Change(Index As Integer)
    DName = txtDriver15(0).Text
    Team
    txtDriver15(1).Text = Team2
    txtDriver15(0).Text = DName
End Sub
Private Sub txtDriver16_Change(Index As Integer)
    DName = txtDriver16(0).Text
    Team
    txtDriver16(1).Text = Team2
    txtDriver16(0).Text = DName
End Sub
Private Sub txtDriver17_Change(Index As Integer)
    DName = txtDriver17(0).Text
    Team
    txtDriver17(1).Text = Team2
    txtDriver17(0).Text = DName
End Sub
Private Sub txtDriver18_Change(Index As Integer)
    DName = txtDriver18(0).Text
    Team
    txtDriver18(1).Text = Team2
    txtDriver18(0).Text = DName
End Sub
Private Sub txtDriver19_Change(Index As Integer)
    DName = txtDriver19(0).Text
    Team
    txtDriver19(1).Text = Team2
    txtDriver19(0).Text = DName
End Sub
Private Sub txtDriver20_Change(Index As Integer)
    DName = txtDriver20(0).Text
    Team
    txtDriver20(1).Text = Team2
    txtDriver20(0).Text = DName
End Sub

Private Sub txtDriver21_Change(Index As Integer)
    DName = txtDriver21(0).Text
    Team
    txtDriver21(1).Text = Team2
    txtDriver21(0).Text = DName
End Sub
Private Sub txtDriver22_Change(Index As Integer)
    DName = txtDriver22(0).Text
    Team
    txtDriver22(1).Text = Team2
    txtDriver22(0).Text = DName
End Sub

Public Sub DiffTime()
    BTime = txtTime1.Text
    Diff = DTime - BTime
End Sub

Private Sub txtTime2_Change()
    If txtTime1 <> "" And txtTime2.Text <> "" And Tryck = False Then
        DTime = txtTime2.Text
        DiffTime
        txtDiff(2).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime3_Change()
    If txtTime1 <> "" And txtTime3.Text <> "" And Tryck = False Then
        DTime = txtTime3.Text
        DiffTime
        txtDiff(3).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime4_Change()
    If txtTime1 <> "" And txtTime4.Text <> "" And Tryck = False Then
        DTime = txtTime4.Text
        DiffTime
        txtDiff(4).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime5_Change()
    If txtTime1 <> "" And txtTime5.Text <> "" And Tryck = False Then
        DTime = txtTime5.Text
        DiffTime
        txtDiff(5).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime6_Change()
    If txtTime1 <> "" And txtTime6.Text <> "" And Tryck = False Then
        DTime = txtTime6.Text
        DiffTime
        txtDiff(6).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime7_Change()
    If txtTime1 <> "" And txtTime7.Text <> "" And Tryck = False Then
        DTime = txtTime7.Text
        DiffTime
        txtDiff(7).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime8_Change()
    If txtTime1 <> "" And txtTime8.Text <> "" And Tryck = False Then
        DTime = txtTime8.Text
        DiffTime
        txtDiff(8).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime9_Change()
    If txtTime1 <> "" And txtTime9.Text <> "" And Tryck = False Then
        DTime = txtTime9.Text
        DiffTime
        txtDiff(9).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime10_Change()
    If txtTime1 <> "" And txtTime10.Text <> "" And Tryck = False Then
        DTime = txtTime10.Text
        DiffTime
        txtDiff(10).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime11_Change()
    If txtTime1 <> "" And txtTime11.Text <> "" And Tryck = False Then
        DTime = txtTime11.Text
        DiffTime
        txtDiff(11).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime12_Change()
    If txtTime1 <> "" And txtTime12.Text <> "" And Tryck = False Then
        DTime = txtTime12.Text
        DiffTime
        txtDiff(12).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime13_Change()
    If txtTime1 <> "" And txtTime13.Text <> "" And Tryck = False Then
        DTime = txtTime13.Text
        DiffTime
        txtDiff(13).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime14_Change()
    If txtTime1 <> "" And txtTime14.Text <> "" And Tryck = False Then
        DTime = txtTime14.Text
        DiffTime
        txtDiff(14).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime15_Change()
    If txtTime1 <> "" And txtTime15.Text <> "" And Tryck = False Then
        DTime = txtTime15.Text
        DiffTime
        txtDiff(15).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime16_Change()
    If txtTime1 <> "" And txtTime16.Text <> "" And Tryck = False Then
        DTime = txtTime16.Text
        DiffTime
        txtDiff(16).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime17_Change()
    If txtTime1 <> "" And txtTime17.Text <> "" And Tryck = False Then
        DTime = txtTime17.Text
        DiffTime
        txtDiff(17).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime18_Change()
    If txtTime1 <> "" And txtTime18.Text <> "" And Tryck = False Then
        DTime = txtTime18.Text
        DiffTime
        txtDiff(18).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime19_Change()
    If txtTime1 <> "" And txtTime19.Text <> "" And Tryck = False Then
        DTime = txtTime19.Text
        DiffTime
        txtDiff(19).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime20_Change()
    If txtTime1 <> "" And txtTime20.Text <> "" And Tryck = False Then
        DTime = txtTime20.Text
        DiffTime
        txtDiff(20).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime21_Change()
    If txtTime1 <> "" And txtTime21.Text <> "" And Tryck = False Then
        DTime = txtTime21.Text
        DiffTime
        txtDiff(21).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub
Private Sub txtTime22_Change()
    If txtTime1 <> "" And txtTime22.Text <> "" And Tryck = False Then
        DTime = txtTime22.Text
        DiffTime
        txtDiff(22).Text = Diff
    Else
        If txtTime1.Text <> "" Then MsgBox "You need to write a time for the first driver first"
    End If
End Sub

Public Sub RenDiff()
    X = 2
    Do Until X = 23
    DiffFul = txtDiff(X).Text
    FulLength = Len(DiffFul)
    If FulLength > 0 Then
        If FulLength = 1 Then
            DiffRen = "+0.00"
            DiffRen = DiffRen + Mid(DiffFul, 1, 1)
            txtDiff(X).Text = DiffRen
        End If
        If FulLength = 2 Then
            DiffRen = "+0.0"
            DiffRen = DiffRen + Mid(DiffFul, 1, 2)
            txtDiff(X).Text = DiffRen
        End If
        If FulLength = 3 Then
            DiffRen = "+0."
            DiffRen = DiffRen + Mid(DiffFul, 1, 3)
            txtDiff(X).Text = DiffRen
        End If
        If FulLength = 4 Then
            DiffRen = "+"
            DiffRen = DiffRen + Mid(DiffFul, 1, 1)
            DiffRen = DiffRen + "."
            DiffRen = DiffRen + Mid(DiffFul, 2, 3)
            txtDiff(X).Text = DiffRen
        End If
    End If
    X = X + 1
    Loop
End Sub

Public Sub WriteHTML()
    FileNum = FreeFile
    Open "g:\internet\f1\superf1.htm" For Output As FileNum
    
    Print #FileNum, "<table border='2'><tr><td><table border='1' cellpadding='0' cellspacing='0'><tr><td width='30'><font size='2' face='Arial'><strong>Pos</strong></font></td><td width='200'><font size='2' face='Arial'><strong>Driver </strong></font></td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'><strong>Team</strong></font></td><td width='60'><font size='2' face='Arial'><strong>Laptime</strong></font></td><td width='60'><font size='2' face='Arial'><strong>Gap</strong></font></td></tr><tr><td width='30'><font size='2' face='Arial'>1</font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver1(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver1(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime1.Text + "</td>"
    Print #FileNum, " <td width='60'>" + txtDiff(1).Text + "</td>"
    Print #FileNum, " </tr><tr><td width='30'><font size='2' face='Arial'>2</font></td>"
    Print #FileNum, " <td width='200'>" + txtDriver2(0).Text + "</td>"
    Print #FileNum, " <td width='150'>" + txtDriver2(1).Text + "</td>"
    Print #FileNum, " <td width='60'>" + txtTime2.Text + "</td>"
    Print #FileNum, " <td width='60'>" + txtDiff(2).Text + "</td>"
    Print #FileNum, " </tr><tr><td width='30'><font size='2' face='Arial'>3</font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver3(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver3(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime3.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(3).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>4</font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver4(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver4(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime4.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(4).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>5</font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver5(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver5(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime5.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(5).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>6</font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver6(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver6(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime6.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(6).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>7</font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver7(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver7(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime7.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(7).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>8</font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver8(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver8(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime8.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(8).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>9</font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver9(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver9(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime9.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(9).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>10</font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver10(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver10(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime10.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(10).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>11</font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver11(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver11(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime11.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(11).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>12</font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver12(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver12(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime12.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(12).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>13</font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver13(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver13(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime13.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(13).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>14</font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver14(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver14(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime14.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(14).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>15 </font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver15(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver15(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime15.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(15).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>16 </font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver16(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver16(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime16.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(16).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>17 </font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver17(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver17(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime17.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(17).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>18 </font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver18(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver18(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime18.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(18).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>19 </font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver19(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver19(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime19.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(19).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>20 </font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver20(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver20(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime20.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(20).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>21 </font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver21(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver21(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime21.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(21).Text + "</td>"
    Print #FileNum, "</tr><tr><td width='30'><font size='2' face='Arial'>22 </font></td>"
    Print #FileNum, "<td width='200'><font size='2' face='Arial'>" + txtDriver22(0).Text + "</td>"
    Print #FileNum, "<td width='150'><font size='2' face='Arial'>" + txtDriver22(1).Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtTime22.Text + "</td>"
    Print #FileNum, "<td width='60'><font size='2' face='Arial'>" + txtDiff(22).Text + "</td>"
    Print #FileNum, "</tr></table></td></tr></table>"
    
    Beep
    MsgBox "HTML texten är färdig, filen ligger i katalogen g:\internet\f1\ och filen heter superf1.htm", vbInformation, "HTML"
    
End Sub

Public Sub RenTid()
    FulLength = Len(txtTime1.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime1.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime1.Text = DiffRen
    End If
    FulLength = Len(txtTime2.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime2.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime2.Text = DiffRen
    End If
    FulLength = Len(txtTime3.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime3.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime3.Text = DiffRen
    End If
    FulLength = Len(txtTime4.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime4.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime4.Text = DiffRen
    End If
    FulLength = Len(txtTime5.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime5.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime5.Text = DiffRen
    End If
    FulLength = Len(txtTime6.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime7.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime6.Text = DiffRen
    End If
    FulLength = Len(txtTime7.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime7.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime7.Text = DiffRen
    End If
    FulLength = Len(txtTime8.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime8.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime8.Text = DiffRen
    End If
    FulLength = Len(txtTime9.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime9.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime9.Text = DiffRen
    End If
    FulLength = Len(txtTime10.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime10.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime10.Text = DiffRen
    End If
    FulLength = Len(txtTime11.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime11.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime11.Text = DiffRen
    End If
    FulLength = Len(txtTime12.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime12.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime12.Text = DiffRen
    End If
    FulLength = Len(txtTime13.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime13.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime13.Text = DiffRen
    End If
    FulLength = Len(txtTime14.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime14.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime14.Text = DiffRen
    End If
    FulLength = Len(txtTime15.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime15.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime15.Text = DiffRen
    End If
    FulLength = Len(txtTime16.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime16.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime16.Text = DiffRen
    End If
    FulLength = Len(txtTime17.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime17.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime17.Text = DiffRen
    End If
    FulLength = Len(txtTime18.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime18.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime18.Text = DiffRen
    End If
    FulLength = Len(txtTime19.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime19.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime19.Text = DiffRen
    End If
    FulLength = Len(txtTime20.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime20.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime20.Text = DiffRen
    End If
    FulLength = Len(txtTime21.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime21.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime21.Text = DiffRen
    End If
    FulLength = Len(txtTime22.Text)
    If FulLength = 6 Then
    DiffRen = ""
    DiffFul = txtTime22.Text
    DiffRen = DiffRen + Mid(DiffFul, 1, 1)
    DiffRen = DiffRen + ":"
    DiffRen = DiffRen + Mid(DiffFul, 2, 2)
    DiffRen = DiffRen + "."
    DiffRen = DiffRen + Mid(DiffFul, 4, 3)
    txtTime22.Text = DiffRen
    End If
End Sub
