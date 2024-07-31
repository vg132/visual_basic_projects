VERSION 5.00
Begin VB.Form frmScore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2145
   ClientLeft      =   3870
   ClientTop       =   3585
   ClientWidth     =   4680
   Icon            =   "Score.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4680
   Begin VB.TextBox txtScore1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtName1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtDate1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtDate3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtName3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtScore3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtDate2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtName2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtScore2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Stäng"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "High Score 1:"
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "High Score 2:"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "High Score 3"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Poäng"
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Namn"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Datum"
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    frmScore.Hide
End Sub

Private Sub Form_Load()
    frmScore.txtScore1.Text = frmYatzy.txtScore1.Text
    frmScore.txtName1.Text = frmYatzy.txtName1.Text
    frmScore.txtDate1.Text = frmYatzy.txtDate1.Text
    frmScore.txtScore2.Text = frmYatzy.txtScore2.Text
    frmScore.txtName2.Text = frmYatzy.txtName2.Text
    frmScore.txtDate2.Text = frmYatzy.txtDate2.Text
    frmScore.txtScore3.Text = frmYatzy.txtScore3.Text
    frmScore.txtName3.Text = frmYatzy.txtName3.Text
    frmScore.txtDate3.Text = frmYatzy.txtDate3.Text
End Sub
