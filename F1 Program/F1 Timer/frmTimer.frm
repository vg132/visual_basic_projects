VERSION 5.00
Begin VB.Form frmTiming 
   BackColor       =   &H005BC9FF&
   Caption         =   "Formula One Timing"
   ClientHeight    =   2235
   ClientLeft      =   4875
   ClientTop       =   3690
   ClientWidth     =   4425
   Icon            =   "frmTimer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2235
   ScaleWidth      =   4425
   Begin VB.TextBox txt107 
      Enabled         =   0   'False
      Height          =   405
      Left            =   2640
      MaxLength       =   6
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton optRace 
      BackColor       =   &H005BC9FF&
      Caption         =   "Race"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.OptionButton optQual 
      BackColor       =   &H005BC9FF&
      Caption         =   "Qual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox txtSvar 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2640
      MaxLength       =   5
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtNästa 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtBästa 
      Height          =   285
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H005BC9FF&
      Caption         =   "107% of the best time:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H005BC9FF&
      Caption         =   "Skilnad i Sekunder:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H005BC9FF&
      Caption         =   "Nästa tid:"
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
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H005BC9FF&
      Caption         =   "Bästa tid i sekunder:"
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
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmTiming"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Count2 As Single
Dim Regel As Single


Private Sub optQual_Click()
    If optQual = True Then
        txtSvar.MaxLength = 5
        frmTiming.Height = 2640
    End If
    If optRace = True Then
        txtSvar.MaxLength = 6
        frmTiming.Height = 2025
    End If
End Sub

Private Sub optRace_Click()
    If optQual = True Then
        txtSvar.MaxLength = 5
        frmTiming.Height = 2640
    End If
    If optRace = True Then
        txtSvar.MaxLength = 6
        frmTiming.Height = 2025
    End If
End Sub

Private Sub txtBästa_Change()
    Räkna
    Regel = txtBästa.Text
    Regel = Regel * 1.07
    If Regel < 100 Then txt107.MaxLength = 6
    If Regel >= 100 Then txt107.MaxLength = 7
    txt107.Text = Regel
End Sub

Private Sub txtNästa_Change()
    Räkna
End Sub

Public Sub Räkna()
    If (txtNästa.Text <> "") And (txtBästa <> "") Then
        Count2 = txtNästa.Text - txtBästa.Text
        txtSvar.Text = Count2
    End If
End Sub
