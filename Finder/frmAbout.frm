VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   2955
   ClientLeft      =   2715
   ClientTop       =   3615
   ClientWidth     =   5025
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmAbout.frx":030A
   ScaleHeight     =   2039.593
   ScaleMode       =   0  'User
   ScaleWidth      =   4718.735
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdInternet 
      Caption         =   "Visit Viktor Gars's Internet page!"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   3255
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0614
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3600
      TabIndex        =   0
      Top             =   2520
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "http://hem.passagen.se/vg132"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2460
      TabIndex        =   6
      Top             =   1905
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   -28.172
      X2              =   5196.712
      Y1              =   1656.522
      Y2              =   1656.522
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmAbout.frx":091E
      ForeColor       =   &H00000000&
      Height          =   1050
      Left            =   1080
      TabIndex        =   2
      Top             =   1125
      Width           =   3765
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   5210.798
      Y1              =   1656.522
      Y2              =   1656.522
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.0.0"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Long
Dim URL As String
Dim Temp As Integer

Private Sub cmdInternet_Click()
    URL = "http://hem.passagen.se/vg132"
    a = ShellExecute(frmAbout.hwnd, "open", URL, vbNullString, vbNullString, 1)
End Sub

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Temp = 1
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseOut
End Sub

Private Sub Label1_Click()
    cmdInternet_Click
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Temp = 1 Then
        Label1.Left = Label1.Left + 10
        Label1.Top = Label1.Top - 10
        Temp = 2
        frmAbout.MousePointer = 99
    End If
End Sub

Private Sub lblDescription_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseOut
End Sub

Public Sub MouseOut()
    If Temp = 2 Then
        Label1.Left = Label1.Left - 10
        Label1.Top = Label1.Top + 10
        Temp = 1
        frmAbout.MousePointer = 0
    End If
End Sub
