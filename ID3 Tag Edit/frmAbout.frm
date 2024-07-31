VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About ID3 Tag Edit"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   315
      Left            =   3360
      TabIndex        =   0
      Top             =   1920
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Version 1.1"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   $"frmAbout.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Copyright © Viktor Gars 1999-2001"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   2400
      Picture         =   "frmAbout.frx":00E0
      Top             =   120
      Width           =   1920
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "ID3 Tag Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1665
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Unload Me
End Sub
