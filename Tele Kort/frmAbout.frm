VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Om TeleKort"
   ClientHeight    =   3255
   ClientLeft      =   2700
   ClientTop       =   3345
   ClientWidth     =   4275
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2246.659
   ScaleMode       =   0  'User
   ScaleWidth      =   4014.446
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAbout.frx":030A
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "TeleKort"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    Unload frmAbout
End Sub
