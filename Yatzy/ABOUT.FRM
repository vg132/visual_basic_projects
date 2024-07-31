VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "Om"
   ClientHeight    =   2805
   ClientLeft      =   4905
   ClientTop       =   2595
   ClientWidth     =   2760
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   2760
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   $"About.frx":0442
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmAbout.Hide
End Sub
