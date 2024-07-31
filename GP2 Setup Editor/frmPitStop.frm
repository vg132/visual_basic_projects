VERSION 5.00
Begin VB.Form frmPitStop 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pit Stop Editor"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2985
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   325
      Left            =   1080
      TabIndex        =   13
      Top             =   2040
      Width           =   1000
   End
   Begin VB.OptionButton opt0 
      Caption         =   "No Stop"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton opt1 
      Caption         =   "1 Stop"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.OptionButton opt2 
      Caption         =   "2 Stops"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.OptionButton opt3 
      Caption         =   "3 Stops"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
   Begin VB.HScrollBar scr3 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   10
      Left            =   1320
      Max             =   100
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.HScrollBar scr1 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   10
      Left            =   1320
      Max             =   100
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.HScrollBar scr2 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   10
      Left            =   1320
      Max             =   100
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "100%"
      Height          =   195
      Left            =   2520
      TabIndex        =   12
      Top             =   120
      Width           =   390
      Visible         =   0   'False
   End
   Begin VB.Label lbl2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "100%"
      Height          =   195
      Left            =   2520
      TabIndex        =   11
      Top             =   720
      Width           =   390
      Visible         =   0   'False
   End
   Begin VB.Label lbl3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "100%"
      Height          =   195
      Left            =   2520
      TabIndex        =   10
      Top             =   1320
      Width           =   390
      Visible         =   0   'False
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "3rd Stop"
      Height          =   195
      Left            =   1320
      TabIndex        =   5
      Top             =   1320
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "2nd Stop"
      Height          =   195
      Left            =   1320
      TabIndex        =   4
      Top             =   720
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "1st Stop"
      Height          =   195
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   585
   End
End
Attribute VB_Name = "frmPitStop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
