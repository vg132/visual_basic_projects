VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHelp 
   Caption         =   "Help - Super YATZY"
   ClientHeight    =   6315
   ClientLeft      =   4650
   ClientTop       =   2595
   ClientWidth     =   8910
   Icon            =   "Help.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   8910
   Begin TabDlg.SSTab tabStarta 
      Height          =   5175
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9128
      _Version        =   327680
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Starta"
      TabPicture(0)   =   "Help.frx":0442
      Tab(0).ControlCount=   0
      Tab(0).ControlEnabled=   0   'False
      TabCaption(1)   =   "Regler"
      TabPicture(1)   =   "Help.frx":045E
      Tab(1).ControlCount=   0
      Tab(1).ControlEnabled=   0   'False
      TabCaption(2)   =   "Tips"
      TabPicture(2)   =   "Help.frx":047A
      Tab(2).ControlCount=   0
      Tab(2).ControlEnabled=   0   'False
      TabCaption(3)   =   "Om Spelet"
      TabPicture(3)   =   "Help.frx":0496
      Tab(3).ControlCount=   0
      Tab(3).ControlEnabled=   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Super YATZY - Hjälp"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

