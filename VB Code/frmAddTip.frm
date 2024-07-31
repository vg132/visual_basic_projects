VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAddTip 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Add Tip to Database"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   315
      Left            =   3540
      TabIndex        =   4
      Top             =   3240
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   3240
      Width           =   1035
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
   Begin RichTextLib.RichTextBox txtTip 
      Height          =   2535
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4471
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmAddTip.frx":0000
   End
   Begin VB.Label Label2 
      Caption         =   "Tip (HTML or Text)"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tip Name:"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   45
      Width           =   735
   End
End
Attribute VB_Name = "frmAddTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    NewItem txtName.Text, txtTip.TextRTF, frmCode.GetDBNr
End Sub

Private Sub Form_Resize()
    txtName.Width = frmAddTip.Width - 990
    txtTip.Width = frmAddTip.Width - 150
    txtTip.Height = frmAddTip.Height - 1395
End Sub

