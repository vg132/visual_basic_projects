VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User Password"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3690
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   360
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   2580
      TabIndex        =   2
      Top             =   360
      Width           =   1035
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label1 
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1950
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Password As String

Private Sub cmdCancel_Click()
    CheckUser = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If txtPassword.Text = Password Then
        CheckUser = True
    Else
        CheckUser = False
        MsgBox "Wrong Password", vbCritical
    End If
    Unload Me
End Sub

Private Sub Form_Load()
Dim vArray As Variant
    Label1.Caption = "Password for " & CheckUserName & ":"
    vArray = GetUser(One, frmLogShow.TreeView1.SelectedItem.Text)
    If IsArray(vArray) Then
        Password = vArray(2, 0)
    End If
End Sub
