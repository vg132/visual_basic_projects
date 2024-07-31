VERSION 5.00
Begin VB.Form frmAddUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Remove User"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   3615
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboName 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   840
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1035
   End
   Begin VB.CommandButton cmdAddUser 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Top             =   840
      Width           =   1035
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1095
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   495
      Width           =   2400
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Password:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   540
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Name:"
      Height          =   195
      Left            =   135
      TabIndex        =   5
      Top             =   180
      Width           =   840
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*************************************
'Function Name: frmAddUser
'Use: Add a new user to the user to the program
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-??-??
'*************************************

Private Sub cmdAddUser_Click()
Dim RetVal As Boolean
    RetVal = AddUser(cboName.Text, txtPassword)
    If RetVal = False Then MsgBox "There is already a user with this user name.", vbInformation
    Form_Load
    cboName.SetFocus
    txtPassword.Text = ""
End Sub

Private Sub cmdClose_Click()
    If Me.Tag = "Login" Then
        Unload Me
        frmLogin.Show , frmMain
    End If
    Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim vArray As Variant
Dim Responce As Integer
    vArray = GetUser(One, cboName.Text)
    If IsArray(vArray) = True Then
        If vArray(2, 0) = txtPassword.Text Then
            Responce = MsgBox("Do you want to delete this user?", vbYesNo + vbQuestion, "Delete User")
            If Responce = vbYes Then
                DeleteUser vArray(0, 0), cboName.Text
            End If
        End If
        Form_Load
    Else
        MsgBox "Wrong UserName or Password.", vbInformation
    End If
End Sub

Private Sub Form_Load()
Dim vArray As Variant
Dim X As Long
    cboName.Clear
    vArray = GetUser(All)
    If IsArray(vArray) = True Then
        For X = 0 To UBound(vArray, 2)
            cboName.AddItem vArray(1, X)
        Next
    End If
End Sub
