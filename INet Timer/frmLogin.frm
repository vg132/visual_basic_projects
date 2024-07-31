VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1215
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   3150
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAddUser 
      Caption         =   "&Add User"
      Height          =   315
      Left            =   105
      TabIndex        =   5
      Top             =   840
      Width           =   1035
   End
   Begin VB.ComboBox cboName 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   105
      Width           =   1935
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   1980
      TabIndex        =   2
      Top             =   840
      Width           =   1035
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Password:"
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   4
      Top             =   525
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   165
      Width           =   840
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddUser_Click()
'*************************************
'Function Name: cmdAddUser
'Use: add a new user in the last minute
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-??-??
'*************************************

    frmAddUser.Tag = "Login"
    Unload Me
    frmAddUser.Show vbModal, frmMain
End Sub

Private Sub cmdOk_Click()
'*************************************
'Function Name: cmdOk
'Use: check password
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-??-??
'*************************************
Dim vArray As Variant
    vArray = GetUser(One, cboName.Text)
    If IsArray(vArray) = True Then
        If txtPassword.Text = vArray(2, 0) & "" Then
            INetData.User = vArray(0, 0)
            frmMain.lblName = cboName.Text
            Unload Me
        Else
            MsgBox "Wrong Password or User Name. Please retry.", vbCritical
        End If
    End If
End Sub

Private Sub Form_Load()
'*************************************
'Function Name: Form_Load
'Use: Move window to always on top (as ICQ)
'the user must login befor the window will be hiden
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-??-??
'*************************************

Dim vArray As Variant
Dim X As Long
Dim Flags As Long  ' the flags specifying how to move the window
Dim RetVal As Long  ' return value
    frmLogin.Show
    Flags = SWP_NOSIZE Or SWP_DRAWFRAME
    RetVal = SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left, Me.Top, Me.Left + Me.Width, Me.Top + Me.Height, Flags) ' move the window

    vArray = GetUser(All)
    If IsArray(vArray) = True Then
        For X = 0 To UBound(vArray, 2)
            cboName.AddItem vArray(1, X)
        Next
        cboName.ListIndex = 0
    End If
End Sub
