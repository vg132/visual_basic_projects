VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Project Manager's Options"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3200
      TabIndex        =   5
      Top             =   1800
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Java Optoions"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton cmdJavaBrowse 
         Caption         =   "..."
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         Top             =   1095
         Width           =   300
      End
      Begin VB.CommandButton cmdJavacBrowse 
         Caption         =   "..."
         Height          =   255
         Left            =   3600
         TabIndex        =   1
         Top             =   495
         Width           =   300
      End
      Begin VB.TextBox txtJava 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtJavac 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "JVM location (Java.exe)"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Javac location (Javac.exe)"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdJavaBrowse_Click()
Dim RetVal As String
    RetVal = BrowseFolders("JVM location", Me.hWnd)
    If RetVal <> "" Then txtJava.Text = RetVal & "\Java.exe"
End Sub

Private Sub cmdJavacBrowse_Click()
Dim RetVal As String
    RetVal = BrowseFolders("Javac location", Me.hWnd)
    If RetVal <> "" Then txtJavac.Text = RetVal & "\Javac.exe"
End Sub

Private Sub cmdOk_Click()
    SaveSetting App.Title, "Java", "Javac", txtJavac.Text
    SaveSetting App.Title, "Java", "Java", txtJava.Text
    Java = txtJava.Text
    Javac = txtJavac.Text
    Unload Me
End Sub

Private Sub Form_Load()
    txtJavac.Text = GetSetting(App.Title, "Java", "Javac", "")
    txtJava.Text = GetSetting(App.Title, "Java", "Java", "")
End Sub
