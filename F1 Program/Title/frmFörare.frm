VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmFörare 
   BackColor       =   &H005BC9FF&
   Caption         =   "Viktorex Driver Link"
   ClientHeight    =   1860
   ClientLeft      =   4320
   ClientTop       =   2865
   ClientWidth     =   4680
   Icon            =   "frmFörare.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   4680
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      DialogTitle     =   "Select a Driver"
      InitDir         =   "g:\internet\f1\olddrivers"
   End
   Begin VB.TextBox txtHTML 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
   Begin VB.CommandButton cmdLänka 
      Caption         =   "Page to Link"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmFörare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileNum As Integer
Dim Length As Integer
Dim File As String
Dim Title As String
Dim X As Integer

Private Sub cmdLänka_Click()
FileNum = FreeFile
    
    On Error GoTo fileerror
    CommonDialog1.ShowOpen
    File = CommonDialog1.filename
    Open File For Binary As FileNum
    X = 1
    Length = 7
    Do While Title <> "<title>"
        Title = String(Length, " ")
        Get #FileNum, X, Title
        X = X + 1
    Loop
    X = X + 6
    Length = 1
    txtHTML.Text = txtHTML.Text + "<a href='" + CommonDialog1.FileTitle + "'>"
    Do While Title <> "<"
        Title = String(Length, " ")
        Get #FileNum, X, Title
        X = X + 1
        If Title <> "<" Then txtHTML.Text = txtHTML.Text + Title
    Loop
    txtHTML.Text = txtHTML.Text + "</a><br>"
    Close FileNum
fileerror:
    Exit Sub
End Sub

Private Sub mnuExit_Click()
    End
End Sub
