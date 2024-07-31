VERSION 5.00
Begin VB.Form frmShow 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mouse"
   ClientHeight    =   465
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   825
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblY 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblX 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "frmShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frmShow.Top = 0
    frmShow.Left = Screen.Width - frmShow.Width
    Size.Height = frmShow.Height
    Size.Left = frmShow.Left
    Size.Top = frmShow.Top
    Size.Width = frmShow.Width
    Call StayOnTop(frmShow.hWnd)
End Sub

Public Sub StayOnTop(ByVal hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, Size.Left / 15, Size.Top / 15, Size.Width / 15, Size.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.mnuShow.Checked = False
End Sub
