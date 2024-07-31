VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Warper"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   1515
      TabIndex        =   0
      Top             =   1680
      Width           =   1035
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4200
      Top             =   3120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Mouse Warper"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4065
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Version 1.0 (C) 2001 Viktor Gars"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   4065
   End
   Begin VB.Label lblINet 
      Alignment       =   2  'Center
      Caption         =   "http://www.vgsoftware.com/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1005
      MouseIcon       =   "Form1.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "viktor.gars@telia.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1245
      MouseIcon       =   "Form1.frx":0594
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Menu mnuSystemTray 
      Caption         =   "mnuSystemTray"
      Visible         =   0   'False
      Begin VB.Menu mnuChange 
         Caption         =   "Disable Mouse Warper"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Show Mouse X, Y"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Mouse Warper"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetCursorPos& Lib "user32" (ByVal X As Long, ByVal Y As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim Go As Boolean

Public Sub SetMousePos(ByVal X As Long, ByVal Y As Long)
    SetCursorPos X, Y
End Sub

Private Sub cmdOk_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    GoSystemTray
    Go = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VBGTray.cbSize = Len(VBGTray)
    VBGTray.hWnd = Me.hWnd
    VBGTray.uId = vbNull
    Call Shell_NotifyIcon(NIM_DELETE, VBGTray)
    End
End Sub

Private Sub Label3_Click()
Dim RetVal As Long
    RetVal = ShellExecute(Me.hWnd, "open", "mailto:viktor.gars@telia.com", vbNullString, vbNullString, 1)
End Sub

Private Sub lblINet_Click()
Dim RetVal As Long
    RetVal = ShellExecute(Me.hWnd, "open", "http://www.vgsoftware.com/", vbNullString, vbNullString, 1)
End Sub

Private Sub mnuAbout_Click()
    Me.Show
End Sub

Private Sub mnuChange_Click()
    If mnuChange.Caption = "Enable Mouse Warper" Then
        mnuChange.Caption = "Disable Mouse Warper"
        Go = True
    Else
        mnuChange.Caption = "Enable Mouse Warper"
        Go = False
    End If
End Sub

Private Sub mnuExit_Click()
    Form_Unload 0
End Sub

Private Sub mnuShow_Click()
    If mnuShow.Checked = False Then
        frmShow.Show
        mnuShow.Checked = True
    Else
        mnuShow.Checked = False
        Unload frmShow
    End If
End Sub

Private Sub Timer1_Timer()
    GetCursorPos Pnt
    If Go = True Then
        If Pnt.X = 0 Then
            SetMousePos Screen.Width / 15, Pnt.Y
        End If
        If Pnt.X = (Screen.Width / 15) - 1 Then
            SetMousePos 1, Pnt.Y
        End If
        If Pnt.Y = 0 Then
            SetMousePos Pnt.X, Screen.Height / 15
        End If
        If Pnt.Y = (Screen.Height / 15) - 1 Then
            SetMousePos Pnt.X, 1
        End If
    End If
    If Form1.mnuShow.Checked = True Then
        frmShow.lblX.Caption = "X=" & Pnt.X
        frmShow.lblY.Caption = "Y=" & Pnt.Y
    End If
End Sub

Private Sub GoSystemTray()

    VBGTray.cbSize = Len(VBGTray)
    VBGTray.hWnd = Me.hWnd
    VBGTray.uId = vbNull
    VBGTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    VBGTray.ucallbackMessage = WM_MOUSEMOVE
    
    VBGTray.hIcon = Me.Icon
    'tooltiptext
    VBGTray.szTip = Me.Caption & vbNullChar
    Call Shell_NotifyIcon(NIM_ADD, VBGTray)
    App.TaskVisible = False 'remove application from taskbar
    Me.Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static lngMsg As Long
Static blnFlag As Boolean
Dim result As Long

    lngMsg = X / Screen.TwipsPerPixelX
    If blnFlag = False Then
        blnFlag = True
        Select Case lngMsg
        Case WM_LBUTTONDBLCLICK
        Me.Show
        'right-click
        Case WM_RBUTTONUP
        result = SetForegroundWindow(Me.hWnd)
        Me.PopupMenu mnuSystemTray
        End Select
        blnFlag = False
    End If
End Sub
