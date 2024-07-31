VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INet Counter"
   ClientHeight    =   2805
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3435
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   3435
   StartUpPosition =   1  'CenterOwner
   WindowState     =   1  'Minimized
   Begin VB.CheckBox chkHoliday 
      Caption         =   "Holiday"
      Height          =   195
      Left            =   2160
      TabIndex        =   9
      Top             =   1640
      Width           =   1215
   End
   Begin VB.Frame fraMenuSplit 
      Height          =   30
      Left            =   0
      TabIndex        =   8
      Top             =   -5
      Width           =   3015
   End
   Begin VB.PictureBox picShow 
      Height          =   255
      Left            =   1920
      ScaleHeight     =   195
      ScaleWidth      =   1275
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
      Visible         =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "This sesion"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.Label lblConName 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   11
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Conection Name:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label lblPrice 
         Height          =   195
         Left            =   1560
         TabIndex        =   6
         Top             =   1080
         Width           =   1605
      End
      Begin VB.Label lblName 
         Height          =   195
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   1605
      End
      Begin VB.Label lblTime 
         Height          =   195
         Left            =   1560
         TabIndex        =   4
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Online Time:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User Name:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   840
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2880
      Top             =   1920
   End
   Begin VB.Image imgOffline 
      Height          =   480
      Left            =   1920
      Picture         =   "frmMain.frx":030A
      Top             =   1920
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Image imgOnline 
      Height          =   480
      Left            =   2400
      Picture         =   "frmMain.frx":0614
      Top             =   1920
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuConect2 
         Caption         =   "Conect"
         Visible         =   0   'False
         Begin VB.Menu mnuConect 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuChange 
         Caption         =   "&Change User"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuQuit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuAuto 
         Caption         =   "Auto Login"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewLog 
         Caption         =   "View INet Log..."
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuUserEditor 
         Caption         =   "User Editor..."
      End
      Begin VB.Menu mnuSetTariff 
         Caption         =   "Tariff Editor..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Dim Start As Boolean

Dim sEnd As String
Dim Icon_Data As NOTIFYICONDATA

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Dim Online As Boolean
Dim bExit As Boolean

Private Sub chkHoliday_Click()
    sEnd = ""
End Sub

Private Sub Form_Load()
Dim RetVal As String
    bExit = False
    On Error Resume Next
    mnuAuto.Checked = GetSetting(App.Title, "Settings", "Auto Login")
    'load databasen
    OpenDataBase
    'put icon in systray
    Icon_Data.cbSize = Len(Icon_Data)
    Icon_Data.hwnd = picShow.hwnd
    Icon_Data.uId = 1&
    Icon_Data.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    Icon_Data.ucallbackMessage = WM_MOUSEMOVE
    Icon_Data.hIcon = imgOffline
    Icon_Data.szTip = "INet Counter - Offline" & Chr(0)
    Shell_NotifyIcon NIM_ADD, Icon_Data
    App.TaskVisible = False

    Me.Caption = "INet Counter - Offline"

    fraMenuSplit.Left = -100
    fraMenuSplit.Width = Me.Width + 200
    Start = True
    Set oReg = New oReg
    SetConMenu
End Sub

Private Sub Form_Resize()
Dim X As Integer
    X = frmMain.WindowState
    If X = 1 Then
        frmMain.Visible = False
    Else
        frmMain.Visible = True
    End If
End Sub

Private Sub mnuAuto_Click()
    If mnuAuto.Checked = False Then
        mnuAuto.Checked = True
        SaveSetting App.Title, "Settings", "Auto Login", mnuAuto.Checked
    Else
        mnuAuto.Checked = False
        SaveSetting App.Title, "Settings", "Auto Login", mnuAuto.Checked
    End If
End Sub

Private Sub mnuChange_Click()
    Online = ActiveConnection
    If Online = True Then SaveCurrent
End Sub

Private Sub mnuConect_Click(Index As Integer)
'*************************************
'Function Name: cmdView_Click
'Use: Start a connection
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-??-??
'*************************************

Dim RetVal As Long
    RetVal = Shell("rundll32.exe rnaui.dll,RnaDial " & mnuConect(Index).Caption, 1)
End Sub

Private Sub mnuExit_Click()
Dim RetVal As Integer
    RetVal = MsgBox("Do you want to Exit INet Counter?", vbYesNo, "Exit")
    If RetVal = vbNo Then Exit Sub
    Unload Me
    End
End Sub

Private Sub mnuQuit_Click()
Dim RetVal As Integer
    RetVal = MsgBox("Do you want to Exit INet Counter?", vbYesNo, "Exit")
    If RetVal = vbNo Then Exit Sub
    Unload Me
    End
End Sub

Private Sub mnuSetTariff_Click()
    frmSetTariff.Show , frmMain
End Sub

Private Sub mnuUserEditor_Click()
    frmAddUser.Show , frmMain
End Sub

Private Sub mnuViewLog_Click()
    frmLogShow.Show , frmMain
End Sub

Private Sub Timer1_Timer()
Dim RetVal As String
    Online = ActiveConnection
    If Online = False Then
        If INetData.OnTime <> "" Then
            mnuChange.Enabled = False
            SaveCurrent
            Me.Caption = "INet Counter - Offline"
            Icon_Data.szTip = "INet Counter - Offline"
        End If
        Exit Sub
    ElseIf Online = True Then
        If Icon_Data.hIcon = imgOnline Then
            Icon_Data.hIcon = imgOffline
            'change icon in systray
            Icon_Data.szTip = "INet Counter - Online Time: " & lblTime.Caption & Chr(0)
            Shell_NotifyIcon NIM_MODIFY, Icon_Data
        Else
            Icon_Data.hIcon = imgOnline
            'change icon in systray
            Icon_Data.szTip = "INet Counter - Online Time: " & lblTime.Caption & Chr(0)
            Shell_NotifyIcon NIM_MODIFY, Icon_Data
        End If
        INetCount
        Me.Caption = "INet Counter - Online"
    End If
End Sub

Private Sub INetCount()
'*************************************
'Function Name: INetCount
'Use: Count the time and price if the user is online
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-17
'*************************************
Dim vArray As Variant
Dim X As Integer
Dim TempTime As String
Dim Read As String
On Error GoTo ErrHandler
    If INetData.OnTime = "" Then
        INetData.OnTime = Now
        INetData.Price = 0
        INetData.Tariff = 0
        If GetConName <> "" Then
            INetData.ConName = GetConName
            lblConName.Caption = INetData.ConName
            INetData.TotPrice = GetConFee(lblConName.Caption)
        Else
            MsgBox "INet Counter was not able to get you Conection name." & _
            vbLf & "INet counter will not be terminated.", vbCritical
            CloseDataBase
            End
        End If
        INetData.CountTime = Now
        frmMain.Show
        If mnuAuto.Checked = False Then
ShowLogin:
            frmLogin.Show , frmMain
        Else
            Read = GetUserName
            vArray = GetUser(One, Read)
            If Not IsArray(vArray) Then GoTo ShowLogin
            INetData.User = vArray(0, 0)
            frmMain.lblName = Read
        End If
        mnuChange.Enabled = True
        sEnd = ""
    End If
    If (sEnd = "") Or (Time = sEnd) Then
        If sEnd <> "" Then
            INetData.TotPrice = INetData.TotPrice + DateDiff("s", INetData.CountTime, Now) * (INetData.Tariff / 60)
            INetData.CountTime = Now
        End If
        If chkHoliday.Value = 0 Then
            X = Weekday(Date, vbMonday)
            vArray = GetCurrentTariff("Day" & X, GetConName)
        ElseIf chkHoliday.Value = 1 Then
            vArray = GetCurrentTariff("Holiday", GetConName)
        End If
        If IsArray(vArray) = False Then
            MsgBox "Exit"
            Exit Sub
        End If
        For X = 0 To UBound(vArray, 2)
            If (vArray(1, X) <= Time) And (vArray(2, X) > Time) Then
                sEnd = vArray(2, X)
                INetData.Tariff = vArray(3, X)
                Exit For
            End If
        Next
    End If
    lblTime.Caption = Sec2Time(DateDiff("s", INetData.OnTime, Now))
    INetData.Price = INetData.TotPrice + (DateDiff("s", INetData.CountTime, Now) * (INetData.Tariff / 60))
    lblPrice = Cent2Dollar(INetData.Price)
Exit Sub
ErrHandler:
    MsgBox Err.Number
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Online = ActiveConnection
    If Online = True Then SaveCurrent
    Icon_Data.cbSize = Len(Icon_Data)
    Icon_Data.hwnd = picShow.hwnd  'Anger länken till bilden
    Icon_Data.uId = 1&
    Shell_NotifyIcon NIM_DELETE, Icon_Data 'Tarbort den från Systrayn
    CloseDataBase
    End
End Sub

Private Sub picShow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static Spara As Boolean
Dim Meddelande As Long
Dim RetVal As Long
Dim Flags As Long  ' the flags specifying how to move the window
Meddelande = X / Screen.TwipsPerPixelX
    If Spara = False Then
       Spara = True
         Select Case Meddelande
            Case WM_LBUTTONDBLCLK: 'Vänster musknapp dubbelklickning
'                Me.Visible = True
'                Flags = SWP_NOSIZE Or SWP_SHOWWINDOW
'                RetVal = SetWindowPos(Me.hwnd, HWND_TOP, Me.Left, Me.Top, Me.Left + Me.Width, Me.Top + Me.Height, Flags)
'                RetVal = ShowWindow(Me.hwnd, 9)
            Case WM_RBUTTONDBLCLK: 'right musknapp dubbelklickning
            Case WM_LBUTTONDOWN:   'left mousebutton down
                Me.Visible = True
                Flags = SWP_NOSIZE Or SWP_SHOWWINDOW
                RetVal = SetWindowPos(Me.hwnd, HWND_TOP, Me.Left, Me.Top, Me.Left + Me.Width, Me.Top + Me.Height, Flags)
                RetVal = ShowWindow(Me.hwnd, 9)
            Case WM_RBUTTONDOWN:   'Höger musknapp ner
            Case WM_LBUTTONUP:     'Vänster musknapp up
            Case WM_RBUTTONUP:     'Höger musknapp Up
                Me.PopupMenu mnuPopup
        End Select
        Spara = False
    End If
End Sub

Public Sub SaveCurrent()
'*************************************
'Function Name: SaveCurrent
'Use: Save current data
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-10-09
'*************************************
    INetData.Price = INetData.TotPrice + (DateDiff("s", INetData.CountTime, Now) * (INetData.Tariff / 60))
    SaveLog lblName.Caption
    INetData.OnTime = ""
    Icon_Data.hIcon = imgOffline
    Icon_Data.szTip = "INet Counter"
    Shell_NotifyIcon NIM_MODIFY, Icon_Data
    frmMain.Caption = "INet Counter"
End Sub

Public Sub SetConMenu()
'*************************************
'Function Name: SaveCurrent
'Use: Load all internet connections and put them in a menu
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-10-09
'*************************************

Dim sRetVal As String
Dim X As Long
Dim Y As Long

    sRetVal = oReg.EnumValue(HKEY_CURRENT_USER, "RemoteAccess\Addresses")
    X = 1
    Y = 0
    Do Until X = 0
        X = InStr(X, sRetVal, Chr(0))
        If X <> 0 Then
            mnuConect2.Visible = True
            If Y <> 0 Then Load mnuConect(Y)
            mnuConect(Y).Caption = Mid(sRetVal, 1, X - 1)
            mnuConect(Y).Visible = True
            sRetVal = Mid(sRetVal, X + 1)
            X = 1
            Y = Y + 1
        End If
    Loop
End Sub
