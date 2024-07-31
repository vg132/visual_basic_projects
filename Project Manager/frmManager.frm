VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManager 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Project Manager"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   2580
   Icon            =   "frmManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   2580
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManager.frx":0442
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManager.frx":0556
            Key             =   "New"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManager.frx":066A
            Key             =   "cpp"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManager.frx":077E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManager.frx":0892
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManager.frx":09A6
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManager.frx":0ABA
            Key             =   "Remove"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManager.frx":0BCE
            Key             =   "java"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManager.frx":0CE2
            Key             =   "other"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManager.frx":0DF6
            Key             =   "txt"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManager.frx":0F0A
            Key             =   "Normal"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManager.frx":125E
            Key             =   "TopMost"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New Project"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open Project"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Project"
            ImageKey        =   "Save"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SaveMenu"
                  Text            =   "&Save"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SaveAsMenu"
                  Text            =   "Save &as"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add"
            Object.ToolTipText     =   "Add file to Project"
            ImageKey        =   "Add"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Remove"
            Object.ToolTipText     =   "Remove file from Project"
            ImageKey        =   "Remove"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Over"
            Object.ToolTipText     =   "Always on top"
            ImageKey        =   "Normal"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstFile 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Files"
         Object.Width           =   3651
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuRun 
         Caption         =   "&Run"
      End
      Begin VB.Menu mnuMake 
         Caption         =   "&Make File"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Options..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private TopMost As Boolean

Private Sub Form_Load()
    lstFile.Left = 0
    lstFile.Top = 360
    lstFile.Height = frmManager.Height - 300 - 50 - Toolbar1.Height
    lstFile.Width = frmManager.Width - 120
    
    Javac = GetSetting(App.Title, "Java", "Javac", "")
    Java = GetSetting(App.Title, "Java", "Java", "")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lstFile.Height = frmManager.Height - 300 - 50 - Toolbar1.Height
    lstFile.Width = frmManager.Width - 120
    lstFile.ColumnHeaders(1).Width = lstFile.Width - 345
End Sub

Private Sub lstFile_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopup
    End If
End Sub

Private Sub lstFile_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    Data.Clear
    Data.Files.Add lstFile.SelectedItem.Key
    Data.SetData , vbCFFiles
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuMake_Click()
    ChDir GetFilePart(lstFile.SelectedItem.Key, GetFilePath)
    X = Shell(Javac & " -d class " & GetFilePart(lstFile.SelectedItem.Key, GetFileName), vbNormalFocus)
End Sub

Private Sub mnuOpt_Click()
    frmOptions.Show vbModal, frmManager
End Sub

Private Sub mnuRun_Click()
Dim Temp As String
Dim FileName As String
    ChDir GetFilePart(lstFile.SelectedItem.Key, GetFilePath)
    Temp = GetFilePart(lstFile.SelectedItem.Key, GetExt)
    FileName = Mid(lstFile.SelectedItem.Text, 1, Len(lstFile.SelectedItem.Text) - Len(Temp) - 1)
    X = Shell(Java & " " & FileName, vbNormalFocus)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim X As Long
    Select Case Button.Key
    Case "New"
        lstFile.ListItems.Clear
        OpenFileName = ""
    Case "Add"
        AddFile
    Case "Save"
        SaveProject Me.hWnd
    Case "Open"
        OpenProject Me.hWnd
    Case "Remove"
        RemoveFile
    Case "Over"
        If TopMost = True Then
            MakeNormal Me.hWnd
            Button.Image = "Normal"
        Else
            MakeTopMost Me.hWnd
            Button.Image = "TopMost"
        End If
    End Select
End Sub

Private Sub MakeNormal(lngHwnd As Long)
    SetWindowPos lngHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    TopMost = False
End Sub

Private Sub MakeTopMost(lngHwnd As Long)
    SetWindowPos lngHwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    TopMost = True
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "SaveMenu"
        SaveProject Me.hWnd
    Case "SaveAsMenu"
        SaveProjectAs Me.hWnd
    End Select
End Sub
