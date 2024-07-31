VERSION 5.00
Begin VB.Form frmId3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VG Software - ID3 Tag Edit"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9000
   Icon            =   "frmId3.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTop 
      Height          =   30
      Left            =   -20
      TabIndex        =   20
      Top             =   0
      Width           =   9000
   End
   Begin VB.CheckBox chkAutoSave 
      Caption         =   "A&uto Save"
      Height          =   195
      Left            =   6240
      TabIndex        =   12
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   ">>"
      Height          =   315
      Left            =   2520
      TabIndex        =   8
      ToolTipText     =   "Show More..."
      Top             =   3120
      Width           =   375
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   6240
      TabIndex        =   11
      Top             =   2760
      Width           =   2655
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   3120
      Pattern         =   "*.mp3"
      TabIndex        =   9
      Top             =   120
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   6240
      TabIndex        =   10
      Top             =   120
      Width           =   2655
   End
   Begin VB.Frame fraSplit 
      Height          =   3600
      Left            =   3000
      TabIndex        =   19
      Top             =   -100
      Width           =   30
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      ToolTipText     =   "Save ID3 Info"
      Top             =   3120
      Width           =   1035
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open..."
      Height          =   315
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Open MP3 File"
      Top             =   3120
      Width           =   1035
   End
   Begin VB.ComboBox cboGenre 
      Height          =   315
      ItemData        =   "frmId3.frx":0442
      Left            =   1200
      List            =   "frmId3.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Genre"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtComments 
      Height          =   285
      Left            =   120
      MaxLength       =   30
      TabIndex        =   5
      ToolTipText     =   "Comments"
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox txtYear 
      Height          =   315
      Left            =   120
      MaxLength       =   4
      TabIndex        =   3
      ToolTipText     =   "Year"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtAlbum 
      Height          =   285
      Left            =   120
      MaxLength       =   30
      TabIndex        =   2
      ToolTipText     =   "Album"
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox txtArtist 
      Height          =   285
      Left            =   120
      MaxLength       =   30
      TabIndex        =   1
      ToolTipText     =   "Artist"
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   120
      MaxLength       =   30
      TabIndex        =   0
      ToolTipText     =   "Song Title"
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "&Genre"
      Height          =   195
      Left            =   1245
      TabIndex        =   18
      Top             =   1920
      Width           =   465
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "&Comments"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   765
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "&Year"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   345
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Alb&um"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "&Artist"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "&Title"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   315
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmId3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function WinHelp Lib "user32.dll" Alias "WinHelpA" (ByVal hWndMain As Long, ByVal lpHelpFile As String, ByVal uCommand As Long, dwData As Any) As Long

Const HELP_CONTENTS = &H3
Const HELP_INDEX = &H3
Const HELP_FORCEFILE = &H9
Const HELP_CONTEXT = &H1

Private Enum GetFilePartEnum
    GetExt = 0
    GetFileName = 1
    GetFilePath = 2
End Enum

Dim FileName As String

Private Sub chkAutoSave_Click()
    SaveSetting App.Title, "Settings", "AutoSave", chkAutoSave.Value
End Sub

Private Sub cmdMore_Click()
Dim X
    If cmdMore.Caption = ">>" Then
        For X = 0 To 39
            Me.Width = Me.Width + 150
            Me.Refresh
        Next
        cmdMore.Caption = "<<"
        cmdMore.ToolTipText = "Show Less..."
        SaveSetting App.Title, "Settings", "More", "True"
    Else
        For X = 0 To 39
            Me.Width = Me.Width - 150
            Me.Refresh
        Next
        cmdMore.Caption = ">>"
        cmdMore.ToolTipText = "Show More..."
        SaveSetting App.Title, "Settings", "More", "False"
    End If
End Sub

Private Sub cmdOpen_Click()
    FileName = ShowOpen("MP3 Files (*.mp3)|*.mp3|All Files (*.*)|*.*|", Me.hWnd, App.Path, "Open MP3 File")
    If FileName <> "" Then
        GetId3 FileName
        txtTitle = RTrim(id3Info.Title)
        txtArtist = RTrim(id3Info.Artist)
        txtAlbum = RTrim(id3Info.Album)
        txtYear = RTrim(id3Info.sYear)
        txtComments = RTrim(id3Info.Comments)
        cboGenre.ListIndex = id3Info.Genre
        cmdSave.Enabled = True
        If bInfo = False Then
            txtTitle.Text = GetFilePart(FileName, GetFileName)
            txtTitle.Text = Mid(txtTitle.Text, 1, Len(txtTitle.Text) - 4)
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    id3Info.Title = txtTitle
    id3Info.Artist = txtArtist
    id3Info.Album = txtAlbum
    id3Info.sYear = txtYear
    id3Info.Comments = txtComments
    id3Info.Genre = cboGenre.ListIndex
    On Error GoTo ErrHandle
    SaveId3 FileName, id3Info
    Exit Sub
ErrHandle:
    If Err.Number = 75 Then
        MsgBox "File is Write Protected"
        Close #1
    Else
        MsgBox Err.Description
        Close #1
    End If
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    If (chkAutoSave.Value = 1) And (FileName <> "") Then
        cmdSave_Click
    End If
    If Right(File1.Path, 1) = "\" Then
        FileName = File1.Path & File1.FileName
    Else
        FileName = File1.Path & "\" & File1.FileName
    End If
    If FileName <> "" Then
        GetId3 FileName
        If RTrim(id3Info.Title) = "" Then
            txtTitle.Text = Mid(File1.FileName, 1, Len(File1.FileName) - 4)
        Else
            txtTitle = RTrim(id3Info.Title)
        End If
        txtArtist = RTrim(id3Info.Artist)
        txtAlbum = RTrim(id3Info.Album)
        txtYear = RTrim(id3Info.sYear)
        txtComments = RTrim(id3Info.Comments)
        On Error Resume Next
        cboGenre.ListIndex = id3Info.Genre
        cmdSave.Enabled = True
        If bInfo = False Then
            txtTitle.Text = GetFilePart(FileName, GetFileName)
            txtTitle.Text = Mid(txtTitle.Text, 1, Len(txtTitle.Text) - 4)
        End If
    End If
End Sub

Private Sub Form_Load()
Dim I As Long
Dim sGet As String
Dim CheckDir As Boolean
    GenreArray = Split(sGenreMatrix, "|")
    For I = LBound(GenreArray) To UBound(GenreArray)
        cboGenre.AddItem GenreArray(I)
    Next
    sGet = GetSetting(App.Title, "Settings", "Path", "")
    If sGet <> "" Then
        CheckDir = FileExists(sGet)
        If CheckDir = True Then
            Drive1.Drive = Mid(sGet, 1, 2)
            Dir1.Path = sGet
            sGet = ""
        Else
            Drive1.Drive = Mid(App.Path, 1, 2)
            Dir1.Path = App.Path
            sGet = ""
        End If
    End If
    sGet = GetSetting(App.Title, "Settings", "More", "False")
    If sGet = "True" Then
        Me.Width = 9090
        cmdMore.Caption = "<<"
        cmdMore.ToolTipText = "Show Less..."
    Else
        Me.Width = 3045
    End If
    I = GetSetting(App.Title, "Settings", "AutoSave", 0)
    chkAutoSave.Value = I
    FileName = ""
    App.HelpFile = App.Path & "\Help.hlp"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FileName = Dir1.Path
    SaveSetting App.Title, "Settings", "Path", FileName
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, frmId3
End Sub

Private Sub mnuExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuHelp_Click()
Dim RetVal
    RetVal = WinHelp(frmId3.hWnd, App.HelpFile, HELP_CONTEXT, ByVal 2)
End Sub

Private Sub mnuOpen_Click()
    cmdOpen_Click
End Sub

Private Sub mnuSave_Click()
    cmdSave_Click
End Sub

Private Sub txtAlbum_GotFocus()
    SelectText
End Sub

Private Sub txtArtist_GotFocus()
    SelectText
End Sub

Private Sub txtComments_GotFocus()
    SelectText
End Sub

Private Sub txtTitle_GotFocus()
    SelectText
End Sub

Private Sub txtYear_GotFocus()
    SelectText
End Sub

Public Function FileExists(ByVal PathName As String) As Boolean
       FileExists = IIf(Dir$(PathName) = "", False, True)
End Function

Private Function GetFilePart(ByVal File As String, ByVal Info As GetFilePartEnum) As String
Dim X As Integer
    If Len(File) <= 3 Then
      GetFilePart = File
      Exit Function
    End If
    Select Case Info
    Case 0
        GetFilePart = LCase(Mid(File, Len(File) - 3))
    Case 1
        For X = Len(File) To 1 Step -1
            If Mid(File, X, 1) = "\" Then Exit For
        Next
        GetFilePart = Mid(File, X + 1)
    Case 2
        For X = Len(File) To 1 Step -1
            If Mid(File, X, 1) = "\" Then Exit For
        Next
        GetFilePart = Mid(File, 1, X - 1)
    End Select
End Function
