VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finder v3.0"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStyle 
      Height          =   30
      Left            =   -100
      TabIndex        =   39
      Top             =   360
      Width           =   15000
   End
   Begin MSComctlLib.ImageList ToolBar 
      Left            =   3000
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":041E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0532
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":064A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":075E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0876
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":098E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DDE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraMusic 
      Caption         =   "Music Info"
      Height          =   6135
      Left            =   -10000
      TabIndex        =   20
      Top             =   410
      Width           =   4575
      Begin VB.CommandButton cmdDown 
         Height          =   735
         Left            =   4200
         Picture         =   "frmMain.frx":1232
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Move selected track down one place"
         Top             =   3360
         Width           =   255
      End
      Begin VB.CommandButton cmdUp 
         Height          =   735
         Left            =   4200
         Picture         =   "frmMain.frx":1577
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Move selected track up one place"
         Top             =   2640
         Width           =   255
      End
      Begin VB.ListBox lstTracks 
         Height          =   3180
         Left            =   135
         TabIndex        =   35
         ToolTipText     =   "Tracks"
         Top             =   1800
         Width           =   3960
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Record"
         Height          =   315
         Index           =   1
         Left            =   3240
         TabIndex        =   34
         ToolTipText     =   "Save record"
         Top             =   5700
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeleteRecord 
         Caption         =   "Delete &Record"
         Height          =   315
         Left            =   120
         TabIndex        =   33
         ToolTipText     =   "Delete record"
         Top             =   5700
         Width           =   1215
      End
      Begin VB.CommandButton cmdNewRecord 
         Caption         =   "&New Record"
         Height          =   315
         Left            =   1680
         TabIndex        =   32
         ToolTipText     =   "New record"
         Top             =   5700
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeleteAll 
         Caption         =   "Re&move All"
         Height          =   315
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Remove all tracks"
         Top             =   5100
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add Track"
         Height          =   315
         Left            =   3240
         TabIndex        =   30
         ToolTipText     =   "Add a track"
         Top             =   5100
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeleteTrack 
         Caption         =   "R&emove Track"
         Height          =   315
         Left            =   1680
         TabIndex        =   29
         ToolTipText     =   "Remove the selected track"
         Top             =   5100
         Width           =   1215
      End
      Begin VB.TextBox txtArtist 
         Height          =   285
         Left            =   140
         TabIndex        =   25
         ToolTipText     =   "Name of the artist"
         Top             =   1080
         Width           =   4315
      End
      Begin VB.TextBox txtRecName 
         Height          =   285
         Left            =   140
         TabIndex        =   24
         ToolTipText     =   "Record name"
         Top             =   480
         Width           =   4315
      End
      Begin VB.Label lblTracks 
         Caption         =   "Tracks"
         Height          =   255
         Left            =   140
         TabIndex        =   28
         Top             =   1560
         Width           =   4315
      End
      Begin VB.Label lblArtist 
         Caption         =   "Artist"
         Height          =   255
         Left            =   140
         TabIndex        =   27
         Top             =   840
         Width           =   4315
      End
      Begin VB.Label lblRecName 
         Caption         =   "Record Name"
         Height          =   255
         Left            =   140
         TabIndex        =   26
         Top             =   240
         Width           =   4315
      End
   End
   Begin VB.Frame fraFilm 
      Caption         =   "Film Info"
      Height          =   6135
      Left            =   3000
      TabIndex        =   19
      Top             =   410
      Width           =   4575
      Begin VB.CommandButton cmdDeleteFilm 
         Caption         =   "Delete Film"
         Height          =   315
         Left            =   120
         TabIndex        =   48
         Top             =   5700
         Width           =   1215
      End
      Begin VB.CommandButton cmdFilmNew 
         Caption         =   "New Film"
         Height          =   315
         Left            =   1680
         TabIndex        =   47
         Top             =   5700
         Width           =   1215
      End
      Begin VB.CommandButton cmdFilmSave 
         Caption         =   "&Save Film"
         Height          =   315
         Left            =   3240
         TabIndex        =   46
         Top             =   5700
         Width           =   1215
      End
      Begin VB.TextBox txtYear 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         TabIndex        =   41
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox cboFilmGrade 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   1200
         Width           =   2295
      End
      Begin VB.ComboBox cboFilmMedia 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtFilmDesc 
         Height          =   3375
         Left            =   120
         TabIndex        =   44
         Top             =   2160
         Width           =   4335
      End
      Begin VB.TextBox txtFilmTitle 
         Height          =   285
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label12 
         Caption         =   "Description about the film"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "Media:"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   1620
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Grade:"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Year:"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   885
         Width           =   1095
      End
      Begin VB.Label lblFilmTitle 
         AutoSize        =   -1  'True
         Caption         =   "Film Title"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fraGame 
      Caption         =   "Game Info"
      Height          =   6135
      Left            =   -10000
      TabIndex        =   11
      Top             =   410
      Width           =   4575
      Begin VB.CommandButton cmdDeleteGame 
         Caption         =   "&Delete Game"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   5700
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New Game"
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Top             =   5700
         Width           =   1215
      End
      Begin VB.TextBox txtMaker 
         Height          =   285
         Left            =   2640
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cboGrade 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1560
         Width           =   1815
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cboMedia 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Game"
         Height          =   315
         Index           =   0
         Left            =   3240
         TabIndex        =   8
         Top             =   5700
         Width           =   1215
      End
      Begin VB.TextBox txtDesc 
         Height          =   3045
         Left            =   140
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2520
         Width           =   4335
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   140
         TabIndex        =   1
         Top             =   480
         Width           =   2415
      End
      Begin VB.ComboBox cboFormat 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Media"
         Height          =   195
         Left            =   140
         TabIndex        =   18
         Top             =   1980
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Grade"
         Height          =   195
         Left            =   140
         TabIndex        =   17
         Top             =   1620
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Game Format"
         Height          =   195
         Left            =   140
         TabIndex        =   16
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Type of Game"
         Height          =   195
         Left            =   140
         TabIndex        =   15
         Top             =   900
         Width           =   1695
      End
      Begin VB.Label lblMaker 
         Caption         =   "Game Maker"
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblName 
         Caption         =   "Game Name"
         Height          =   255
         Left            =   140
         TabIndex        =   13
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Description about the game"
         Height          =   255
         Left            =   140
         TabIndex        =   12
         Top             =   2280
         Width           =   4335
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   12
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6000
      Left            =   0
      TabIndex        =   0
      Top             =   530
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   10583
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraAbout 
      Caption         =   "Program Info"
      Height          =   6135
      Left            =   -10000
      TabIndex        =   21
      Top             =   410
      Width           =   4575
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "The idé for this program came from Tobias Freij in 1998"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1560
         Width           =   4395
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Copyright © 1998-1999 Viktor Gars"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   4395
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Finder v3.0"
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
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   4395
      End
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Visible         =   0   'False
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuDelGame 
         Caption         =   "Delete Game"
      End
   End
   Begin VB.Menu mnuMusic 
      Caption         =   "Music"
      Visible         =   0   'False
      Begin VB.Menu mnuNewRec 
         Caption         =   "New Record"
      End
      Begin VB.Menu mnuDelRec 
         Caption         =   "Delete Record"
      End
   End
   Begin VB.Menu mnuFilm 
      Caption         =   "Film"
      Visible         =   0   'False
      Begin VB.Menu mnuNewFilm 
         Caption         =   "New Film"
      End
      Begin VB.Menu mnuDelFilm 
         Caption         =   "Delete Film"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFilmGrade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub cboFilmMedia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub cmdDeleteAll_Click()
    Responce = MsgBox("Do you want to delete ALL Tracks on this record?", vbQuestion + vbYesNo, "Finder")
    If Responce = vbYes Then
        lstTracks.Clear
    End If
End Sub

Private Sub cmdDeleteFilm_Click()
Dim Id As Long
    Responce = MsgBox("Do you want to delete this film from the database?", vbQuestion + vbYesNo, "Delete Film")
    If Responce = vbYes Then
        Id = Mid(TreeView1.SelectedItem.Key, 6, Len(TreeView1.SelectedItem.Key) - 5)
        oDB.DeleteFilm Id
        TreeView1.Nodes.Remove (TreeView1.SelectedItem.Index)
        FilmClick
    End If
End Sub

Private Sub cmdDeleteGame_Click()
Dim Id As Long
    Responce = MsgBox("Do you want to delete this game from the database?", vbQuestion + vbYesNo, "Delete Game")
    If Responce = vbYes Then
        Id = Mid(TreeView1.SelectedItem.Key, 6, Len(TreeView1.SelectedItem.Key) - 5)
        oDB.DeleteGame Id
        TreeView1.Nodes.Remove (TreeView1.SelectedItem.Index)
        GameClick
    End If
End Sub

Private Sub cmdDeleteRecord_Click()
Dim Id As Long
    Responce = MsgBox("Do you want to delete this record from the database?", vbQuestion + vbYesNo, "Delete Record")
    If Responce = vbYes Then
        Id = Mid(TreeView1.SelectedItem.Key, 7, Len(TreeView1.SelectedItem.Key) - 6)
        oDB.DeleteMusic Id
        TreeView1.Nodes.Remove (TreeView1.SelectedItem.Index)
        MusicClick
    End If
End Sub

Private Sub cmdDeleteTrack_Click()
Dim Start As Long
Dim Read As String
    If lstTracks.ListIndex <> "-1" Then
        On Error Resume Next
        Start = lstTracks.ListIndex
        lstTracks.RemoveItem (lstTracks.ListIndex)
        If Not Start = lstTracks.ListCount Then
            For X = Start To lstTracks.ListCount - 1
                Start = InStr(1, lstTracks.List(X), ".")
                Read = X + 1 & ". " & Mid(lstTracks.List(X), Start + 2, Len(lstTracks.List(X)) - Start)
                lstTracks.List(X) = Read
            Next
        End If
    End If
End Sub

Private Sub cmdDown_Click()
Dim Read As String
Dim Read2 As String
Dim Start As Long
    If lstTracks.ListIndex <> "-1" Then
        If Not lstTracks.ListIndex = lstTracks.ListCount - 1 Then
            X = lstTracks.ListIndex
            Read = lstTracks.List(X)
            Read2 = lstTracks.List(X + 1)
            
            Start = InStr(1, Read2, ".")
            lstTracks.List(X) = X + 1 & ". " & Mid(Read2, Start + 2, Len(Read2) - Start)
            Start = InStr(1, Read, ".")
            lstTracks.List(X + 1) = X + 2 & ". " & Mid(Read, Start + 2, Len(Read) - Start)
            lstTracks.Selected(X + 1) = True
        End If
    End If
End Sub

Private Sub cmdFilmNew_Click()
Dim Temp As Long
    En True
    Temp = oDB.NewFilm
    TreeView1.Nodes.Add "Film", tvwChild, "Film-" & Temp, "New Film", 1, 2
    TreeView1.Nodes("Film-" & Temp).Selected = True
    txtFilmDesc = ""
    txtYear.Text = ""
    txtFilmTitle = ""
    cboFilmMedia.ListIndex = 0
    cboFilmGrade.ListIndex = 0
    txtFilmTitle.SetFocus
End Sub

Private Sub cmdFilmSave_Click()
    SaveData
End Sub

Private Sub cmdNew_Click()
Dim Temp As Long
    En True
    Temp = oDB.NewGame
    TreeView1.Nodes.Add "Game", tvwChild, "Game-" & Temp, "New Game", 1, 2
    TreeView1.Nodes("Game-" & Temp).Selected = True
    TreeView1.SelectedItem.EnsureVisible
    txtName.Text = ""
    txtMaker.Text = ""
    txtDesc.Text = ""
    cboType.ListIndex = 0
    cboGrade.ListIndex = 0
    cboMedia.ListIndex = 0
    cboFormat.ListIndex = 0
    txtName.SetFocus
End Sub

Private Sub cmdNewRecord_Click()
Dim Temp As Long
     En True
     Temp = oDB.NewRecord
     TreeView1.Nodes.Add "Music", tvwChild, "Music-" & Temp, "New Record", 1, 2
     TreeView1.Nodes("Music-" & Temp).Selected = True
     TreeView1.SelectedItem.EnsureVisible
     MusicClick
End Sub

Private Sub cmdSave_Click(Index As Integer)
    SaveData
End Sub

Private Sub cmdUp_Click()
Dim Read As String
Dim Read2 As String
Dim Start As Long
    If lstTracks.ListIndex <> "-1" Then
        X = lstTracks.ListIndex
        Read = lstTracks.List(X)
        Read2 = lstTracks.List(X - 1)
        
        Start = InStr(1, Read2, ".")
        lstTracks.List(X) = X + 1 & ". " & Mid(Read2, Start + 2, Len(Read2) - Start)
        Start = InStr(1, Read, ".")
        lstTracks.List(X - 1) = X & ". " & Mid(Read, Start + 2, Len(Read) - Start)
        lstTracks.Selected(X - 1) = True
    End If
End Sub

Private Sub Command1_Click()
Dim Read As String
    Read = InputBox("Track Title", "Add Track Title")
    lstTracks.AddItem lstTracks.ListCount + 1 & ". " & Read
End Sub

Private Sub Form_Load()
    InitData
    AddData
    GetData
    TreeView1_NodeClick TreeView1.Nodes(TreeView1.SelectedItem.Index)
End Sub

Public Sub AddData()
    cboFormat.AddItem "PC"
    cboFormat.AddItem "Sony PlayStation"
    cboFormat.AddItem "N64"
    cboFormat.AddItem "Sega Dreamcast"
    cboFormat.AddItem "{PlayStation 2}"
    cboFormat.AddItem "{Nintendo Dolphin}"
    cboFormat.AddItem "Super Nintendo"
    cboFormat.AddItem "Nintendo GameBoy"
    cboFormat.AddItem "Nintendo 8-bit"
    cboFormat.AddItem "Sega Saturn"
    cboFormat.AddItem "Sega Mega Drive"
    cboFormat.AddItem "Sega Game Gear"
    cboFormat.AddItem "Sega Master System"
    cboFormat.AddItem "Atari Jaguar"
    cboFormat.AddItem "Atari Lynxs"
    cboFormat.AddItem "Other"

    cboGrade.AddItem "0"
    cboGrade.AddItem "1"
    cboGrade.AddItem "2"
    cboGrade.AddItem "3"
    cboGrade.AddItem "4"
    cboGrade.AddItem "5"

    cboFilmGrade.AddItem "0"
    cboFilmGrade.AddItem "1"
    cboFilmGrade.AddItem "2"
    cboFilmGrade.AddItem "3"
    cboFilmGrade.AddItem "4"
    cboFilmGrade.AddItem "5"

    cboMedia.AddItem "CD"
    cboMedia.AddItem "Cart"
    cboMedia.AddItem "HD (Hard Disk)"
    cboMedia.AddItem "DVD"
    cboMedia.AddItem "Diskett"
    cboMedia.AddItem "Zip"
    cboMedia.AddItem "Other"

    cboFilmMedia.AddItem "VHS"
    cboFilmMedia.AddItem "S-VHS"
    cboFilmMedia.AddItem "C-VHS"
    cboFilmMedia.AddItem "VCD"
    cboFilmMedia.AddItem "DVD"
    cboFilmMedia.AddItem "Movie"
    cboFilmMedia.AddItem "Super 8"
    cboFilmMedia.AddItem "Other"
    
    
    cboType.AddItem "Action"
    cboType.AddItem "Driving"
    cboType.AddItem "RPG"
    cboType.AddItem "Beat 'em up"
    cboType.AddItem "Shoot 'em up"
    cboType.AddItem "Simulator"
    cboType.AddItem "Adventure"
    cboType.AddItem "Platform"
    cboType.AddItem "Horror"
    cboType.AddItem "Other"
End Sub

Private Sub mnuDelFilm_Click()
    cmdDeleteFilm_Click
End Sub

Private Sub mnuDelGame_Click()
    cmdDeleteGame_Click
End Sub

Private Sub mnuDelRec_Click()
    cmdDeleteRecord_Click
End Sub

Private Sub mnuNewFilm_Click()
    cmdFilmNew_Click
End Sub

Private Sub mnuNewGame_Click()
    cmdNew_Click
End Sub

Private Sub mnuNewRec_Click()
    cmdNewRecord_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Exit"
            End
        Case "Copy"
            SendKeys "^C"
        Case "Cut"
            SendKeys "^X"
        Case "Paste"
            SendKeys "^V"
        Case "Undo"
            SendKeys "^Z"
        Case "Save"
            SaveData
        Case "Find"
            frmFind.Show , frmMain
        Case "Print"
            If fraAbout.Left = 3000 Then
                frmPrint.Tag = "About"
                frmPrint.Show vbModal, frmMain
            ElseIf fraGame.Left = 3000 Then
                frmPrint.Tag = "Game"
                frmPrint.Show vbModal, frmMain
            ElseIf fraMusic.Left = 3000 Then
                frmPrint.Tag = "Music"
                frmPrint.Show vbModal, frmMain
            Else
                frmPrint.Tag = "Film"
                frmPrint.Show vbModal, frmMain
            End If
    End Select
End Sub

Private Sub TreeView1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Temp As Long
Dim No As String
    If KeyCode = vbKeyF5 Then
        Temp = TreeView1.SelectedItem.Index
        GetData
        TreeView1.Nodes(Temp).Selected = True
        TreeView1.SelectedItem.EnsureVisible
    ElseIf KeyCode = vbKeyDelete Then
        Temp = InStr(1, TreeView1.SelectedItem.Key, "-")
        No = Mid(TreeView1.SelectedItem.Key, 1, 1)
        If (No = "G") And (Temp > 3) Then
            cmdDeleteGame_Click
        ElseIf (No = "M") And (Temp > 3) Then
            cmdDeleteRecord_Click
        ElseIf (No = "F") And (Temp > 3) Then
            cmdDeleteFilm_Click
        End If
    End If
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If fraGame.Left = 3000 Then
            PopupMenu mnuGame
        ElseIf fraMusic.Left = 3000 Then
            PopupMenu mnuMusic
        ElseIf fraFilm.Left = 3000 Then
            PopupMenu mnuFilm
        End If
    End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    NodeClick
End Sub

Public Sub GameClick()
Dim Id As Long
    If Len(TreeView1.SelectedItem.Key) > 4 Then
        Id = Mid(TreeView1.SelectedItem.Key, 6, Len(TreeView1.SelectedItem.Key) - 5)
        vArray = oDB.GetGameData(Id)
        txtName = "" & vArray(0, 0)
        cboType = "" & vArray(1, 0)
        txtMaker.Text = "" & vArray(2, 0)
        cboGrade = "" & vArray(3, 0)
        cboFormat = "" & vArray(4, 0)
        cboMedia = "" & vArray(5, 0)
        txtDesc.Text = "" & vArray(6, 0)
        En True
        Toolbar1.Buttons(2).Enabled = True
    Else
        txtName = ""
        cboType.ListIndex = 0
        txtMaker.Text = ""
        cboGrade.ListIndex = 0
        cboFormat.ListIndex = 0
        cboMedia.ListIndex = 0
        txtDesc.Text = ""
        En False
        Toolbar1.Buttons(2).Enabled = False
        cmdNew.Enabled = True
    End If
End Sub

Public Sub GetData()
    TreeView1.Nodes.Clear
    With TreeView1.Nodes
        .Add , , "Root", "Finder", 1, 2
        .Add "Root", tvwChild, "Film", "Film", 1, 2
        .Add "Root", tvwChild, "Music", "Music", 1, 2
        .Add "Root", tvwChild, "Game", "Games", 1, 2
    End With

    TreeView1.Nodes(2).Selected = True
    TreeView1.SelectedItem.EnsureVisible
    TreeView1.Nodes(1).Selected = True

    'Hämta Spel Data
    vArray = oDB.GetGame
    Y = oDB.GetNr
    If Y > 1 Then
        For X = 0 To Y - 1
            TreeView1.Nodes.Add "Game", tvwChild, "Game-" & vArray(1, X), vArray(0, X), 1, 2
        Next
    ElseIf Y = 1 Then
        TreeView1.Nodes.Add "Game", tvwChild, "Game-" & vArray(1, 0), vArray(0, 0), 1, 2
    End If
    
    Set vArray = Nothing
    'Hämta Music data
    vArray = oDB.GetMusicName
    Y = oDB.GetNr
    If Y > 1 Then
        For X = 0 To Y - 1
            TreeView1.Nodes.Add "Music", tvwChild, "Music-" & vArray(1, X), vArray(0, X), 1, 2
        Next
    ElseIf Y = 1 Then
        TreeView1.Nodes.Add "Music", tvwChild, "Music-" & vArray(1, 0), vArray(0, 0), 1, 2
    End If
    
    Set vArray = Nothing
    'Hämta Music Data
    vArray = oDB.GetFilmName
    Y = oDB.GetNr
    If Y > 1 Then
        For X = 0 To Y - 1
            TreeView1.Nodes.Add "Film", tvwChild, "Film-" & vArray(1, X), vArray(0, X), 1, 2
        Next
    ElseIf Y = 1 Then
        TreeView1.Nodes.Add "Film", tvwChild, "Film-" & vArray(1, 0), vArray(0, 0), 1, 2
    End If
End Sub

Public Sub MusicClick()
Dim Id As Long
Dim Read As String
Dim Start As Long
Dim Stopp As Long

    If Len(TreeView1.SelectedItem.Key) > 5 Then
        lstTracks.Clear
        Id = Mid(TreeView1.SelectedItem.Key, 7, Len(TreeView1.SelectedItem.Key) - 6)
        vArray = oDB.GetMusic(Id)
        txtRecName = "" & vArray(1, 0)
        txtArtist = "" & vArray(2, 0)
        Read = "" & vArray(3, 0)
        Stopp = 1
        Start = 1
        For X = 1 To 30
            Stopp = InStr(Start, Read, "|-|")
            If Stopp > Start Then
                Id = Stopp - Start
                lstTracks.AddItem X & ". " & Mid(Read, Start, Id)
                Start = Stopp + 3
            Else
                Exit For
            End If
        Next
        Toolbar1.Buttons(2).Enabled = True
        En True
    Else
        txtRecName.Text = ""
        Toolbar1.Buttons(2).Enabled = False
        En False
        txtArtist = ""
        lstTracks.Clear
        cmdNewRecord.Enabled = True
    End If
End Sub

Public Sub TextSelected()

Dim i As Integer
Dim oMyTextBox As Object

Set oMyTextBox = Screen.ActiveControl
    If TypeName(oMyTextBox) = "TextBox" Then
        i = Len(oMyTextBox.Text)
        oMyTextBox.SelStart = 0
        oMyTextBox.SelLength = i
    End If

End Sub

Private Sub txtArtist_GotFocus()
    TextSelected
End Sub

Private Sub txtArtist_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Sub txtDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
End Sub

Private Sub txtDesc_GotFocus()
    TextSelected
End Sub

Private Sub txtFilmDesc_GotFocus()
    TextSelected
End Sub

Private Sub txtFilmDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub


Private Sub txtFilmTitle_GotFocus()
    TextSelected
End Sub

Private Sub txtFilmTitle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Sub txtMaker_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub txtMaker_GotFocus()
    TextSelected
End Sub

Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub txtName_GotFocus()
    TextSelected
End Sub

Sub cboType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Sub cboMedia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Sub cboFormat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Sub cboGrade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Public Sub SaveData()
Dim Read As String
Dim Id As Long
    If fraMusic.Left = 3000 Then
        'Spara Music
        For X = 0 To lstTracks.ListCount - 1
            Id = InStr(1, lstTracks.List(X), ".")
            Read = Read & "|-|" & Mid(lstTracks.List(X), Id + 2, Len(lstTracks.List(X)) - Id)
        Next
        If Read <> "" Then
            Read = Mid(Read, 4, Len(Read) - 3) & "|-|"
        End If
        Id = Mid(TreeView1.SelectedItem.Key, 7, Len(TreeView1.SelectedItem.Key) - 6)
        oDB.SaveMusic Id, txtRecName, txtArtist, Read
        TreeView1.SelectedItem.Text = txtRecName.Text
    ElseIf fraGame.Left = 3000 Then
        'Spara spel data
        Id = Mid(TreeView1.SelectedItem.Key, 6, Len(TreeView1.SelectedItem.Key) - 5)
        oDB.SaveGame txtName.Text, cboType.Text, txtMaker.Text, cboGrade.Text, cboFormat.Text, cboMedia.Text, txtDesc.Text, Id
        TreeView1.SelectedItem.Text = txtName.Text
    ElseIf fraFilm.Left = 3000 Then
        Id = Mid(TreeView1.SelectedItem.Key, 6, Len(TreeView1.SelectedItem.Key) - 5)
        oDB.SaveFilm Id, txtFilmTitle.Text, cboFilmMedia.Text, txtDesc.Text, txtYear.Text, cboFilmGrade.Text
        TreeView1.SelectedItem.Text = txtFilmTitle.Text
    End If
End Sub

Private Sub txtRecName_GotFocus()
    TextSelected
End Sub

Private Sub txtRecName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub En(ByVal sInput As Boolean)
Dim oCtl As Control
        For Each oCtl In fraGame.Container
            If TypeOf oCtl Is TextBox Then
                oCtl.Enabled = sInput
            End If
            If TypeOf oCtl Is ComboBox Then
                oCtl.Enabled = sInput
            End If
            If TypeOf oCtl Is CommandButton Then
                oCtl.Enabled = sInput
            End If
            If TypeOf oCtl Is Label Then
                oCtl.Enabled = sInput
            End If
        Next
    Set oCtl = Nothing
End Sub

Public Sub NodeClick()
    If Mid(TreeView1.SelectedItem.Key, 1, 1) = "G" Then
        fraGame.Left = 3000
        fraMusic.Left = -10000
        fraFilm.Left = -10000
        fraAbout.Left = -10000
        Toolbar1.Buttons(2).Enabled = True
        GameClick
    ElseIf Mid(TreeView1.SelectedItem.Key, 1, 1) = "M" Then
        fraGame.Left = -10000
        fraMusic.Left = 3000
        fraFilm.Left = -10000
        fraAbout.Left = -10000
        Toolbar1.Buttons(2).Enabled = True
        MusicClick
    ElseIf Mid(TreeView1.SelectedItem.Key, 1, 1) = "R" Then
        fraGame.Left = -10000
        fraMusic.Left = -10000
        fraFilm.Left = -10000
        fraAbout.Left = 3000
        Toolbar1.Buttons(2).Enabled = False
        En True
    Else
        fraGame.Left = -10000
        fraMusic.Left = -10000
        fraFilm.Left = 3000
        fraAbout.Left = -10000
        Toolbar1.Buttons(2).Enabled = True
        FilmClick
    End If
End Sub

Public Sub FilmClick()
Dim Id As Long
    If Len(TreeView1.SelectedItem.Key) > 4 Then
        En True
        Id = Mid(TreeView1.SelectedItem.Key, 6, Len(TreeView1.SelectedItem.Key) - 5)
        vArray = oDB.GetFilmData(Id)
        txtFilmTitle.Text = "" & vArray(0, 0)
        cboFilmGrade.Text = "" & vArray(2, 0)
        cboFilmMedia.Text = "" & vArray(3, 0)
        txtFilmDesc.Text = "" & vArray(4, 0)
        txtYear.Text = "" & vArray(5, 0)
        En True
        Toolbar1.Buttons(2).Enabled = True
    Else
        txtFilmTitle.Text = ""
        cboFilmGrade.ListIndex = 0
        cboFilmMedia.ListIndex = 0
        txtFilmDesc.Text = ""
        En False
        Toolbar1.Buttons(2).Enabled = False
        cmdFilmNew.Enabled = True
    End If
End Sub

Private Sub txtYear_GotFocus()
    TextSelected
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub
