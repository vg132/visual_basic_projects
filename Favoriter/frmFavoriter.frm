VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmFavoriter 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Favoriter"
   ClientHeight    =   4200
   ClientLeft      =   2655
   ClientTop       =   2025
   ClientWidth     =   10020
   Icon            =   "frmFavoriter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   10020
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   13
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Ny"
            Object.ToolTipText     =   "Ny Favorit"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Spara"
            Object.ToolTipText     =   "Spara"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Move"
            Object.ToolTipText     =   "Klipp ut"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Kopiera"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Klistra in"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Ta bort"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Favorit Mapp"
            Object.ToolTipText     =   "Favorit Mapp"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "About"
            Object.ToolTipText     =   "Om Bookmark Editor"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Avsluta"
            Object.ToolTipText     =   "Avsluta"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   2535
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFFFFF&
      Height          =   2625
      Left            =   2760
      Pattern         =   "*.url"
      TabIndex        =   5
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtURL 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1080
      Width           =   8775
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   8775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "URL:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Namn:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFavoriter.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFavoriter.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFavoriter.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFavoriter.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFavoriter.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFavoriter.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFavoriter.frx":0AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFavoriter.frx":0BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFavoriter.frx":0EDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFavoriter.frx":11F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFavoriter.frx":150E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFavoriter.frx":1828
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArkiv 
      Caption         =   "Arkiv"
      Visible         =   0   'False
      Begin VB.Menu mnuFlytta 
         Caption         =   "Klipp ut"
      End
      Begin VB.Menu mnuKopiera 
         Caption         =   "Kopiera"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Klistra in"
      End
      Begin VB.Menu mnuTabort 
         Caption         =   "Ta bort"
      End
   End
   Begin VB.Menu mnuDir 
      Caption         =   "Dir"
      Visible         =   0   'False
      Begin VB.Menu mnuNyKatalog 
         Caption         =   "Ny Mapp"
      End
      Begin VB.Menu mnuTabortKatalogen 
         Caption         =   "Ta bort Mappen"
      End
   End
End
Attribute VB_Name = "frmFavoriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub SaveFile()
    
    On Error GoTo commonerror
    'Lägga till ändelsen .url till filens namn, om det inte redan finns
    FileSize = 0
    FileSize = Len(txtName.Text)
    FileName = txtName.Text
    FileSize = FileSize - 3
    If Mid(FileName, FileSize, 4) <> ".url" Then
        txtName.Text = txtName.Text + ".url"
    End If
    'Fixa fram rätt katalog och filnamn
    FileNum = FreeFile
    File = Dir1.Path + "\" + txtName.Text
    Open File For Output As FileNum
    'Skriva filen
    Print #FileNum, "[InternetShortcut]"
    Print #FileNum, "URL=" + txtURL.Text
    Print #FileNum, "Modified=2036FFC6D346BD01F2"
    Close FileNum
    File1.Refresh
    Exit Sub
commonerror:
    Select Case Err.Number
        Case 53 'File not Found
            MsgBox "File not Found"
        Case 32755 'Cansel selected
            Exit Sub
        Case Else
            MsgBox "Ett Fel"
    End Select
End Sub

Public Sub DeleteFile()
    'Först måste man kolla om den valda filen är sparad, annars går det inte att ta bort den
    If File <> "" Then
        Responce = MsgBox("Vill du ta bort denna Favorit?", vbYesNo, "Ta bort")
        If Responce = vbNo Then Exit Sub
        If Responce = vbYes Then
            Kill (File)
            NewURL
            File1.Refresh
        End If
    End If
End Sub

Private Sub cmdMapp_Click()
    SaveSetting "Favoriter", "Config", "Dir", Dir1.Path
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    'Ladda en fil
    FileNum = FreeFile
    FileSize = 0
    NewSize = 0
    On Error GoTo fileerror
    File = Dir1.Path + "\" + File1.FileName
    FileSize = FileLen(File)
    FileSize = FileSize - 24
    NewSize = FileSize - 31
    txtName = File1.FileName
    Open File For Binary As FileNum
    URL = String(NewSize, " ")
    Get #FileNum, 25, URL
    txtURL.Text = URL
    Close FileNum
    'Ta bort det sista i namnet, .url
    Temp = ""
    FileName = txtName.Text
    Nummer = Len(FileName)
    Nummer = Nummer - 4
    FileName = Temp + Mid(FileName, 1, Nummer)
    txtName.Text = FileName
    Exit Sub
fileerror:
    MsgBox "Ett Fel"
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Visa ändast popup meny när en fil är markerad
    On Error GoTo errorstop
    If (Button = 2) Then
        PopupMenu mnuArkiv
    End If
    Exit Sub
errorstop:
    MsgBox "Ett Fel"
End Sub

Private Sub Form_Load()
    On Error GoTo errortrap
    Toolbar1.Buttons(6).Enabled = False
    mnuPaste.Enabled = False
    MoveFile2 = 0
    Dir1.Path = GetSetting("Favoriter", "Config", "Dir", "Dir")
    Exit Sub
errortrap:
    MsgBox "Du har inte valt vilken katalog som dina Favoriter ligger i, gå till den katalogen och tryck på iconen 'Favorit Mapp', du behöver bara göra detta en gång. Din Favorit Mapp är troligtvis 'c:\Windows\Favoriter'", vbInformation, "Favorit katalog"
    Dir1.Path = "C:\windows"
End Sub

Private Sub mnuFlytta_Click()
    MoveFile2 = 1
    Copy = 1
    CopyFile
End Sub

Private Sub mnuKopiera_Click()
    Copy = 1
    CopyFile
End Sub

Private Sub mnuTabort_Click()
    DeleteFile
End Sub
Private Sub mnuPaste_click()
    Copy = 2
    CopyFile
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
    Case Is = "Ny"
        NewURL
    Case Is = "Spara"
        SaveFile
    Case Is = "Delete"
        DeleteFile
    Case Is = "Copy"
        Copy = 1
        MoveFile2 = 0
        CopyFile
    Case Is = "Move"
        MoveFile2 = 1
        Copy = 1
        CopyFile
    Case Is = "About"
        About
    Case Is = "Favorit Mapp"
        cmdMapp_Click
    Case Is = "Avsluta"
        End
    Case Is = "Paste"
        Copy = 2
        CopyFile
    End Select
End Sub

Public Sub NewURL()
    File = ""
    txtName = ""
    txtURL = ""
End Sub

Public Sub CopyFile()
    If File = "" Then
        MsgBox "Du måste markera en fil för att kunna flytta/kopiera den.", vbInformation
        Exit Sub
    End If
    If Copy = 1 Then
        CopyFile2 = File
        CopyFileName = txtName.Text
        Toolbar1.Buttons(6).Enabled = True
        mnuPaste.Enabled = True
    End If
    If Copy = 2 Then
        FileCopy CopyFile2, Dir1.Path + "\" + CopyFileName + ".url"
        File1.Refresh
        If MoveFile2 = 1 Then
            Kill (CopyFile2)
            MoveFile2 = 0
            Toolbar1.Buttons(6).Enabled = False
            mnuPaste.Enabled = False
        End If
    End If
End Sub

Public Sub About()
    frmHelp.Show
End Sub

Private Sub mnuNyKatalog_Click()
    'On Error GoTo errortrap
    Temp = InputBox("Namn på katalogen", "Ny Katalog")
    Temp = Dir1.Path + "\" + Temp
    MkDir Temp
    Dir1.Refresh
    Exit Sub
errortrap:
    Select Case Err.Number
        Case 75 'Katalogen Finns redan
            MsgBox "Katalogen finns redan.", vbOKOnly, "Ny Katalog"
        Case Else
            MsgBox "Ett Fel"
    End Select
End Sub

Private Sub Dir1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Visa ändast popup meny när en fil är markerad
    On Error GoTo errorstop
    If Button = 2 Then
        PopupMenu mnuDir
    End If
    Exit Sub
errorstop:
    MsgBox "Ett Fel"
End Sub

Private Sub mnuTabortKatalogen_Click()
   On Error GoTo errortrap
    'Fråga om användaren vill ta bort mappen
    Responce = MsgBox("Vill du ta bort mappen " + Dir1.List(Nummer + Dir1.ListIndex) + "?", vbYesNo, "Ta bort Mappen")
    If Responce = vbYes Then
        Temp = Dir1.List(Nummer + Dir1.ListIndex)
        If Temp = Dir1.Path Then
            Dir1.Path = Dir1.List(Nummer + Dir1.ListIndex - 1)
        End If
        'txtName.Text = Dir1.ListIndex
        RmDir (Temp)
        Dir1.Refresh
    End If
    Exit Sub
errortrap:
    Select Case Err.Number
        Case 75 'Katalogen Finns redan
            MsgBox "Du kan inte ta bort denna katalog, den är troligtvis skrivarskyddad", vbOKOnly
        Case Else
            MsgBox "Ett Fel!"
    End Select
End Sub
