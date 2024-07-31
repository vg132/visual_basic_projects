VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Code Safe"
   ClientHeight    =   7470
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   9630
      _CBHeight       =   390
      _Version        =   "6.0.8169"
      Child1          =   "Toolbar1"
      MinHeight1      =   330
      Width1          =   3600
      NewRow1         =   0   'False
      Child2          =   "Toolbar2"
      MinHeight2      =   330
      Width2          =   2655
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   3795
         TabIndex        =   6
         Top             =   30
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imgToolbar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
         Begin VB.ComboBox cboSize 
            Height          =   315
            Left            =   2280
            TabIndex        =   8
            Text            =   "Combo2"
            Top             =   0
            Width           =   855
         End
         Begin VB.ComboBox cboFont 
            Height          =   315
            Left            =   0
            Sorted          =   -1  'True
            TabIndex        =   7
            Text            =   "Combo1"
            Top             =   0
            Width           =   2175
         End
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   165
         TabIndex        =   5
         Top             =   30
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imgToolbar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               Object.ToolTipText     =   "New"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Object.ToolTipText     =   "Save Item"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Delete"
               Object.ToolTipText     =   "Delete Item"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cut"
               Object.ToolTipText     =   "Cut"
               ImageKey        =   "Cut"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Object.ToolTipText     =   "Copy"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paste"
               Object.ToolTipText     =   "Paste"
               ImageKey        =   "Paste"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Undo"
               Object.ToolTipText     =   "Undo"
               ImageKey        =   "Undo"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Exit"
               Object.ToolTipText     =   "Exit"
               ImageKey        =   "Exit"
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picRight 
      Align           =   4  'Align Right
      Height          =   7080
      Left            =   3900
      ScaleHeight     =   7020
      ScaleWidth      =   5670
      TabIndex        =   2
      Top             =   390
      Width           =   5730
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   2175
         Left            =   1080
         TabIndex        =   9
         Top             =   3960
         Width           =   3135
         ExtentX         =   5530
         ExtentY         =   3836
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3135
         Left            =   840
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   5530
         _Version        =   393217
         BorderStyle     =   0
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgToolbar 
      Left            =   120
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":00C9
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":01E9
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0309
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":041D
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0875
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0995
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AB5
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BD5
            Key             =   "Paste"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTreeView 
      Left            =   720
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CF1
            Key             =   "icoClose"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E11
            Key             =   "icoOpen"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F31
            Key             =   "icoText"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLeft 
      Align           =   3  'Align Left
      Height          =   7080
      Left            =   0
      ScaleHeight     =   7020
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   390
      Width           =   3975
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   735
         Left            =   480
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   6975
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   12303
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgTreeView"
         Appearance      =   0
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuAddItem 
         Caption         =   "Add &Item"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteItem 
         Caption         =   "&Delete Item"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private sShow As String
'
'Private Sub cboFont_Click()
'    If cboFont.Text <> "" Then
'        RichTextBox1.SelFontName = cboFont.Text
'        'RichTextBox1.Font = cboFont.Text
'    End If
'End Sub
'
'Private Sub cboSize_Change()
'    If cboSize.Text <> "" Then
'        RichTextBox1.SelFontSize = cboSize.Text
'        'RichTextBox1.Font = cboFont.Text
'    End If
'End Sub
'
'Private Sub cboSize_Click()
'    If cboSize.Text <> "" Then
'        RichTextBox1.SelFontSize = cboSize.Text
'        'RichTextBox1.Font = cboFont.Text
'    End If
'End Sub
'
'Private Sub Command1_Click()
'    frmCode.Show
'End Sub
'
'Private Sub MDIForm_Load()
'Dim intCount
'Dim X As Long
'    cboFont.Clear
'    For intCount = 1 To Screen.FontCount
'        cboFont.AddItem Screen.Fonts(intCount)
'    Next intCount
'    cboSize.Clear
'    For intCount = 1 To 10
'        X = X + 2
'        cboSize.AddItem X
'    Next
'    cboSize.Text = 10
'    cboFont.Text = "Arial"
'    OpenDataBase
'    LoadTree
'End Sub
'
'Private Sub MDIForm_Resize()
'    picRight.Width = Me.Width - picLeft.Width
'End Sub
'
'Private Sub mnuAddItem_Click()
'Dim sName As String
'    If TreeView1.SelectedItem.Key = "Root" Then
'        sName = InputBox("Name")
'        If sName <> "" Then
'            AddMainItem sName
'            LoadTree
'        End If
'    ElseIf Mid(TreeView1.SelectedItem.Key, 1, 4) = "Main" Then
'        sName = InputBox("Tips Name")
'        If sName <> "" Then
'            RichTextBox1.TextRTF = ""
'            NewItem sName, GetDBNr
'            sShow = sName
'            LoadTree
'            sShow = ""
'        End If
'    End If
'End Sub
'
'Private Sub mnuDeleteItem_Click()
'Dim RetVal As Integer
'    RetVal = MsgBox("Do you want to delete this item?", vbYesNo + vbQuestion)
'    If RetVal = vbYes Then
'        If Mid(TreeView1.SelectedItem.Key, 1, 4) = "Tips" Then
'            DeleteItem GetDBNr, GetTipNr
'        ElseIf Mid(TreeView1.SelectedItem.Key, 1, 4) = "Main" Then
'            DeleteTabell GetDBNr
'        End If
'        LoadTree
'        RichTextBox1.TextRTF = ""
'    End If
'End Sub
'
'Private Sub picLeft_Resize()
'    TreeView1.Width = picLeft.Width - 170
'    TreeView1.Height = picLeft.Height
'    TreeView1.Left = 0
'    TreeView1.Top = 0
'End Sub
'
'Private Sub picright_Resize()
'    RichTextBox1.Width = picRight.Width - 100
'    RichTextBox1.Height = picRight.Height
'    RichTextBox1.Left = 50
'    RichTextBox1.Top = 0
'    WebBrowser1.Width = picRight.Width - 100
'    WebBrowser1.Height = picRight.Height
'    WebBrowser1.Left = 50
'    WebBrowser1.Top = 0
'End Sub
'
'Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 9 Then Exit Sub
'End Sub
'
'Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'    Select Case Button.Key
'    Case "Exit"
'        End
'    Case "Copy"
'        SendKeys "^c"
'    Case "Cut"
'        SendKeys "^x"
'    Case "Paste"
'        SendKeys "^v"
'    Case "Undo"
'        SendKeys "^z"
'    Case "New"
'        mnuAddItem_Click
'    Case "Save"
'        SaveTip GetDBNr, GetTipNr, RichTextBox1.TextRTF
'    End Select
'End Sub
'
'Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 2 Then PopupMenu mnuPopup
'End Sub
'
'Private Sub LoadTree()
''*************************************
''Function Name: LoadTree
''Use:
''Remarks:
''History:
''Programmer: Viktor Gars
''Date: 1999-10-22
''*************************************
'Dim vMain As Variant
'Dim vTips As Variant
'Dim X As Integer
'Dim Y As Long
'Dim nodX As Node
'    TreeView1.Nodes.Clear
'    Set nodX = TreeView1.Nodes.Add(, , "Root", "Code", "icoClose", "icoOpen")
'    vMain = GetMain
'    If Not IsArray(vMain) Then Exit Sub
'    For X = 0 To UBound(vMain, 2)
'        Set nodX = TreeView1.Nodes.Add("Root", tvwChild, "Main-" & vMain(0, X), vMain(1, X), "icoClose", "icoOpen")
'        vTips = GetTips(vMain(0, X))
'        If IsArray(vTips) Then
'            For Y = 0 To UBound(vTips, 2)
'                TreeView1.Nodes.Add "Main-" & vMain(0, X), tvwChild, "Tips-" & vTips(1, Y), vTips(0, Y), "icoText", "icoText"
'            Next
'        End If
'    Next
'    nodX.EnsureVisible
'End Sub
'
'Private Function GetDBNr() As String
'Dim sTemp As String
'Dim X As Long
'    sTemp = Mid(TreeView1.SelectedItem.Key, 6)
'    X = InStr(1, sTemp, "-")
'    If X <> 0 Then
'        sTemp = Mid(sTemp, 1, X)
'    End If
'    GetDBNr = sTemp
'End Function
'
'Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
'Dim sData As String
'Dim FileNum As Integer
'    If Node.Key = "Root" Then
'        mnuDeleteItem.Enabled = False
'        mnuAddItem.Caption = "Add Tabell..."
'        RichTextBox1.TextRTF = ""
'    ElseIf Mid(TreeView1.SelectedItem.Key, 1, 4) = "Main" Then
'        mnuDeleteItem.Enabled = True
'        mnuDeleteItem.Caption = "Delete Tabell..."
'        mnuAddItem.Caption = "Add Tip..."
'        RichTextBox1.TextRTF = ""
'    ElseIf Mid(TreeView1.SelectedItem.Key, 1, 4) = "Tips" Then
'        mnuDeleteItem.Enabled = True
'        mnuDeleteItem.Caption = "Delete Tip..."
'        mnuAddItem.Caption = "Add Tip..."
'        sData = GetTip(GetDBNr, GetTipNr)
'        If LCase(Mid(sData, 1, 6)) = "<html>" Then
'            On Error Resume Next
'            Kill App.Path & "\temphtml.htm"
'            FileNum = FreeFile
'            Open App.Path & "\TempHtml.htm" For Binary As FileNum
'            Put #FileNum, 1, sData
'            Close FileNum
'            WebBrowser1.Navigate2 App.Path & "\temphtml.htm"
'            WebBrowser1.Visible = True
'            RichTextBox1.Visible = False
'            TreeView1.SelectedItem.Selected = True
'        Else
'            WebBrowser1.Visible = False
'            RichTextBox1.Visible = True
'            RichTextBox1.TextRTF = sData
'        End If
'    End If
'End Sub
'
'Private Function GetTipNr() As String
'Dim X As Long
'Dim sTemp As String
'    sTemp = Mid(TreeView1.SelectedItem.Parent.Key, 6)
'    X = InStr(1, sTemp, "-")
'    If X <> 0 Then
'        sTemp = Mid(sTemp, 1, X)
'    End If
'    GetTipNr = sTemp
'End Function
