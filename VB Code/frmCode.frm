VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmCode 
   Caption         =   "Code Finder (C) Viktor Gars 1999"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8130
   Icon            =   "frmCode.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   855
      Left            =   4680
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
      Visible         =   0   'False
      _ExtentX        =   2566
      _ExtentY        =   1508
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      TextRTF         =   $"frmCode.frx":0442
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "Add New Tip"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageKey        =   "Exit"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2175
      Left            =   4440
      TabIndex        =   1
      Top             =   3600
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
   Begin MSComctlLib.ImageList imgToolbar 
      Left            =   6600
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCode.frx":0517
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCode.frx":0637
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCode.frx":074B
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCode.frx":085F
            Key             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTreeView 
      Left            =   7200
      Top             =   3120
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
            Picture         =   "frmCode.frx":0CB7
            Key             =   "icoClose"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCode.frx":0DD7
            Key             =   "icoOpen"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCode.frx":0EF7
            Key             =   "icoText"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6975
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   12303
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imgTreeView"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Height          =   5055
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuAddItem 
         Caption         =   "&Add Item"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteItem 
         Caption         =   "&Delete Item"
      End
   End
End
Attribute VB_Name = "frmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sShow As String
Private mbResizing As Boolean

Private Sub Form_Load()
Dim intCount
Dim X As Long
    OpenDataBase
    LoadTree
    WebBrowser1.Navigate2 App.Path & "\index.htm"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    TreeView1.Move 0, Toolbar1.Height, Me.ScaleWidth / 3, Me.ScaleHeight - Toolbar1.Height
    WebBrowser1.Move (Me.ScaleWidth / 3) + 50, Toolbar1.Height, (Me.ScaleWidth * 2 / 3) - 50, Me.ScaleHeight - Toolbar1.Height
    rtfText.Move (Me.ScaleWidth / 3) + 50, Toolbar1.Height, (Me.ScaleWidth * 2 / 3) - 50, Me.ScaleHeight - Toolbar1.Height
    Label1.Move Me.ScaleWidth / 3, 0, 100, Me.ScaleHeight
    Label1.MousePointer = vbSizeWE
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then mbResizing = True
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mbResizing Then
        Dim nX As Single
        nX = Label1.Left + X
        If nX < 500 Then Exit Sub
        If nX > Me.ScaleWidth - 500 Then Exit Sub
        TreeView1.Width = nX
        WebBrowser1.Left = nX + 50
        rtfText.Left = nX + 50
        WebBrowser1.Width = Me.ScaleWidth - nX - 50
        rtfText.Width = Me.ScaleWidth - nX - 50
        Label1.Left = nX
    End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mbResizing = False
End Sub

Private Sub mnuAddItem_Click()
Dim sName As String
    If TreeView1.SelectedItem.Key = "Root" Then
        sName = InputBox("Name")
        If sName <> "" Then
            AddMainItem sName
            LoadTree
        End If
    ElseIf Mid(TreeView1.SelectedItem.Key, 1, 4) = "Main" Then
        frmAddTip.Show vbModal, frmCode
        LoadTree
    End If
End Sub

Private Sub mnuDeleteItem_Click()
Dim RetVal As Integer
    RetVal = MsgBox("Do you want to delete this item?", vbYesNo + vbQuestion)
    If RetVal = vbYes Then
        If Mid(TreeView1.SelectedItem.Key, 1, 4) = "Tips" Then
            DeleteItem GetDBNr, GetTipNr
        ElseIf Mid(TreeView1.SelectedItem.Key, 1, 4) = "Main" Then
            DeleteTabell GetDBNr
        End If
        LoadTree
    End If
End Sub

Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Then Exit Sub
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Exit"
        End
    Case "Find"
        frmFind.Show , frmCode
    End Select
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub LoadTree()
'*************************************
'Function Name: LoadTree
'Use:
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-10-22
'*************************************
Dim vMain As Variant
Dim vTips As Variant
Dim X As Integer
Dim Y As Long
Dim nodX As Node
    TreeView1.Nodes.Clear
    Set nodX = TreeView1.Nodes.Add(, , "Root", "Code", "icoClose", "icoOpen")
    vMain = GetMain
    If Not IsArray(vMain) Then Exit Sub
    For X = 0 To UBound(vMain, 2)
        Set nodX = TreeView1.Nodes.Add("Root", tvwChild, "Main-" & vMain(0, X), vMain(1, X), "icoClose", "icoOpen")
        vTips = GetTips(vMain(0, X))
        If IsArray(vTips) Then
            For Y = 0 To UBound(vTips, 2)
                TreeView1.Nodes.Add "Main-" & vMain(0, X), tvwChild, "Tips-" & vTips(1, Y), vTips(0, Y), "icoText", "icoText"
            Next
        End If
    Next
    nodX.EnsureVisible
    TreeView1.Sorted = True
End Sub

Public Function GetDBNr() As String
Dim sTemp As String
Dim X As Long
    sTemp = Mid(TreeView1.SelectedItem.Key, 6)
    X = InStr(1, sTemp, "-")
    If X <> 0 Then
        sTemp = Mid(sTemp, 1, X)
    End If
    GetDBNr = sTemp
End Function

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
Dim sData As String
Dim FileNum As Integer
Dim i As Integer
    If Node.Key = "Root" Then
        mnuDeleteItem.Enabled = False
        mnuAddItem.Caption = "Add Tabell..."
        WebBrowser1.Navigate2 App.Path & "\index.htm"
    ElseIf Mid(TreeView1.SelectedItem.Key, 1, 4) = "Main" Then
        mnuDeleteItem.Enabled = True
        mnuDeleteItem.Caption = "Delete Tabell..."
        mnuAddItem.Caption = "Add Tip..."
        WebBrowser1.Navigate2 App.Path & "\index.htm"
    ElseIf Mid(TreeView1.SelectedItem.Key, 1, 4) = "Tips" Then
        mnuDeleteItem.Enabled = True
        mnuDeleteItem.Caption = "Delete Tip..."
        mnuAddItem.Caption = "Add Tip..."
        sData = GetTip(GetDBNr, GetTipNr)
        i = InStr(1, sData, "<html>", vbTextCompare)
        If i <> 0 Then
            On Error Resume Next
            Kill App.Path & "\temphtml.htm"
            FileNum = FreeFile
            Open App.Path & "\TempHtml.htm" For Binary As FileNum
            Put #FileNum, 1, sData
            Close FileNum
            rtfText.Visible = False
            WebBrowser1.Visible = True
            WebBrowser1.Navigate2 App.Path & "\temphtml.htm"
            TreeView1.SelectedItem.Selected = True
        Else
            On Error Resume Next
            Kill App.Path & "\temprtf.rtf"
            FileNum = FreeFile
            Open App.Path & "\Temprtf.rtf" For Binary As FileNum
            Put #FileNum, 1, sData
            Close FileNum
            WebBrowser1.Visible = False
            rtfText.Visible = True
            rtfText.FileName = App.Path & "\temprtf.rtf"
            TreeView1.SelectedItem.Selected = True
        End If
    End If
End Sub

Private Function GetTipNr() As String
Dim X As Long
Dim sTemp As String
    sTemp = Mid(TreeView1.SelectedItem.Parent.Key, 6)
    X = InStr(1, sTemp, "-")
    If X <> 0 Then
        sTemp = Mid(sTemp, 1, X)
    End If
    GetTipNr = sTemp
End Function
