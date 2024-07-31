VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLogShow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INet Counter - Log View"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ListView lstLog 
      Height          =   5175
      Left            =   1920
      TabIndex        =   15
      Top             =   60
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9128
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Log On Date/Time"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Logoff Date/Time"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Time"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Price"
         Object.Width           =   1656
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Conection Name"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.ComboBox cboDay 
      Height          =   315
      Index           =   1
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   5475
      Width           =   615
   End
   Begin VB.ComboBox cboMonth 
      Height          =   315
      Index           =   1
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   5475
      Width           =   1215
   End
   Begin VB.ComboBox cboYear 
      Height          =   315
      Index           =   1
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   5475
      Width           =   735
   End
   Begin VB.ComboBox cboYear 
      Height          =   315
      Index           =   0
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5460
      Width           =   735
   End
   Begin VB.ComboBox cboMonth 
      Height          =   315
      Index           =   0
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   5460
      Width           =   1215
   End
   Begin VB.ComboBox cboDay 
      Height          =   315
      Index           =   0
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   5460
      Width           =   615
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Default         =   -1  'True
      Height          =   315
      Left            =   7440
      TabIndex        =   2
      Top             =   5475
      Width           =   1035
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   5175
      Left            =   60
      TabIndex        =   16
      Top             =   60
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   9128
      _Version        =   327682
      Indentation     =   353
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLogShow.frx":0000
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLogShow.frx":0112
            Key             =   "Open"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Month:"
      Height          =   195
      Index           =   1
      Left            =   5520
      TabIndex        =   14
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Year:"
      Height          =   195
      Index           =   1
      Left            =   4800
      TabIndex        =   13
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Day:"
      Height          =   195
      Index           =   1
      Left            =   6720
      TabIndex        =   12
      Top             =   5280
      Width           =   330
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Day:"
      Height          =   195
      Index           =   0
      Left            =   3840
      TabIndex        =   8
      Top             =   5265
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Year:"
      Height          =   195
      Index           =   0
      Left            =   1920
      TabIndex        =   7
      Top             =   5265
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Month:"
      Height          =   195
      Index           =   0
      Left            =   2640
      TabIndex        =   6
      Top             =   5265
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "to:"
      Height          =   195
      Left            =   4530
      TabIndex        =   1
      Top             =   5580
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "View Logdata from:"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   5580
      Width           =   1365
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Logdata"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Item"
      End
   End
End
Attribute VB_Name = "frmLogShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sUName As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cboMonth_Click(Index As Integer)
Dim X As Byte
Dim RetVal As Boolean
    cboDay(Index).Clear
    X = cboMonth(Index).ListIndex + 1
    If X = 2 Then
        RetVal = IsLeapYear(cboYear(Index))
        If RetVal = True Then
            For X = 1 To 29
                cboDay(Index).AddItem X
            Next
        Else
            For X = 1 To 28
                cboDay(Index).AddItem X
            Next
        End If
    Else
        If (X = 1) Or (X = 3) Or (X = 5) Or (X = 7) Or (X = 8) Or (X = 10) Or (X = 12) Then
            For X = 1 To 31
                cboDay(Index).AddItem X
            Next
        Else
            For X = 1 To 30
                cboDay(Index).AddItem X
            Next
        End If
    End If
    cboDay(Index).ListIndex = 0
End Sub

Private Sub cboYear_Click(Index As Integer)
    If (IsLeapYear(cboYear(Index).Text) = True) And (cboMonth(Index).ListIndex = 1) Then
        cboMonth_Click (Index)
    End If
End Sub

Private Sub cmdView_Click()
'*************************************
'Function Name: cmdView_Click
'Use: View data from a specific date
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-??-??
'*************************************

Dim vArray As Variant
Dim tmpDate(0 To 1) As Date
    If TreeView1.SelectedItem.Text = "Users" Then
        lstLog.ListItems.Clear
        Exit Sub
    End If
    Me.MousePointer = 11
    tmpDate(0) = cboYear(0).Text & "-" & cboMonth(0).ListIndex + 1 & "-" & cboDay(0).Text
    tmpDate(1) = cboYear(1).Text & "-" & cboMonth(1).ListIndex + 1 & "-" & cboDay(1).Text
    lstLog.ListItems.Clear
    vArray = GetLog(TreeView1.SelectedItem.Text, tmpDate(0), tmpDate(1))
    ShowLog vArray
    Me.MousePointer = 0
End Sub

Private Sub Form_Load()
'*************************************
'Function Name: Form_Load
'Use: Load all users and show them in the tree
''Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-??-??
'*************************************

Dim nodX As Node
Dim vArray As Variant
Dim X As Long
    sUName = ""
    TreeView1.Nodes.Add , , "Root", "Users", 1, 2
    vArray = GetUser(All)
    If IsArray(vArray) = False Then Exit Sub
    For X = 0 To UBound(vArray, 2)
        Set nodX = TreeView1.Nodes.Add("Root", tvwChild, "U" & vArray(1, X), vArray(1, X), "Close", "Open")
        sUName = sUName & vArray(1, X) & ";"
    Next
    TreeView1.Nodes.Add "Root", tvwChild, "All", "All users", "Close", "Open"
    nodX.EnsureVisible

    cmdView.Enabled = False

    For X = 1990 To 2050
        cboYear(0).AddItem X
        cboYear(1).AddItem X
    Next
    cboYear(0).ListIndex = Year(Date) - 1990
    cboYear(1).ListIndex = Year(Date) - 1990
    For X = 0 To 1
        cboMonth(X).AddItem "January"
        cboMonth(X).AddItem "February"
        cboMonth(X).AddItem "Mars"
        cboMonth(X).AddItem "April"
        cboMonth(X).AddItem "May"
        cboMonth(X).AddItem "June"
        cboMonth(X).AddItem "July"
        cboMonth(X).AddItem "August"
        cboMonth(X).AddItem "September"
        cboMonth(X).AddItem "October"
        cboMonth(X).AddItem "November"
        cboMonth(X).AddItem "December"
        cboMonth(X).ListIndex = Month(Date) - 1
        cboDay(X).ListIndex = Day(Date)
    Next
End Sub

Private Sub lstLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuDelete_Click()
Dim I As Integer
Dim vArray As Variant
    vArray = GetUser(One, TreeView1.SelectedItem.Text)
    If IsArray(vArray) Then
        If vArray(2, 0) <> "" Then
            CheckUser = False
            CheckUserName = TreeView1.SelectedItem.Text
            frmPassword.Show vbModal, frmLogShow
            If CheckUser = False Then Exit Sub
        End If
    End If

    Me.MousePointer = 11
    If lstLog.SelectedItem Is Nothing Then Exit Sub
    For I = 1 To lstLog.ListItems.Count
        If lstLog.ListItems(I).Selected = True Then
            DeleteLogItem TreeView1.SelectedItem.Text, Mid(lstLog.ListItems(I).Key, 3)
        End If
    Next I
    lstLog.ListItems.Clear
    vArray = GetLog(TreeView1.SelectedItem.Text)
    ShowLog vArray
    Me.MousePointer = 0
End Sub

Private Sub mnuPrint_Click()
    PrintLogData
End Sub

Public Sub ShowLog(ByVal vArray As Variant)
Dim X As Long
Dim lItem As ListItem
Dim dPrice As Double
Dim sPrice As String
Dim lTime As Long
Dim lTotTime As Long

    If IsArray(vArray) = False Then Exit Sub
    For X = 0 To UBound(vArray, 2)
        Set lItem = lstLog.ListItems.Add(, "id" & vArray(0, X), vArray(1, X) & "")
        lItem.SubItems(1) = vArray(2, X) & ""
        lTime = DateDiff("s", vArray(1, X), vArray(2, X))
        lTotTime = lTotTime + lTime
        lItem.SubItems(2) = Sec2Time(lTime)
        sPrice = vArray(4, X)
        sPrice = Cent2Dollar(Round(sPrice, 2))
        dPrice = dPrice + sPrice
        lItem.SubItems(3) = sPrice
        lItem.SubItems(4) = vArray(3, X)
    Next
    lstLog.ListItems.Add
    Set lItem = lstLog.ListItems.Add(, "TellInfo", "First Logon Date")
    lItem.SubItems(1) = "Last logon date"
    lItem.SubItems(2) = "Total Time"
    lItem.SubItems(3) = "Total Price"
    lItem.SubItems(4) = "-"
    Set lItem = lstLog.ListItems.Add(, "TotLog", vArray(1, 0))
    lItem.SubItems(1) = vArray(2, UBound(vArray, 2))
    lItem.SubItems(2) = Sec2Time(lTotTime)
    lItem.SubItems(3) = Cent2Dollar(Round(dPrice, 2))
    lItem.SubItems(4) = "-"
End Sub

Public Sub PrintLogData()
Dim I As Integer
    Printer.FontSize = 10
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    
    Printer.Print vbTab & "INet Counter v2.0 - Log Data"
    Printer.Print ""
    Printer.Print vbTab & "User: " & TreeView1.SelectedItem.Text
    Printer.Print ""
    Printer.Print vbTab & "Logon date        " & vbTab & "Logoff date       " & vbTab & vbTab & "Time" & vbTab & vbTab & "Price" & vbTab & vbTab & "Conection Name"
    Printer.Print ""
    For I = 1 To lstLog.ListItems.Count - 2
        Printer.Print vbTab & lstLog.ListItems(I).Text & vbTab & lstLog.ListItems(I).SubItems(1) & vbTab & lstLog.ListItems(I).SubItems(2) & vbTab & vbTab & lstLog.ListItems(I).SubItems(3) & vbTab & vbTab & lstLog.ListItems(I).SubItems(4)
    Next I
    Printer.Print vbTab & "First Logon date  " & vbTab & "Last Logoff date  " & vbTab & "Total Time" & vbTab & "Total Price" & vbTab & "-"
    I = lstLog.ListItems.Count
    Printer.Print vbTab & lstLog.ListItems(I).Text & vbTab & lstLog.ListItems(I).SubItems(1) & vbTab & lstLog.ListItems(I).SubItems(2) & vbTab & vbTab & lstLog.ListItems(I).SubItems(3) & vbTab & vbTab & lstLog.ListItems(I).SubItems(4)
    Printer.EndDoc
End Sub

Public Sub ShowAll()
Dim X As Long
Dim lItem As ListItem
Dim dPrice As Double
Dim sPrice As String
Dim lTime As Long
Dim lTotTime As Long
Dim sUser As String
Dim NodNr As Long
    NodNr = 2
    Do Until NodNr = TreeView1.Nodes("All").Index
        CheckUser = True
        sUser = TreeView1.Nodes(NodNr).Text
        vArray = GetUser(One, sUser)
        If vArray(2, 0) <> "" Then
            CheckUser = False
            CheckUserName = sUser
            frmPassword.Show vbModal, frmMain
        End If
        vArray = GetLog(sUser)
        If (IsArray(vArray)) And (CheckUser = True) Then
            For X = 0 To UBound(vArray, 2)
                Set lItem = lstLog.ListItems.Add(, "id" & sUser & vArray(0, X), vArray(1, X))
                lItem.SubItems(1) = vArray(2, X)
                lTime = DateDiff("s", vArray(1, X), vArray(2, X))
                lTotTime = lTotTime + lTime
                lItem.SubItems(2) = Sec2Time(lTime)
                sPrice = vArray(4, X)
                sPrice = Cent2Dollar(Round(sPrice, 2))
                dPrice = dPrice + sPrice
                lItem.SubItems(3) = sPrice
                lItem.SubItems(4) = sUser
            Next
        End If
        NodNr = NodNr + 1
    Loop
    If lstLog.ListItems.Count = 0 Then Exit Sub
    lstLog.SortKey = 1
    lstLog.Sorted = True
    lstLog.Sorted = False
    lstLog.ListItems.Add
    Set lItem = lstLog.ListItems.Add(, "TellInfo", "First Logon Date")
    lItem.SubItems(1) = "Last logon date"
    lItem.SubItems(2) = "Total Time"
    lItem.SubItems(3) = "Total Price"
    lItem.SubItems(4) = "-"
    Set lItem = lstLog.ListItems.Add(, "TotLog", lstLog.ListItems(1).Text)
    lItem.SubItems(1) = lstLog.ListItems(lstLog.ListItems.Count - 3).SubItems(1)
    lItem.SubItems(2) = Sec2Time(lTotTime)
    lItem.SubItems(3) = Cent2Dollar(Round(dPrice, 2))
    lItem.SubItems(4) = "-"
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
Dim vArray As Variant
    Me.MousePointer = 11
    lstLog.ListItems.Clear
    lstLog.ColumnHeaders(5).Text = "Conection Name"
    If Node.Text = "Users" Then
        cmdView.Enabled = False
    ElseIf Node.Text = "All users" Then
        cmdView.Enabled = False
        lstLog.ColumnHeaders(5).Text = "User name"
        ShowAll
    Else
        vArray = GetUser(One, Node.Text)
        If vArray(2, 0) <> "" Then
            CheckUser = False
            CheckUserName = Node.Text
            frmPassword.Show vbModal, frmMain
            If CheckUser = False Then
                Me.MousePointer = 0
                Exit Sub
            End If
        End If
        cmdView.Enabled = True
        vArray = GetLog(Node.Text)
        ShowLog vArray
    End If
    Me.MousePointer = 0
End Sub
