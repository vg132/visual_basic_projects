VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find..."
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4515
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lstFound 
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   5424
      EndProperty
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Search"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   615
      Left            =   3720
      Picture         =   "frmFind.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Type in the &word(s) to search for:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFind_Click()
Dim vArray As Variant
Dim X As Long
    Me.MousePointer = 11
    lstFound.ListItems.Clear
    vArray = Find(txtFind.Text, 18)
    If IsArray(vArray) Then
        For X = 0 To UBound(vArray, 2)
            lstFound.ListItems.Add , "id" & vArray(0, X), vArray(1, X)
        Next
    End If
    lstFound.SortKey = 0
    lstFound.Sorted = True
    Me.MousePointer = 0
End Sub

Private Sub lstFound_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim sData As String
    sData = ""
    With frmCode
        .mnuDeleteItem.Enabled = True
        .mnuDeleteItem.Caption = "Delete Tip..."
        .mnuAddItem.Caption = "Add Tip..."
        sData = GetTip(Mid(lstFound.SelectedItem.Key, 3), 18)
        On Error Resume Next
        Kill App.Path & "\temphtml.htm"
        FileNum = FreeFile
        Open App.Path & "\TempHtml.htm" For Binary As FileNum
        Put #FileNum, 1, sData
        Close FileNum
        .WebBrowser1.Navigate2 App.Path & "\temphtml.htm"
    
        .TreeView1.Nodes("Tips-" & Mid(lstFound.SelectedItem.Key, 3)).Selected = True
        .TreeView1.SelectedItem.EnsureVisible
    End With
End Sub

Private Sub txtFind_Change()
    If txtFind.Text = "" Then
        cmdFind.Enabled = False
    Else
        cmdFind.Enabled = True
    End If
End Sub
