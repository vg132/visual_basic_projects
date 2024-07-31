VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Category"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Title/Name"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search Options"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4575
      Begin VB.CheckBox chkCase 
         Caption         =   "Match Ca&se"
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkGame 
         Caption         =   "&Game"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkMusic 
         Caption         =   "M&usic"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkFilm 
         Caption         =   "Fil&m"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Find"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   315
         Left            =   3360
         TabIndex        =   2
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Found"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   4575
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFind_Click()
Dim NrOf As Long
Dim TempNr As Long
Dim Found As Boolean
    Found = False
    ListView1.ListItems.Clear
    If chkGame.Value = 1 Then
        vArray = oDB.GetAllGames
        NrOf = oDB.GetNr
        For X = 0 To NrOf - 1
            For Y = 0 To 6
                For TempNr = 1 To Len("" & vArray(Y, X)) - Len(txtFind.Text)
                    If chkCase.Value = 1 Then
                        If Mid(vArray(Y, X), TempNr, Len(txtFind.Text)) = txtFind.Text Then
                            Set itemX = ListView1.ListItems.Add(, "K" & vArray(0, X), vArray(0, X))
                            itemX.SubItems(1) = "Game"
                            itemX.SubItems(2) = vArray(1, X)
                            Found = True
                            Exit For
                        End If
                    ElseIf chkCase.Value = 0 Then
                        If Mid(UCase(vArray(Y, X)), TempNr, Len(txtFind.Text)) = UCase(txtFind.Text) Then
                            Set itemX = ListView1.ListItems.Add(, "K" & vArray(0, X), vArray(0, X))
                            itemX.SubItems(1) = "Game"
                            itemX.SubItems(2) = vArray(1, X)
                            Found = True
                            Exit For
                        End If
                    End If
                Next
                If Found = True Then Exit For
            Next
            Found = False
        Next
    End If
    
    If chkMusic.Value = 1 Then
        vArray = oDB.GetAllMusic
        NrOf = oDB.GetNr
        TempNr = 0
        If oDB.GetNr > 0 Then
            For X = 0 To NrOf - 1
                For Y = 0 To 3
                    For TempNr = 1 To Len("" & vArray(Y, X)) - Len(txtFind.Text)
                        If chkCase.Value = 1 Then
                            If Mid(vArray(Y, X), TempNr, Len(txtFind.Text)) = txtFind.Text Then
                                Set itemX = ListView1.ListItems.Add(, "K" & vArray(0, X), vArray(0, X))
                                itemX.SubItems(1) = "Music"
                                itemX.SubItems(2) = vArray(1, X)
                                Found = True
                                Exit For
                            End If
                        ElseIf chkCase.Value = 0 Then
                            If Mid(UCase(vArray(Y, X)), TempNr, Len(txtFind.Text)) = UCase(txtFind.Text) Then
                                Set itemX = ListView1.ListItems.Add(, "K" & vArray(0, X), vArray(0, X))
                                itemX.SubItems(1) = "Music"
                                itemX.SubItems(2) = vArray(1, X)
                                Found = True
                                Exit For
                            End If
                        End If
                    Next
                    If Found = True Then Exit For
                Next
                Found = False
            Next
        End If
    End If
    
    
    
    
    
    
    If chkMusic.Value = 1 Then
        vArray = oDB.GetAllFilm
        NrOf = oDB.GetNr
        TempNr = 0
        If oDB.GetNr > 0 Then
            For X = 0 To NrOf - 1
                For Y = 0 To 3
                    For TempNr = 1 To Len("" & vArray(Y, X)) - Len(txtFind.Text)
                        If chkCase.Value = 1 Then
                            If Mid(vArray(Y, X), TempNr, Len(txtFind.Text)) = txtFind.Text Then
                                Set itemX = ListView1.ListItems.Add(, "K" & vArray(0, X), vArray(0, X))
                                itemX.SubItems(1) = "Film"
                                itemX.SubItems(2) = vArray(1, X)
                                Found = True
                                Exit For
                            End If
                        ElseIf chkCase.Value = 0 Then
                            If Mid(UCase(vArray(Y, X)), TempNr, Len(txtFind.Text)) = UCase(txtFind.Text) Then
                                Set itemX = ListView1.ListItems.Add(, "K" & vArray(0, X), vArray(0, X))
                                itemX.SubItems(1) = "Film"
                                itemX.SubItems(2) = vArray(1, X)
                                Found = True
                                Exit For
                            End If
                        End If
                    Next
                    If Found = True Then Exit For
                Next
                Found = False
            Next
        End If
    End If

    
    
    
    
    
    
    
    
    If ListView1.ListItems.Count > 0 Then
        frmFind.Height = 4980
    Else
        frmFind.Height = 1875
        MsgBox "Search text is not found.", vbInformation + vbOKOnly, "Find"
    End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.Sorted = True
    If ColumnHeader.Index = 1 Then
        If ListView1.SortOrder = lvwAscending Then
            ListView1.SortKey = 0
            ListView1.SortOrder = lvwDescending
        Else
            ListView1.SortKey = 0
            ListView1.SortOrder = lvwAscending
        End If
    ElseIf ColumnHeader.Index = 2 Then
        If ListView1.SortOrder = lvwAscending Then
            ListView1.SortKey = 1
            ListView1.SortOrder = lvwDescending
        Else
            ListView1.SortKey = 1
            ListView1.SortOrder = lvwAscending
        End If
    ElseIf ColumnHeader.Index = 3 Then
        If ListView1.SortOrder = lvwAscending Then
            ListView1.SortKey = 2
            ListView1.SortOrder = lvwDescending
        Else
            ListView1.SortKey = 2
            ListView1.SortOrder = lvwAscending
        End If
    End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    frmMain.TreeView1.Nodes(Item.SubItems(1) & "-" & Item.Text).Selected = True
    frmMain.TreeView1.SelectedItem.EnsureVisible
    frmMain.NodeClick
End Sub
