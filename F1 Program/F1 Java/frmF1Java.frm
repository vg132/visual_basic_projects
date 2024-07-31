VERSION 5.00
Begin VB.Form frmF1Java 
   BackColor       =   &H0060C8FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formula One JavaScript"
   ClientHeight    =   4830
   ClientLeft      =   5415
   ClientTop       =   3675
   ClientWidth     =   4635
   Icon            =   "frmF1Java.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   4635
   Tag             =   "Tabell1"
   Begin VB.TextBox txtStart 
      DataField       =   "Start Pos"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1200
      TabIndex        =   19
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdDNF 
      Caption         =   "DNF"
      Height          =   375
      Left            =   3240
      TabIndex        =   18
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtJava 
      Height          =   285
      Left            =   1200
      TabIndex        =   17
      Top             =   3960
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Write Java"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   975
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Tabell1"
      Top             =   4485
      Width           =   4635
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Next Race"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   8
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   8
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Coments"
      DataSource      =   "Data1"
      Height          =   1125
      Index           =   7
      Left            =   1200
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "WC Pos"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   6
      Left            =   1200
      TabIndex        =   6
      Top             =   2085
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Points"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      Top             =   1755
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Pos"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   4
      Left            =   1200
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Race"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   3
      Top             =   765
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Team"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   2
      Top             =   435
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Driver"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0060C8FF&
      Caption         =   "Start Pos"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0066CCFF&
      Caption         =   "Next Race:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   15
      Tag             =   "Next Race:"
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0066CCFF&
      Caption         =   "Coments:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Tag             =   "Coments:"
      Top             =   2415
      Width           =   855
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0066CCFF&
      Caption         =   "WC Pos:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Tag             =   "WC Pos:"
      Top             =   2100
      Width           =   855
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0066CCFF&
      Caption         =   "Total Points:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Tag             =   "Points:"
      Top             =   1785
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0066CCFF&
      Caption         =   "Pos:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Tag             =   "Pos:"
      Top             =   1455
      Width           =   855
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0066CCFF&
      Caption         =   "Race:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Tag             =   "Race:"
      Top             =   780
      Width           =   855
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0066CCFF&
      Caption         =   "Team:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Tag             =   "Team:"
      Top             =   465
      Width           =   855
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0066CCFF&
      Caption         =   "Driver:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Tag             =   "Driver:"
      Top             =   135
      Width           =   855
   End
End
Attribute VB_Name = "frmF1Java"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    Data1.Recordset.AddNew
End Sub


Private Sub cmdDelete_Click()
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    With Data1.Recordset
        .Delete
        .MoveNext
        If .EOF Then .MoveLast
    End With
End Sub


Private Sub cmdRefresh_Click()
    'this is really only needed for multi user apps
    Data1.Refresh
End Sub


Private Sub cmdUpdate_Click()
    Data1.UpdateRecord
    Data1.Recordset.Bookmark = Data1.Recordset.LastModified
End Sub


Private Sub cmdGrid_Click()
    On Error GoTo cmdGrid_ClickErr
    
    Set f.Data1.Recordset = Data1.Recordset
    f.Caption = Me.Caption & " Grid"
    f.Show


    Exit Sub
cmdGrid_ClickErr:
End Sub


Private Sub cmdDNF_Click()
    txtFields(4).Text = "DNF"
End Sub

Private Sub Command1_Click()
    txtJava = "<strong>Driver: </strong>" + txtFields(1).Text + "<br><strong>Team: </strong>" + txtFields(2).Text + "<br><strong>Race: </strong>" + txtFields(3).Text + "<br><strong>Starting Pos:</Strong></br>" + txtStart.Text + "<br><strong>Pos:</strong> " + txtFields(4).Text + "<br><strong>Points:</strong> " + txtFields(5).Text + "<br><strong>WC Pos:</strong> " + txtFields(6).Text + "<br><strong>Coments: </strong>" + txtFields(7).Text + "<br><strong>Next Race:</strong> " + txtFields(8).Text
End Sub

Private Sub Data1_Error(DataErr As Integer, Response As Integer)
    'This is where you would put error handling code
    'If you want to ignore errors, comment out the next line
    'If you want to trap them, add code here to handle them
    MsgBox "Data error event hit err:" & Error$(DataErr)
    Response = 0  'throw away the error
End Sub


Private Sub Data1_Reposition()
    Screen.MousePointer = vbDefault
    On Error Resume Next
    'This will display the current record position
    'for dynasets and snapshots
    Data1.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1)
    'for the table object you must set the index property when
    'the recordset gets created and use the following line
    'Data1.Caption = "Record: " & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub


Private Sub Data1_Validate(Action As Integer, Save As Integer)
    'This is where you put validation code
    'This event gets called when the following actions occur
    Select Case Action
        Case vbDataActionMoveFirst
        Case vbDataActionMovePrevious
        Case vbDataActionMoveNext
        Case vbDataActionMoveLast
        Case vbDataActionAddNew
        Case vbDataActionUpdate
        Case vbDataActionDelete
        Case vbDataActionFind
        Case vbDataActionBookmark
        Case vbDataActionClose
            Screen.MousePointer = vbDefault
    End Select
    Screen.MousePointer = vbHourglass
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

