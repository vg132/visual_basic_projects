VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmTabell1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Date Hacker - Version 2.0"
   ClientHeight    =   1440
   ClientLeft      =   4305
   ClientTop       =   3210
   ClientWidth     =   5550
   ControlBox      =   0   'False
   Icon            =   "Hacker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   5550
   Tag             =   "Tabell1"
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "G:\program\Date Hacker\data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Tabell1"
      Top             =   1095
      Width           =   5550
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Date"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   5
      Top             =   765
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Location"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   3
      Top             =   435
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Name"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Date (e.g: 12-28-1997)"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Tag             =   "Date:"
      Top             =   780
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Location:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Tag             =   "Location:"
      Top             =   465
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Name:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Tag             =   "Name:"
      Top             =   135
      Width           =   1815
   End
   Begin VB.Menu mnuH 
      Caption         =   "H"
      Begin VB.Menu mnuRun 
         Caption         =   "Run"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuLocation 
         Caption         =   "Location"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReload 
         Caption         =   "Reload"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmTabell1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RealDate As Date
Dim NewDate As Date
Dim Starta As String
Dim RetVal

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
    If (txtFields(2) <> "") And (txtFields(3).Text <> "") Then
        NewDate = txtFields(3)
        Date = NewDate
        Starta = txtFields(2)
        RetVal = Shell(Starta, 1)
    End If
End Sub

Private Sub cmdUpdate_Click()
    CommonDialog1.ShowOpen
    txtFields(2).Text = CommonDialog1.filename
End Sub


Private Sub cmdGrid_Click()
    Data1.Refresh
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

Private Sub Form_Load()
    RealDate = Date
    mnuH.Visible = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuH
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub lblLabels_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuH
    End If
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Date Hacker v 2.0 (C) Viktor Gars 1998. This program was made by Viktor Gars from Sweden 1998. This version of Date Hacker is made in Microsoft Visual Basic 5, I hope to be able to use Microsoft Visual C++ for version 3.0 of Date Hacker (if ther will ever be a third version).", vbInformation, "About Date Hacker"
End Sub

Private Sub mnuAdd_Click()
    cmdAdd_Click
End Sub

Private Sub mnuDelete_Click()
    cmdDelete_Click
End Sub

Private Sub mnuExit_Click()
    Date = RealDate
    End
End Sub

Private Sub mnuLocation_Click()
    cmdUpdate_Click
End Sub

Private Sub mnuReload_Click()
    cmdGrid_Click
End Sub

Private Sub mnuRun_Click()
    cmdRefresh_Click
End Sub

Private Sub txtFields_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuH
    End If
End Sub
