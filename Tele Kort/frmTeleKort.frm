VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmTeleKort 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Viktorex TeleKort"
   ClientHeight    =   3450
   ClientLeft      =   2445
   ClientTop       =   3405
   ClientWidth     =   4575
   Icon            =   "frmTeleKort.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4575
   Tag             =   "Tabell1"
   Begin VB.TextBox txtDir 
      Height          =   285
      Left            =   240
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2040
      Width           =   495
      Visible         =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Antal"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   5
      Top             =   1395
      Width           =   3375
   End
   Begin VB.CommandButton cmdBild 
      Caption         =   "Visa Bild"
      Height          =   300
      Left            =   1560
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Uppdatera"
      Height          =   300
      Left            =   2880
      TabIndex        =   8
      Tag             =   "&Refresh"
      Top             =   2625
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Tabort"
      Height          =   300
      Left            =   1800
      TabIndex        =   10
      Tag             =   "&Delete"
      Top             =   2625
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Lägg till"
      Height          =   300
      Left            =   720
      TabIndex        =   9
      Tag             =   "&Add"
      Top             =   2625
      Width           =   975
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "kort.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Tabell1"
      Top             =   3105
      Width           =   4575
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Motiv"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   5
      Left            =   1080
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1725
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Upplaga"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   4
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Markeringar"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   3
      Top             =   765
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Årtal"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   2
      Top             =   435
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artikelnr"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Antal"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1395
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Motiv:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Tag             =   "Motiv:"
      Top             =   1725
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Upplaga:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Tag             =   "Upplaga:"
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Markeringar:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Tag             =   "Markeringar:"
      Top             =   765
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Årtal:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Tag             =   "Årtal:"
      Top             =   435
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Artikelnr:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Tag             =   "Artikelnr:"
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu mnuArkiv 
      Caption         =   "Arkiv"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuAdd 
         Caption         =   "Lägg till"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuUppdatera 
         Caption         =   "&Uppdatera"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Tabort"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAvsluta 
         Caption         =   "&Avsluta"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuRedigera 
      Caption         =   "&Redigera"
      Begin VB.Menu mnuSök 
         Caption         =   "Sök..."
      End
      Begin VB.Menu mnuGåtill 
         Caption         =   "Gå till..."
      End
   End
   Begin VB.Menu mnuHjälp 
      Caption         =   "Hjälp"
      Begin VB.Menu mnuOm 
         Caption         =   "Om TeleKort"
      End
   End
End
Attribute VB_Name = "frmTeleKort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GåTill As Integer
Dim Record As Integer
Dim myDB As Database
Dim myRS As Recordset
Dim Dir As String
Dim StartRecord As Integer
Dim Sök As String
Dim Y As Integer

Private Sub cmdAdd_Click()
    Data1.Recordset.AddNew
    txtFields(1).SetFocus
End Sub


Private Sub cmdBild_Click()
    On Error GoTo errortrap
    If cmdBild.Caption = "Visa Bild" Then
        frmBild.Show
        cmdBild.Caption = "Stäng Bilden"
    Else
        Unload frmBild
        cmdBild.Caption = "Visa Bild"
    End If
errortrap:
Exit Sub
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo errortrap
    With Data1.Recordset
        .Delete
        .MoveNext
        If .EOF Then .MoveLast
    End With

errortrap: Exit Sub

End Sub


Private Sub cmdRefresh_Click()
    Data1.Refresh
End Sub

Private Sub Data1_Reposition()
    ' ska finnas i färdigt program
    'Set myDB = OpenDatabase(Dir + "\kort.mdb")
    Set myDB = OpenDatabase("G:\program\tele kort\kort.mdb")
    Set myRS = myDB.OpenRecordset("tabell1", dbOpenTable)
    Screen.MousePointer = vbDefault
    On Error Resume Next
    Data1.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1) & "/" & myRS.RecordCount
End Sub

Private Sub Form_Load()
    Dir = CurDir
    txtDir = CurDir
End Sub

Private Sub mnuAdd_Click()
    cmdAdd_Click
End Sub

Private Sub mnuAvsluta_Click()
    End
End Sub

Private Sub mnuDelete_Click()
    cmdDelete_Click
End Sub

Private Sub mnuGåtill_Click()
    ' ska finnas i färdigt program
    'Set myDB = OpenDatabase(Dir + "\kort.mdb")
    Set myDB = OpenDatabase("G:\program\tele kort\kort.mdb")
    Set myRS = myDB.OpenRecordset("tabell1", dbOpenTable)
    
    On Error GoTo errortrap
    GåTill = InputBox("Gå till record nr:", "Gå till")
    If GåTill < myRS.RecordCount + 1 Then
        Data1.Recordset.MoveFirst
        Data1.Recordset.Move (GåTill - 1)
    End If
errortrap:
Exit Sub
End Sub

Private Sub mnuOm_Click()
    frmAbout.Show
End Sub

Private Sub mnuSök_Click()
    ' ska finnas i färdigt program
    'Set myDB = OpenDatabase(Dir + "\kort.mdb")
    Set myDB = OpenDatabase("G:\program\tele kort\kort.mdb")
    Set myRS = myDB.OpenRecordset("tabell1", dbOpenTable)
    On Error GoTo errortrap
    Y = 0
    Sök = InputBox("Sök efter artikel nr:", "Sök")
    If Sök <> "" Then
        Data1.Recordset.MoveFirst
        Do Until Y = myRS.RecordCount
            If UCase(txtFields(1).Text) = UCase(Sök) Then
                Exit Sub
            End If
            Data1.Recordset.MoveNext
            Y = Y + 1
        Loop
    End If
    MsgBox "Det artikel nr som du söker efter finns inte i databasen.", vbInformation, "Sök"
    Exit Sub
errortrap:
Exit Sub
End Sub

Private Sub mnuUppdatera_Click()
    cmdRefresh_Click
End Sub

