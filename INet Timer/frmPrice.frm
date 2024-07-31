VERSION 5.00
Begin VB.Form frmTariff 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Tariff"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   315
      Left            =   3240
      TabIndex        =   36
      Top             =   2880
      Width           =   1035
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   1560
      TabIndex        =   35
      Top             =   2880
      Width           =   1035
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Index           =   3
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   2070
      Width           =   1215
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Index           =   0
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Index           =   4
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   2430
      Width           =   1215
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Index           =   2
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   1710
      Width           =   1215
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Index           =   1
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   1350
      Width           =   1215
   End
   Begin VB.ComboBox cboName 
      Height          =   315
      Left            =   1080
      TabIndex        =   28
      Top             =   120
      Width           =   4575
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   10
      TabIndex        =   27
      Top             =   990
      Width           =   1215
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Index           =   1
      Left            =   120
      MaxLength       =   10
      TabIndex        =   26
      Top             =   1350
      Width           =   1215
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Index           =   2
      Left            =   120
      MaxLength       =   10
      TabIndex        =   25
      Top             =   1710
      Width           =   1215
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Index           =   3
      Left            =   120
      MaxLength       =   10
      TabIndex        =   24
      Top             =   2070
      Width           =   1215
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Index           =   4
      Left            =   120
      MaxLength       =   10
      TabIndex        =   23
      Top             =   2430
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   120
      TabIndex        =   22
      Top             =   2880
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   4620
      TabIndex        =   21
      Top             =   2880
      Width           =   1035
   End
   Begin VB.OptionButton optNr 
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   19
      Top             =   585
      Width           =   375
   End
   Begin VB.OptionButton optNr 
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   18
      Top             =   585
      Width           =   375
   End
   Begin VB.OptionButton optNr 
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   17
      Top             =   585
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton optNr 
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   16
      Top             =   585
      Width           =   375
   End
   Begin VB.OptionButton optNr 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   15
      Top             =   585
      Width           =   375
   End
   Begin VB.ComboBox cboTime 
      Height          =   315
      Index           =   4
      Left            =   4440
      TabIndex        =   4
      Text            =   "cboTime"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox cboTime 
      Height          =   315
      Index           =   3
      Left            =   4440
      TabIndex        =   3
      Text            =   "cboTime"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox cboTime 
      Height          =   315
      Index           =   2
      Left            =   4440
      TabIndex        =   2
      Text            =   "cboTime"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox cboTime 
      Height          =   315
      Index           =   0
      Left            =   4440
      TabIndex        =   0
      Text            =   "cboTime"
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox cboTime 
      Height          =   315
      Index           =   1
      Left            =   4440
      TabIndex        =   1
      Text            =   "cboTime"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Tariff Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   180
      Width           =   870
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Nr of difrent tariffs/day:"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   600
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Use this price from"
      Height          =   195
      Left            =   1440
      TabIndex        =   14
      Top             =   2520
      Width           =   1305
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Use this price from"
      Height          =   195
      Left            =   1440
      TabIndex        =   13
      Top             =   1080
      Width           =   1305
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Use this price from"
      Height          =   195
      Left            =   1440
      TabIndex        =   12
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Use this price from"
      Height          =   195
      Left            =   1440
      TabIndex        =   11
      Top             =   2160
      Width           =   1305
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Use this price from"
      Height          =   195
      Left            =   1440
      TabIndex        =   10
      Top             =   1800
      Width           =   1305
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "to"
      Height          =   195
      Left            =   4170
      TabIndex        =   9
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "to"
      Height          =   195
      Left            =   4170
      TabIndex        =   8
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "to"
      Height          =   195
      Left            =   4170
      TabIndex        =   7
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "to"
      Height          =   195
      Left            =   4170
      TabIndex        =   6
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "to"
      Height          =   195
      Left            =   4170
      TabIndex        =   5
      Top             =   1080
      Width           =   135
   End
End
Attribute VB_Name = "frmTariff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurOpt As Integer

Private Sub cboName_Click()
Dim vArray As Variant
Dim X As Long
    If cboName.Text = "" Then Exit Sub
    vArray = GetTariff(cboName.Text)
    If IsArray(vArray) = True Then
        Reset
        X = UBound(vArray, 2)
        optNr(X).Value = True
        For X = 0 To UBound(vArray, 2)
            If vArray(2, X) = "23:59:59" Then vArray(2, X) = "24:00:00"
            If vArray(1, X) = "23:59:59" Then vArray(1, X) = "24:00:00"
            txtPrice(vArray(5, X)).Text = vArray(3, X)
            txtTime(vArray(5, X)).Text = vArray(1, X)
            cboTime(vArray(5, X)).Text = vArray(2, X)
        Next
    End If
End Sub

Private Sub cboTime_Change(Index As Integer)
    If (Index <> 4) And (CurOpt <> Index) Then txtTime(Index + 1).Text = cboTime(Index).Text
End Sub

Private Sub cboTime_Click(Index As Integer)
    If (Index <> 4) And (CurOpt <> Index) Then txtTime(Index + 1).Text = cboTime(Index).Text
End Sub

Private Sub cboTime_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Len(cboTime(Index).Text) > 5 Then
        cboTime(Index).Text = Mid(cboTime(Index).Text, 1, 5)
        cboTime(Index).SelStart = Len(cboTime(Index).Text)
        Exit Sub
    End If
    If KeyCode = 110 Then
        cboTime(Index).Text = Mid(cboTime(Index).Text, 1, Len(cboTime(Index).Text) - 1) & "."
        cboTime(Index).SelStart = Len(cboTime(Index).Text)
        Exit Sub
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
    frmSetTariff.Show , frmMain
End Sub

Private Sub cmdDelete_Click()
    DeleteTariff cboName.Text
    Form_Load
End Sub

Private Sub cmdNew_Click()
Dim X As Integer
    For X = 0 To 4
        cboTime(X).ListIndex = 0
        txtTime(X).Text = "00:00:00"
        txtPrice(X).Text = ""
    Next
    cboName.Text = ""
    optNr(2).Value = True
    cboTime(2).Text = "24:00:00"
End Sub

Private Sub cmdSave_Click()
Dim optCheck As Integer
    If cboName.Text = "" Then
        MsgBox "You have to enter a name for this tariff.", vbInformation
        cboName.SetFocus
        Exit Sub
    End If
    For optCheck = 0 To 4
        If optNr(optCheck).Value = True Then Exit For
    Next
    SaveTariff cboName.Text, optCheck
End Sub

Private Sub Form_Load()
Dim vArray As Variant
Dim X As Integer
Dim Y As Integer
    For X = 0 To 4
        cboTime(X).Clear
        cboTime(X).AddItem "00:00:00"
        cboTime(X).AddItem "01:00:00"
        cboTime(X).AddItem "02:00:00"
        cboTime(X).AddItem "03:00:00"
        cboTime(X).AddItem "04:00:00"
        cboTime(X).AddItem "05:00:00"
        cboTime(X).AddItem "06:00:00"
        cboTime(X).AddItem "07:00:00"
        cboTime(X).AddItem "08:00:00"
        cboTime(X).AddItem "09:00:00"
        cboTime(X).AddItem "10:00:00"
        cboTime(X).AddItem "11:00:00"
        cboTime(X).AddItem "12:00:00"
        cboTime(X).AddItem "13:00:00"
        cboTime(X).AddItem "14:00:00"
        cboTime(X).AddItem "15:00:00"
        cboTime(X).AddItem "16:00:00"
        cboTime(X).AddItem "17:00:00"
        cboTime(X).AddItem "18:00:00"
        cboTime(X).AddItem "19:00:00"
        cboTime(X).AddItem "20:00:00"
        cboTime(X).AddItem "21:00:00"
        cboTime(X).AddItem "22:00:00"
        cboTime(X).AddItem "23:00:00"
        cboTime(X).AddItem "24:00:00"
        cboTime(X).ListIndex = 0
        txtTime(X).Text = "00:00:00"
    Next
    cboName.Clear
    optNr_Click (2)
    vArray = GetPrice(All)
    If IsArray(vArray) = True Then
        For X = 0 To UBound(vArray, 2)
            cboName.AddItem vArray(1, X)
        Next
    End If
    For X = 0 To 4
        txtPrice(X) = ""
    Next
End Sub

Private Sub optNr_Click(Index As Integer)
Dim X As Integer
    CurOpt = Index
    For X = 0 To Index
        cboTime(X).Enabled = True
        txtTime(X).Enabled = True
        txtPrice(X).Enabled = True
        cboTime(X).Locked = False
    Next
    cboTime(X - 1).Text = "24:00:00"
    cboTime(X - 1).Locked = True
    For X = Index + 1 To 4
        cboTime(X).Enabled = False
        txtTime(X).Enabled = False
        txtPrice(X).Enabled = False
    Next
End Sub

Private Sub Reset()
'*************************************
'Function Name: Reset
'Use: Clear all controls
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-21
'*************************************
Dim X As Integer
    For X = 0 To 4
        cboTime(X).ListIndex = 0
        txtTime(X).Text = "00:00:00"
    Next
    For X = 0 To 4
        txtPrice(X).Text = ""
    Next
End Sub
