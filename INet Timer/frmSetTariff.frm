VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSetTariff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tariff Editor"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4770
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAddTariff 
      Caption         =   "&Add Tariff"
      Height          =   315
      Left            =   3480
      TabIndex        =   21
      Top             =   600
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   3480
      TabIndex        =   20
      Top             =   960
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   3480
      TabIndex        =   19
      Top             =   1320
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   18
      Text            =   "Thursday:"
      Top             =   1695
      Width           =   1245
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "Wednesday:"
      Top             =   1335
      Width           =   1245
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   16
      Text            =   "Tuesday:"
      Top             =   975
      Width           =   1245
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   15
      Text            =   "Monday:"
      Top             =   615
      Width           =   1245
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   14
      Text            =   "Connection fee:"
      Top             =   3480
      Width           =   1245
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   13
      Text            =   "Holiday:"
      Top             =   3120
      Width           =   1245
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   12
      Text            =   "Sunday:"
      Top             =   2760
      Width           =   1245
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   11
      Text            =   "Saturday:"
      Top             =   2400
      Width           =   1245
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   10
      Text            =   "Friday:"
      Top             =   2055
      Width           =   1245
   End
   Begin VB.ComboBox cboUse 
      Height          =   315
      Index           =   0
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox cboUse 
      Height          =   315
      Index           =   6
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2745
      Width           =   1215
   End
   Begin VB.ComboBox cboUse 
      Height          =   315
      Index           =   5
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2385
      Width           =   1215
   End
   Begin VB.ComboBox cboUse 
      Height          =   315
      Index           =   4
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox cboUse 
      Height          =   315
      Index           =   3
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox cboUse 
      Height          =   315
      Index           =   2
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtCon 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ComboBox cboUse 
      Height          =   315
      Index           =   1
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox cboUse 
      Height          =   315
      Index           =   7
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3105
      Width           =   1215
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6800
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "frmSetTariff.frx":0000
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSetTariff.frx":031A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSetTariff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddTariff_Click()
    Unload Me
    frmTariff.Show , frmMain
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
Dim X As Integer
    For X = 0 To 7
        If cboUse(X).Text = "" Then
            MsgBox "You have to add a tariff for every day of the week!", vbInformation
            Exit Sub
        End If
    Next
    SaveWeek TabStrip1.SelectedItem.Caption
End Sub

Private Sub Form_Load()
Dim sRetVal As String
Dim X As Long
Dim Y As Integer
Dim vArray As Variant
    Me.Show
    X = TabStrip1.Tabs.Count
    If X > 0 Then
        TabStrip1.Tabs.Remove (1)
    End If
    'Ladda in alla Remote Access profiler
    sRetVal = oReg.EnumValue(HKEY_CURRENT_USER, "RemoteAccess\Addresses")
    X = 1
    Do Until X = 0
        X = InStr(X, sRetVal, Chr(0))
        If X <> 0 Then
            TabStrip1.Tabs.Add , , Mid(sRetVal, 1, X - 1), 1
            sRetVal = Mid(sRetVal, X + 1)
            X = 1
        End If
    Loop
    'Ladda in den första veckan
    vArray = GetPrice(All)
    If IsArray(vArray) = True Then
        For Y = 0 To 7
            For X = 0 To UBound(vArray, 2)
                cboUse(Y).AddItem vArray(1, X)
            Next
            cboUse(Y).ListIndex = 0
        Next
        GetWeek TabStrip1.Tabs(1).Caption
    End If
End Sub

Private Sub TabStrip1_Click()
Dim Read As String
Dim X As Integer
    Read = GetWeek(TabStrip1.SelectedItem.Caption)
    If Read = "No" Then
        For X = 0 To 7
            cboUse(X).ListIndex = -1
        Next
        txtCon.Text = "0"
    End If
End Sub
