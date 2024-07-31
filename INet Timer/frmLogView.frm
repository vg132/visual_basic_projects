VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogView 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INet Counter Log View"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   315
      Left            =   6660
      TabIndex        =   1
      Top             =   3360
      Width           =   1035
   End
   Begin MSComctlLib.ListView lstLog 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5741
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Log On Date/Time"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Logoff Date/Time"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Conection Name"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmLogView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim vArray As Variant
Dim X As Long
Dim ItemX As ListItem
Dim vArray2 As Variant
Dim TotPrice As Double
Dim StartDate As String
    TotPrice = 0
    vArray = GetLog
    If Not IsArray(vArray) Then Exit Sub
    StartDate = vArray(2, 0)
    For X = 0 To UBound(vArray, 2)
        Set ItemX = lstLog.ListItems.Add(, , vArray(1, X))
        With ItemX
            .SubItems(1) = vArray(2, X)
            .SubItems(2) = vArray(3, X)
            .SubItems(3) = Cent2Dollar(vArray(4, X))
            .SubItems(4) = "" & vArray(5, X)
        End With
        TotPrice = TotPrice + vArray(4, X)
    Next
    lstLog.ListItems.Add
    Set ItemX = lstLog.ListItems.Add(, , "Total")
    With ItemX
        .SubItems(1) = StartDate
        .SubItems(2) = vArray(3, X - 1)
        .SubItems(3) = Cent2Dollar(TotPrice)
    End With
End Sub
