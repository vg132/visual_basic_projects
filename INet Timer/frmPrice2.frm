VERSION 5.00
Begin VB.Form frmPrice 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Edit a Price"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   1110
      TabIndex        =   5
      Top             =   960
      Width           =   900
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   2115
      TabIndex        =   1
      Top             =   960
      Width           =   900
   End
   Begin VB.CommandButton cmdCansel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   900
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblPrice 
      AutoSize        =   -1  'True
      Caption         =   "Price/min (Cent):"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   525
      Width           =   1185
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Price Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   165
      Width           =   870
   End
End
Attribute VB_Name = "frmPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboName_Click()
'När användaren klickar på ett pris i combo så syns det i textbox
Dim vArray As Variant
    vArray = GetPrice(One, cboName.Text)
    txtPrice.Text = vArray(2, 0)
End Sub

Private Sub cmdCansel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    DeletePrice cboName.Text
    Form_Load
End Sub

Private Sub cmdSave_Click()
    SavePrice cboName.Text, txtPrice.Text
    Unload Me
End Sub

Private Sub Form_Load()
Dim vArray As Variant
Dim X As Integer
    cboName.Clear
    vArray = GetPrice(All)
    If IsArray(vArray) = False Then Exit Sub
    For X = 0 To UBound(vArray, 2)
        cboName.AddItem vArray(1, X)
    Next
End Sub
