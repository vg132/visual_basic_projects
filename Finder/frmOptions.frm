VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Screen Options"
   ClientHeight    =   2055
   ClientLeft      =   2970
   ClientTop       =   3030
   ClientWidth     =   1710
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   1710
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Screen Size"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK!"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton opt3 
         Caption         =   "1024*768 (or more)"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton opt2 
         Caption         =   "800*600"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton opt1 
         Caption         =   "640*480"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Denna del av programet värkar vara klar
Dim Screen As Integer 'Screen size, 1=640*480 2=800*600 3=1024*768
Dim Start As Boolean 'Har programet startats förut?? True = JA

Private Sub cmdOk_Click()
    Start = True
    If opt1 = True Then Screen = 1
    If opt2 = True Then Screen = 2
    If opt3 = True Then Screen = 3
    'Spara skärmens storlek i registret
    SaveSetting "Finder", "Config", "Screen Size", Screen
    'medela programet att det inte är första gången som det startas
    SaveSetting "Finder", "Config", "Start", Start
    'Avsluta
    Unload frmOptions
End Sub

Private Sub Form_Load()
    On Error GoTo errortrap
    Screen = GetSetting("Finder", "Config", "Screen Size", "Dir")
    If Screen = 1 Then opt1 = True
    If Screen = 2 Then opt2 = True
    If Screen = 3 Then opt3 = True
errortrap:
Exit Sub
End Sub
