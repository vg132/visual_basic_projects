VERSION 5.00
Begin VB.Form frmMeny 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3600
   ClientLeft      =   390
   ClientTop       =   3180
   ClientWidth     =   1950
   ControlBox      =   0   'False
   Icon            =   "frmMeny.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   1950
   Begin VB.Image imgMOpt 
      Height          =   615
      Left            =   0
      ToolTipText     =   "Options"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Image imgOpt2 
      Height          =   600
      Left            =   0
      Picture         =   "frmMeny.frx":030A
      Top             =   6720
      Width           =   1950
      Visible         =   0   'False
   End
   Begin VB.Image imgEnsam 
      Height          =   600
      Left            =   4080
      Picture         =   "frmMeny.frx":0826
      Top             =   4560
      Width           =   1950
      Visible         =   0   'False
   End
   Begin VB.Image imgFlera2 
      Height          =   600
      Left            =   4080
      Picture         =   "frmMeny.frx":0C2A
      Top             =   5280
      Width           =   1950
      Visible         =   0   'False
   End
   Begin VB.Image imgOpt 
      Height          =   600
      Left            =   4080
      Picture         =   "frmMeny.frx":10AA
      Top             =   6000
      Width           =   1950
      Visible         =   0   'False
   End
   Begin VB.Image imgAbout2 
      Height          =   600
      Left            =   2040
      Picture         =   "frmMeny.frx":1567
      Top             =   4560
      Width           =   1950
      Visible         =   0   'False
   End
   Begin VB.Image imgFlera 
      Height          =   600
      Left            =   2040
      Picture         =   "frmMeny.frx":1A4A
      Top             =   5280
      Width           =   1950
      Visible         =   0   'False
   End
   Begin VB.Image imgQuit2 
      Height          =   600
      Left            =   2040
      Picture         =   "frmMeny.frx":1E8F
      Top             =   6000
      Width           =   1950
      Visible         =   0   'False
   End
   Begin VB.Image imgQuit 
      Height          =   600
      Left            =   0
      Picture         =   "frmMeny.frx":231A
      Top             =   6000
      Width           =   1950
      Visible         =   0   'False
   End
   Begin VB.Image imgEnsam2 
      Height          =   600
      Left            =   0
      Picture         =   "frmMeny.frx":2753
      Top             =   5280
      Width           =   1950
      Visible         =   0   'False
   End
   Begin VB.Image imgAbout 
      Height          =   600
      Left            =   0
      Picture         =   "frmMeny.frx":2BA3
      Top             =   4560
      Width           =   1950
      Visible         =   0   'False
   End
   Begin VB.Image imgMQuit 
      Height          =   600
      Left            =   0
      ToolTipText     =   "Exit Finder"
      Top             =   3000
      Width           =   1950
   End
   Begin VB.Image imgMAbout 
      Height          =   600
      Left            =   0
      ToolTipText     =   "About Finder"
      Top             =   2400
      Width           =   1950
   End
   Begin VB.Image imgMFind 
      Height          =   600
      Left            =   0
      ToolTipText     =   "Search The Database"
      Top             =   1200
      Width           =   1950
   End
   Begin VB.Image imgMEnsama 
      Height          =   600
      Left            =   0
      ToolTipText     =   "Ordinary CD's"
      Top             =   0
      Width           =   1950
   End
   Begin VB.Image imgMFlera 
      Height          =   600
      Left            =   0
      ToolTipText     =   "Collection CD's"
      Top             =   600
      Width           =   1950
   End
   Begin VB.Image imgFind2 
      Height          =   600
      Left            =   4080
      Picture         =   "frmMeny.frx":301F
      Top             =   6720
      Width           =   1950
      Visible         =   0   'False
   End
   Begin VB.Image imgFind 
      Height          =   600
      Left            =   2040
      Picture         =   "frmMeny.frx":3512
      Top             =   6720
      Width           =   1950
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMeny"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Denna del av programet är såvitt jag vet klar!
Dim Start As Boolean 'Har programet startats förut? True=JA
Dim Name3 As String

Private Sub Form_Load()
    On Error GoTo errorstop
    'Ge start ett värde, om programet inte har startats så går det vidare till errorstop
    Start = GetSetting("Finder", "Config", "Start", "Dir")
    If Start = True Then
        imgMFlera.Picture = imgFlera2.Picture
        imgMEnsama.Picture = imgEnsam2.Picture
        imgMQuit.Picture = imgQuit2.Picture
        imgMAbout.Picture = imgAbout2.Picture
        imgMFind.Picture = imgFind2.Picture
        imgMOpt.Picture = imgOpt2.Picture
        Exit Sub
    End If
errorstop:
imgMFlera.Picture = imgFlera2.Picture
imgMEnsama.Picture = imgEnsam2.Picture
imgMQuit.Picture = imgQuit2.Picture
imgMAbout.Picture = imgAbout2.Picture
imgMFind.Picture = imgFind2.Picture
imgMOpt.Picture = imgOpt2.Picture
'Starta Skärmvals meny
frmOptions.Show
End Sub

Private Sub imgMAbout_Click()
    frmAbout.Show
End Sub

Private Sub imgMAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMFlera.Picture = imgFlera2.Picture
    imgMEnsama.Picture = imgEnsam2.Picture
    imgMQuit.Picture = imgQuit2.Picture
    imgMAbout.Picture = imgAbout.Picture
    imgMFind.Picture = imgFind2.Picture
    imgMOpt.Picture = imgOpt2.Picture
End Sub

Private Sub imgMEnsama_Click()
    frmsammaartist.Show
End Sub

Private Sub imgMEnsama_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMFlera.Picture = imgFlera2.Picture
    imgMEnsama.Picture = imgEnsam.Picture
    imgMQuit.Picture = imgQuit2.Picture
    imgMAbout.Picture = imgAbout2.Picture
    imgMFind.Picture = imgFind2.Picture
    imgMOpt.Picture = imgOpt2.Picture
End Sub

Private Sub imgMFind_Click()
    frmFinde.Show
End Sub

Private Sub imgMFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMFlera.Picture = imgFlera2.Picture
    imgMEnsama.Picture = imgEnsam2.Picture
    imgMQuit.Picture = imgQuit2.Picture
    imgMAbout.Picture = imgAbout2.Picture
    imgMFind.Picture = imgFind.Picture
    imgMOpt.Picture = imgOpt2.Picture
End Sub

Private Sub imgMFlera_Click()
    frmsammling.Show
End Sub

Private Sub imgMFlera_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMFlera.Picture = imgFlera.Picture
    imgMEnsama.Picture = imgEnsam2.Picture
    imgMQuit.Picture = imgQuit2.Picture
    imgMAbout.Picture = imgAbout2.Picture
    imgMFind.Picture = imgFind2.Picture
    imgMOpt.Picture = imgOpt2.Picture
End Sub

Private Sub imgMOpt_Click()
    frmOptions.Show
End Sub

Private Sub imgMOpt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMFlera.Picture = imgFlera2.Picture
    imgMEnsama.Picture = imgEnsam2.Picture
    imgMQuit.Picture = imgQuit2.Picture
    imgMAbout.Picture = imgAbout2.Picture
    imgMFind.Picture = imgFind2.Picture
    imgMOpt.Picture = imgOpt.Picture
End Sub

Private Sub imgMQuit_Click()
    End
End Sub

Private Sub imgMQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMFlera.Picture = imgFlera2.Picture
    imgMEnsama.Picture = imgEnsam2.Picture
    imgMQuit.Picture = imgQuit.Picture
    imgMAbout.Picture = imgAbout2.Picture
    imgMFind.Picture = imgFind2.Picture
    imgMOpt.Picture = imgOpt2.Picture
End Sub

