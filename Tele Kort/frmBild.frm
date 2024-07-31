VERSION 5.00
Begin VB.Form frmBild 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bilder på Telefonkortet"
   ClientHeight    =   3795
   ClientLeft      =   7080
   ClientTop       =   3075
   ClientWidth     =   4785
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   Begin VB.Image imgBak 
      Height          =   3810
      Left            =   2400
      Top             =   0
      Width           =   2430
   End
   Begin VB.Image imgFram 
      Height          =   3810
      Left            =   0
      Top             =   0
      Width           =   2460
   End
End
Attribute VB_Name = "frmBild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Fram As String
Dim Bak As String
Dim Dir As String
Dim ArtikelNr As String
Dim FileName As String
Dim CharLoop As Integer

Private Sub Form_Load()
    ArtikelNr = ""
    FileName = ""
    ArtikelNr = frmTeleKort.txtFields(1).Text
    For CharLoop = 1 To Len(ArtikelNr)
        If IsNumeric(Mid(ArtikelNr, CharLoop, 1)) = True Then
            FileName = FileName + Mid(ArtikelNr, CharLoop, 1)
        End If
    Next CharLoop
    'Detta ska användas i färdig version
    'Dir = frmTeleKort.txtDir.Text
    'Bak = Dir + "\bilder\" + FileName + ".gif"
    'Fram = Dir + "\bilder\" + FileName + ".jpg"
    
    'Detta används vid utväckling
    Bak = "g:\program\tele kort\bilder\" + FileName + ".gif"
    Fram = "g:\program\tele kort\bilder\" + FileName + ".jpg"
    
    On Error GoTo errortrap
    Set imgBak = LoadPicture(Bak)
    Set imgFram = LoadPicture(Fram)
    Storlek
    Exit Sub
errortrap:
MsgBox "Kunde inte hitta en bild på detta kort.", vbOKOnly, "Viktorex TeleKort"
Unload frmBild
End Sub

Public Sub Storlek()
    frmBild.Width = imgBak.Width + imgFram.Width + 100
    If imgBak.Height > imgFram.Height Then
        frmBild.Height = imgBak.Height + 300
    Else
        frmBild.Height = imgFram.Height + 300
    End If
    imgBak.Left = imgFram.Width + imgFram.Left + 10
End Sub
