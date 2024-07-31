VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Print Options"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   2580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Print Options"
      Height          =   2175
      Left            =   100
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.OptionButton optAllFilm 
         Caption         =   "Print list of all films"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optAllMusic 
         Caption         =   "Print list of all records"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1000
         Width           =   1935
      End
      Begin VB.OptionButton optAllGames 
         Caption         =   "Print list of all games"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1280
         Width           =   2175
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   1680
         Width           =   1035
      End
      Begin VB.OptionButton optCurrent 
         Caption         =   "optCurrent"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
    Me.MousePointer = 11
    If optCurrent.Value = True Then
        If frmPrint.Tag = "Game" Then
            Printer.FontSize = 10
            Printer.Print ""
            Printer.Print ""
            Printer.Print "      " & "Game name:     " & Trim(frmMain.txtName.Text)
            Printer.Print "      " & "Game maker:    " & Trim(frmMain.txtMaker.Text)
            Printer.Print "      " & "Game format:    " & Trim(frmMain.cboFormat.Text)
            Printer.Print "      " & "Type of game:   " & Trim(frmMain.cboType.Text)
            Printer.Print "      " & "Grade:            " & Trim(frmMain.cboGrade.Text)
            Printer.Print "      " & "Media:            " & Trim(frmMain.cboMedia.Text)
            Printer.EndDoc
        ElseIf frmPrint.Tag = "Music" Then
            Printer.FontSize = 10
            Printer.Print ""
            Printer.Print ""
            Printer.Print "      " & "Record Title:     " & frmMain.txtRecName.Text
            Printer.Print "      " & "Artist:              " & frmMain.txtArtist.Text
            Printer.Print ""
            Printer.Print "      " & "Tracks"
            For X = 0 To frmMain.lstTracks.ListCount - 1
                Printer.Print "      " & frmMain.lstTracks.List(X)
            Next
            Printer.EndDoc
        ElseIf frmPrint.Tag = "Film" Then
            
        ElseIf frmPrint.Tag = "About" Then
            
        End If
    ElseIf optAllGames.Value = True Then
        vArray = oDB.GetGame
        X = oDB.GetNr
        If X < 1 Then
            Me.MousePointer = 0
            MsgBox "You don't have any games in the database.", vbInformation, "Print"
            Exit Sub
        End If
        Printer.FontSize = 10
        Printer.Print ""
        Printer.Print ""
        For Y = 0 To X - 1
            Printer.Print "      " & vArray(0, Y)
        Next
        Printer.EndDoc
    ElseIf optAllMusic.Value = True Then
        vArray = oDB.GetMusicName
        X = oDB.GetNr
        If X < 1 Then
            Me.MousePointer = 0
            MsgBox "You don't have any records in the database.", vbInformation, "Print"
            Exit Sub
        End If
        Printer.FontSize = 10
        Printer.Print ""
        Printer.Print ""
        For Y = 0 To X - 1
            Printer.Print "      " & vArray(0, Y)
        Next
        Printer.EndDoc
    End If
    Me.MousePointer = 0
    Unload frmPrint
End Sub

Private Sub Form_Activate()
    If frmPrint.Tag = "Game" Then
        optCurrent.Caption = "Print selected Game data"
    ElseIf frmPrint.Tag = "Music" Then
        optCurrent.Caption = "Print selected Music data"
    ElseIf frmPrint.Tag = "Film" Then
        optCurrent.Caption = "Print selected Film data"
    ElseIf frmPrint.Tag = "About" Then
        optCurrent.Caption = "Print the about screen"
    End If
End Sub
