VERSION 5.00
Begin VB.Form frmNamn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Namn"
   ClientHeight    =   2040
   ClientLeft      =   4890
   ClientTop       =   3075
   ClientWidth     =   3690
   Icon            =   "Namn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3690
   Begin VB.OptionButton opt5 
      Caption         =   "5 Spelare"
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton opt4 
      Caption         =   "4 Spelare"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton opt3 
      Caption         =   "3 Spelare"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.OptionButton opt2 
      Caption         =   "2 Spelare"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   360
      Width           =   1215
   End
   Begin VB.OptionButton opt1 
      Caption         =   "1 Spelare"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtFem 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      MaxLength       =   7
      TabIndex        =   5
      Text            =   "Namn"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtFyra 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      MaxLength       =   7
      TabIndex        =   4
      Text            =   "Namn"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtTre 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      MaxLength       =   7
      TabIndex        =   3
      Text            =   "Namn"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtTvå 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      MaxLength       =   7
      TabIndex        =   2
      Text            =   "Namn"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtEn 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      MaxLength       =   7
      TabIndex        =   1
      Text            =   "Namn"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdKlar 
      Caption         =   "Starta"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Spelare 5"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Spelare 4"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Spelare 3"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Spelare 2"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Spelare 1"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmNamn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Spelare As Integer

Private Sub cmdKlar_Click()
    If frmYatzy.mnuEn.Checked = True Then
        Spelare = 1
        frmYatzy.lblSpelare1.Caption = txtEn.Text
        Starta
        frmNamn.Hide
        Exit Sub
    End If
    If frmYatzy.mnutvå.Checked = True Then
        Spelare = 2
        frmYatzy.lblSpelare1.Caption = txtEn.Text
        frmYatzy.lblSpelare2.Caption = txtTvå.Text
        Starta
        frmNamn.Hide
        Exit Sub
    End If
    If frmYatzy.mnuTre.Checked = True Then
        Spelare = 3
        frmYatzy.lblSpelare1.Caption = txtEn.Text
        frmYatzy.lblSpelare2.Caption = txtTvå.Text
        frmYatzy.lblSpelare3.Caption = txtTre.Text
        Starta
        frmNamn.Hide
        Exit Sub
    End If
    If frmYatzy.mnuFyra.Checked = True Then
        Spelare = 4
        frmYatzy.lblSpelare1.Caption = txtEn.Text
        frmYatzy.lblSpelare2.Caption = txtTvå.Text
        frmYatzy.lblSpelare3.Caption = txtTre.Text
        frmYatzy.lblSpelare4.Caption = txtFyra.Text
        Starta
        frmNamn.Hide
        Exit Sub
    End If
    If frmYatzy.mnuFem.Checked = True Then
        Spelare = 5
        frmYatzy.lblSpelare1.Caption = txtEn.Text
        frmYatzy.lblSpelare2.Caption = txtTvå.Text
        frmYatzy.lblSpelare3.Caption = txtTre.Text
        frmYatzy.lblSpelare4.Caption = txtFyra.Text
        frmYatzy.lblSpelare5.Caption = txtFem.Text
        Starta
        frmNamn.Hide
        Exit Sub
    End If
End Sub

Public Sub Starta()
    
    If Spelare = 1 Then
        frmYatzy.cmdEttor1.Visible = True
        frmYatzy.cmdTvåor1.Visible = True
        frmYatzy.cmdTreor1.Visible = True
        frmYatzy.cmdFyror1.Visible = True
        frmYatzy.cmdFemmor1.Visible = True
        frmYatzy.cmdSexor1.Visible = True
        frmYatzy.cmdSumma1.Visible = True
        frmYatzy.cmdBonus1.Visible = True
        frmYatzy.cmd1Par1.Visible = True
        frmYatzy.cmd2Par1.Visible = True
        frmYatzy.cmdTretal1.Visible = True
        frmYatzy.cmdFyrtal1.Visible = True
        frmYatzy.cmdLiten1.Visible = True
        frmYatzy.cmdStor1.Visible = True
        frmYatzy.cmdKåk1.Visible = True
        frmYatzy.cmdChans1.Visible = True
        frmYatzy.cmdYatzy1.Visible = True
        frmYatzy.cmd2Summa1.Visible = True
    
        frmYatzy.cmdEttor2.Visible = False
        frmYatzy.cmdTvåor2.Visible = False
        frmYatzy.cmdTreor2.Visible = False
        frmYatzy.cmdFyror2.Visible = False
        frmYatzy.cmdFemmor2.Visible = False
        frmYatzy.cmdSexor2.Visible = False
        frmYatzy.cmdSumma2.Visible = False
        frmYatzy.cmdBonus2.Visible = False
        frmYatzy.cmd1Par2.Visible = False
        frmYatzy.cmd2Par2.Visible = False
        frmYatzy.cmdTretal2.Visible = False
        frmYatzy.cmdFyrtal2.Visible = False
        frmYatzy.cmdLiten2.Visible = False
        frmYatzy.cmdStor2.Visible = False
        frmYatzy.cmdKåk2.Visible = False
        frmYatzy.cmdChans2.Visible = False
        frmYatzy.cmdYatzy2.Visible = False
        frmYatzy.cmd2Summa2.Visible = False
    
        frmYatzy.cmdEttor3.Visible = False
        frmYatzy.cmdTvåor3.Visible = False
        frmYatzy.cmdTreor3.Visible = False
        frmYatzy.cmdFyror3.Visible = False
        frmYatzy.cmdFemmor3.Visible = False
        frmYatzy.cmdSexor3.Visible = False
        frmYatzy.cmdSumma3.Visible = False
        frmYatzy.cmdBonus3.Visible = False
        frmYatzy.cmd1Par3.Visible = False
        frmYatzy.cmd2Par3.Visible = False
        frmYatzy.cmdTretal3.Visible = False
        frmYatzy.cmdFyrtal3.Visible = False
        frmYatzy.cmdLiten3.Visible = False
        frmYatzy.cmdStor3.Visible = False
        frmYatzy.cmdKåk3.Visible = False
        frmYatzy.cmdChans3.Visible = False
        frmYatzy.cmdYatzy3.Visible = False
        frmYatzy.cmd2Summa3.Visible = False
    
        frmYatzy.cmdEttor4.Visible = False
        frmYatzy.cmdTvåor4.Visible = False
        frmYatzy.cmdTreor4.Visible = False
        frmYatzy.cmdFyror4.Visible = False
        frmYatzy.cmdFemmor4.Visible = False
        frmYatzy.cmdSexor4.Visible = False
        frmYatzy.cmdSumma4.Visible = False
        frmYatzy.cmdBonus4.Visible = False
        frmYatzy.cmd1Par4.Visible = False
        frmYatzy.cmd2Par4.Visible = False
        frmYatzy.cmdTretal4.Visible = False
        frmYatzy.cmdFyrtal4.Visible = False
        frmYatzy.cmdLiten4.Visible = False
        frmYatzy.cmdStor4.Visible = False
        frmYatzy.cmdKåk4.Visible = False
        frmYatzy.cmdChans4.Visible = False
        frmYatzy.cmdYatzy4.Visible = False
        frmYatzy.cmd2Summa4.Visible = False
    
        frmYatzy.cmdEttor5.Visible = False
        frmYatzy.cmdTvåor5.Visible = False
        frmYatzy.cmdTreor5.Visible = False
        frmYatzy.cmdFyror5.Visible = False
        frmYatzy.cmdFemmor5.Visible = False
        frmYatzy.cmdSexor5.Visible = False
        frmYatzy.cmdSumma5.Visible = False
        frmYatzy.cmdBonus5.Visible = False
        frmYatzy.cmd1Par5.Visible = False
        frmYatzy.cmd2Par5.Visible = False
        frmYatzy.cmdTretal5.Visible = False
        frmYatzy.cmdFyrtal5.Visible = False
        frmYatzy.cmdLiten5.Visible = False
        frmYatzy.cmdStor5.Visible = False
        frmYatzy.cmdKåk5.Visible = False
        frmYatzy.cmdChans5.Visible = False
        frmYatzy.cmdYatzy5.Visible = False
        frmYatzy.cmd2Summa5.Visible = False
    
        frmYatzy.lblSpelare2.Caption = ""
        frmYatzy.lblSpelare3.Caption = ""
        frmYatzy.lblSpelare4.Caption = ""
        frmYatzy.lblSpelare5.Caption = ""
    End If
    
    If Spelare = 2 Then
        
        frmYatzy.cmdEttor1.Visible = True
        frmYatzy.cmdTvåor1.Visible = True
        frmYatzy.cmdTreor1.Visible = True
        frmYatzy.cmdFyror1.Visible = True
        frmYatzy.cmdFemmor1.Visible = True
        frmYatzy.cmdSexor1.Visible = True
        frmYatzy.cmdSumma1.Visible = True
        frmYatzy.cmdBonus1.Visible = True
        frmYatzy.cmd1Par1.Visible = True
        frmYatzy.cmd2Par1.Visible = True
        frmYatzy.cmdTretal1.Visible = True
        frmYatzy.cmdFyrtal1.Visible = True
        frmYatzy.cmdLiten1.Visible = True
        frmYatzy.cmdStor1.Visible = True
        frmYatzy.cmdKåk1.Visible = True
        frmYatzy.cmdChans1.Visible = True
        frmYatzy.cmdYatzy1.Visible = True
        frmYatzy.cmd2Summa1.Visible = True
        
        frmYatzy.cmdEttor2.Visible = True
        frmYatzy.cmdTvåor2.Visible = True
        frmYatzy.cmdTreor2.Visible = True
        frmYatzy.cmdFyror2.Visible = True
        frmYatzy.cmdFemmor2.Visible = True
        frmYatzy.cmdSexor2.Visible = True
        frmYatzy.cmdSumma2.Visible = True
        frmYatzy.cmdBonus2.Visible = True
        frmYatzy.cmd1Par2.Visible = True
        frmYatzy.cmd2Par2.Visible = True
        frmYatzy.cmdTretal2.Visible = True
        frmYatzy.cmdFyrtal2.Visible = True
        frmYatzy.cmdLiten2.Visible = True
        frmYatzy.cmdStor2.Visible = True
        frmYatzy.cmdKåk2.Visible = True
        frmYatzy.cmdChans2.Visible = True
        frmYatzy.cmdYatzy2.Visible = True
        frmYatzy.cmd2Summa2.Visible = True
    
        frmYatzy.cmdEttor3.Visible = False
        frmYatzy.cmdTvåor3.Visible = False
        frmYatzy.cmdTreor3.Visible = False
        frmYatzy.cmdFyror3.Visible = False
        frmYatzy.cmdFemmor3.Visible = False
        frmYatzy.cmdSexor3.Visible = False
        frmYatzy.cmdSumma3.Visible = False
        frmYatzy.cmdBonus3.Visible = False
        frmYatzy.cmd1Par3.Visible = False
        frmYatzy.cmd2Par3.Visible = False
        frmYatzy.cmdTretal3.Visible = False
        frmYatzy.cmdFyrtal3.Visible = False
        frmYatzy.cmdLiten3.Visible = False
        frmYatzy.cmdStor3.Visible = False
        frmYatzy.cmdKåk3.Visible = False
        frmYatzy.cmdChans3.Visible = False
        frmYatzy.cmdYatzy3.Visible = False
        frmYatzy.cmd2Summa3.Visible = False
    
        frmYatzy.cmdEttor4.Visible = False
        frmYatzy.cmdTvåor4.Visible = False
        frmYatzy.cmdTreor4.Visible = False
        frmYatzy.cmdFyror4.Visible = False
        frmYatzy.cmdFemmor4.Visible = False
        frmYatzy.cmdSexor4.Visible = False
        frmYatzy.cmdSumma4.Visible = False
        frmYatzy.cmdBonus4.Visible = False
        frmYatzy.cmd1Par4.Visible = False
        frmYatzy.cmd2Par4.Visible = False
        frmYatzy.cmdTretal4.Visible = False
        frmYatzy.cmdFyrtal4.Visible = False
        frmYatzy.cmdLiten4.Visible = False
        frmYatzy.cmdStor4.Visible = False
        frmYatzy.cmdKåk4.Visible = False
        frmYatzy.cmdChans4.Visible = False
        frmYatzy.cmdYatzy4.Visible = False
        frmYatzy.cmd2Summa4.Visible = False
        
        frmYatzy.cmdEttor5.Visible = False
        frmYatzy.cmdTvåor5.Visible = False
        frmYatzy.cmdTreor5.Visible = False
        frmYatzy.cmdFyror5.Visible = False
        frmYatzy.cmdFemmor5.Visible = False
        frmYatzy.cmdSexor5.Visible = False
        frmYatzy.cmdSumma5.Visible = False
        frmYatzy.cmdBonus5.Visible = False
        frmYatzy.cmd1Par5.Visible = False
        frmYatzy.cmd2Par5.Visible = False
        frmYatzy.cmdTretal5.Visible = False
        frmYatzy.cmdFyrtal5.Visible = False
        frmYatzy.cmdLiten5.Visible = False
        frmYatzy.cmdStor5.Visible = False
        frmYatzy.cmdKåk5.Visible = False
        frmYatzy.cmdChans5.Visible = False
        frmYatzy.cmdYatzy5.Visible = False
        frmYatzy.cmd2Summa5.Visible = False
        
        frmYatzy.lblSpelare3.Caption = ""
        frmYatzy.lblSpelare4.Caption = ""
        frmYatzy.lblSpelare5.Caption = ""
    End If
    
    If Spelare = 3 Then
        frmYatzy.cmdEttor1.Visible = True
        frmYatzy.cmdTvåor1.Visible = True
        frmYatzy.cmdTreor1.Visible = True
        frmYatzy.cmdFyror1.Visible = True
        frmYatzy.cmdFemmor1.Visible = True
        frmYatzy.cmdSexor1.Visible = True
        frmYatzy.cmdSumma1.Visible = True
        frmYatzy.cmdBonus1.Visible = True
        frmYatzy.cmd1Par1.Visible = True
        frmYatzy.cmd2Par1.Visible = True
        frmYatzy.cmdTretal1.Visible = True
        frmYatzy.cmdFyrtal1.Visible = True
        frmYatzy.cmdLiten1.Visible = True
        frmYatzy.cmdStor1.Visible = True
        frmYatzy.cmdKåk1.Visible = True
        frmYatzy.cmdChans1.Visible = True
        frmYatzy.cmdYatzy1.Visible = True
        frmYatzy.cmd2Summa1.Visible = True
        
        frmYatzy.cmdEttor2.Visible = True
        frmYatzy.cmdTvåor2.Visible = True
        frmYatzy.cmdTreor2.Visible = True
        frmYatzy.cmdFyror2.Visible = True
        frmYatzy.cmdFemmor2.Visible = True
        frmYatzy.cmdSexor2.Visible = True
        frmYatzy.cmdSumma2.Visible = True
        frmYatzy.cmdBonus2.Visible = True
        frmYatzy.cmd1Par2.Visible = True
        frmYatzy.cmd2Par2.Visible = True
        frmYatzy.cmdTretal2.Visible = True
        frmYatzy.cmdFyrtal2.Visible = True
        frmYatzy.cmdLiten2.Visible = True
        frmYatzy.cmdStor2.Visible = True
        frmYatzy.cmdKåk2.Visible = True
        frmYatzy.cmdChans2.Visible = True
        frmYatzy.cmdYatzy2.Visible = True
        frmYatzy.cmd2Summa2.Visible = True
        
        
        frmYatzy.cmdEttor3.Visible = True
        frmYatzy.cmdTvåor3.Visible = True
        frmYatzy.cmdTreor3.Visible = True
        frmYatzy.cmdFyror3.Visible = True
        frmYatzy.cmdFemmor3.Visible = True
        frmYatzy.cmdSexor3.Visible = True
        frmYatzy.cmdSumma3.Visible = True
        frmYatzy.cmdBonus3.Visible = True
        frmYatzy.cmd1Par3.Visible = True
        frmYatzy.cmd2Par3.Visible = True
        frmYatzy.cmdTretal3.Visible = True
        frmYatzy.cmdFyrtal3.Visible = True
        frmYatzy.cmdLiten3.Visible = True
        frmYatzy.cmdStor3.Visible = True
        frmYatzy.cmdKåk3.Visible = True
        frmYatzy.cmdChans3.Visible = True
        frmYatzy.cmdYatzy3.Visible = True
        frmYatzy.cmd2Summa3.Visible = True
    
        frmYatzy.cmdEttor4.Visible = False
        frmYatzy.cmdTvåor4.Visible = False
        frmYatzy.cmdTreor4.Visible = False
        frmYatzy.cmdFyror4.Visible = False
        frmYatzy.cmdFemmor4.Visible = False
        frmYatzy.cmdSexor4.Visible = False
        frmYatzy.cmdSumma4.Visible = False
        frmYatzy.cmdBonus4.Visible = False
        frmYatzy.cmd1Par4.Visible = False
        frmYatzy.cmd2Par4.Visible = False
        frmYatzy.cmdTretal4.Visible = False
        frmYatzy.cmdFyrtal4.Visible = False
        frmYatzy.cmdLiten4.Visible = False
        frmYatzy.cmdStor4.Visible = False
        frmYatzy.cmdKåk4.Visible = False
        frmYatzy.cmdChans4.Visible = False
        frmYatzy.cmdYatzy4.Visible = False
        frmYatzy.cmd2Summa4.Visible = False
    
        frmYatzy.cmdEttor5.Visible = False
        frmYatzy.cmdTvåor5.Visible = False
        frmYatzy.cmdTreor5.Visible = False
        frmYatzy.cmdFyror5.Visible = False
        frmYatzy.cmdFemmor5.Visible = False
        frmYatzy.cmdSexor5.Visible = False
        frmYatzy.cmdSumma5.Visible = False
        frmYatzy.cmdBonus5.Visible = False
        frmYatzy.cmd1Par5.Visible = False
        frmYatzy.cmd2Par5.Visible = False
        frmYatzy.cmdTretal5.Visible = False
        frmYatzy.cmdFyrtal5.Visible = False
        frmYatzy.cmdLiten5.Visible = False
        frmYatzy.cmdStor5.Visible = False
        frmYatzy.cmdKåk5.Visible = False
        frmYatzy.cmdChans5.Visible = False
        frmYatzy.cmdYatzy5.Visible = False
        frmYatzy.cmd2Summa5.Visible = False
    
        frmYatzy.lblSpelare4.Caption = ""
        frmYatzy.lblSpelare5.Caption = ""
    End If
    
    If Spelare = 4 Then
        frmYatzy.cmdEttor1.Visible = True
        frmYatzy.cmdTvåor1.Visible = True
        frmYatzy.cmdTreor1.Visible = True
        frmYatzy.cmdFyror1.Visible = True
        frmYatzy.cmdFemmor1.Visible = True
        frmYatzy.cmdSexor1.Visible = True
        frmYatzy.cmdSumma1.Visible = True
        frmYatzy.cmdBonus1.Visible = True
        frmYatzy.cmd1Par1.Visible = True
        frmYatzy.cmd2Par1.Visible = True
        frmYatzy.cmdTretal1.Visible = True
        frmYatzy.cmdFyrtal1.Visible = True
        frmYatzy.cmdLiten1.Visible = True
        frmYatzy.cmdStor1.Visible = True
        frmYatzy.cmdKåk1.Visible = True
        frmYatzy.cmdChans1.Visible = True
        frmYatzy.cmdYatzy1.Visible = True
        frmYatzy.cmd2Summa1.Visible = True
        
        frmYatzy.cmdEttor2.Visible = True
        frmYatzy.cmdTvåor2.Visible = True
        frmYatzy.cmdTreor2.Visible = True
        frmYatzy.cmdFyror2.Visible = True
        frmYatzy.cmdFemmor2.Visible = True
        frmYatzy.cmdSexor2.Visible = True
        frmYatzy.cmdSumma2.Visible = True
        frmYatzy.cmdBonus2.Visible = True
        frmYatzy.cmd1Par2.Visible = True
        frmYatzy.cmd2Par2.Visible = True
        frmYatzy.cmdTretal2.Visible = True
        frmYatzy.cmdFyrtal2.Visible = True
        frmYatzy.cmdLiten2.Visible = True
        frmYatzy.cmdStor2.Visible = True
        frmYatzy.cmdKåk2.Visible = True
        frmYatzy.cmdChans2.Visible = True
        frmYatzy.cmdYatzy2.Visible = True
        frmYatzy.cmd2Summa2.Visible = True
        
        
        frmYatzy.cmdEttor4.Visible = True
        frmYatzy.cmdTvåor4.Visible = True
        frmYatzy.cmdTreor4.Visible = True
        frmYatzy.cmdFyror4.Visible = True
        frmYatzy.cmdFemmor4.Visible = True
        frmYatzy.cmdSexor4.Visible = True
        frmYatzy.cmdSumma4.Visible = True
        frmYatzy.cmdBonus4.Visible = True
        frmYatzy.cmd1Par4.Visible = True
        frmYatzy.cmd2Par4.Visible = True
        frmYatzy.cmdTretal4.Visible = True
        frmYatzy.cmdFyrtal4.Visible = True
        frmYatzy.cmdLiten4.Visible = True
        frmYatzy.cmdStor4.Visible = True
        frmYatzy.cmdKåk4.Visible = True
        frmYatzy.cmdChans4.Visible = True
        frmYatzy.cmdYatzy4.Visible = True
        frmYatzy.cmd2Summa4.Visible = True
    
        frmYatzy.cmdEttor3.Visible = True
        frmYatzy.cmdTvåor3.Visible = True
        frmYatzy.cmdTreor3.Visible = True
        frmYatzy.cmdFyror3.Visible = True
        frmYatzy.cmdFemmor3.Visible = True
        frmYatzy.cmdSexor3.Visible = True
        frmYatzy.cmdSumma3.Visible = True
        frmYatzy.cmdBonus3.Visible = True
        frmYatzy.cmd1Par3.Visible = True
        frmYatzy.cmd2Par3.Visible = True
        frmYatzy.cmdTretal3.Visible = True
        frmYatzy.cmdFyrtal3.Visible = True
        frmYatzy.cmdLiten3.Visible = True
        frmYatzy.cmdStor3.Visible = True
        frmYatzy.cmdKåk3.Visible = True
        frmYatzy.cmdChans3.Visible = True
        frmYatzy.cmdYatzy3.Visible = True
        frmYatzy.cmd2Summa3.Visible = True
    
        frmYatzy.cmdEttor5.Visible = False
        frmYatzy.cmdTvåor5.Visible = False
        frmYatzy.cmdTreor5.Visible = False
        frmYatzy.cmdFyror5.Visible = False
        frmYatzy.cmdFemmor5.Visible = False
        frmYatzy.cmdSexor5.Visible = False
        frmYatzy.cmdSumma5.Visible = False
        frmYatzy.cmdBonus5.Visible = False
        frmYatzy.cmd1Par5.Visible = False
        frmYatzy.cmd2Par5.Visible = False
        frmYatzy.cmdTretal5.Visible = False
        frmYatzy.cmdFyrtal5.Visible = False
        frmYatzy.cmdLiten5.Visible = False
        frmYatzy.cmdStor5.Visible = False
        frmYatzy.cmdKåk5.Visible = False
        frmYatzy.cmdChans5.Visible = False
        frmYatzy.cmdYatzy5.Visible = False
        frmYatzy.cmd2Summa5.Visible = False
    
        frmYatzy.lblSpelare5.Caption = ""
    End If

    If Spelare = 5 Then
        frmYatzy.cmdEttor1.Visible = True
        frmYatzy.cmdTvåor1.Visible = True
        frmYatzy.cmdTreor1.Visible = True
        frmYatzy.cmdFyror1.Visible = True
        frmYatzy.cmdFemmor1.Visible = True
        frmYatzy.cmdSexor1.Visible = True
        frmYatzy.cmdSumma1.Visible = True
        frmYatzy.cmdBonus1.Visible = True
        frmYatzy.cmd1Par1.Visible = True
        frmYatzy.cmd2Par1.Visible = True
        frmYatzy.cmdTretal1.Visible = True
        frmYatzy.cmdFyrtal1.Visible = True
        frmYatzy.cmdLiten1.Visible = True
        frmYatzy.cmdStor1.Visible = True
        frmYatzy.cmdKåk1.Visible = True
        frmYatzy.cmdChans1.Visible = True
        frmYatzy.cmdYatzy1.Visible = True
        frmYatzy.cmd2Summa1.Visible = True
        
        frmYatzy.cmdEttor2.Visible = True
        frmYatzy.cmdTvåor2.Visible = True
        frmYatzy.cmdTreor2.Visible = True
        frmYatzy.cmdFyror2.Visible = True
        frmYatzy.cmdFemmor2.Visible = True
        frmYatzy.cmdSexor2.Visible = True
        frmYatzy.cmdSumma2.Visible = True
        frmYatzy.cmdBonus2.Visible = True
        frmYatzy.cmd1Par2.Visible = True
        frmYatzy.cmd2Par2.Visible = True
        frmYatzy.cmdTretal2.Visible = True
        frmYatzy.cmdFyrtal2.Visible = True
        frmYatzy.cmdLiten2.Visible = True
        frmYatzy.cmdStor2.Visible = True
        frmYatzy.cmdKåk2.Visible = True
        frmYatzy.cmdChans2.Visible = True
        frmYatzy.cmdYatzy2.Visible = True
        frmYatzy.cmd2Summa2.Visible = True
        
        frmYatzy.cmdEttor3.Visible = True
        frmYatzy.cmdTvåor3.Visible = True
        frmYatzy.cmdTreor3.Visible = True
        frmYatzy.cmdFyror3.Visible = True
        frmYatzy.cmdFemmor3.Visible = True
        frmYatzy.cmdSexor3.Visible = True
        frmYatzy.cmdSumma3.Visible = True
        frmYatzy.cmdBonus3.Visible = True
        frmYatzy.cmd1Par3.Visible = True
        frmYatzy.cmd2Par3.Visible = True
        frmYatzy.cmdTretal3.Visible = True
        frmYatzy.cmdFyrtal3.Visible = True
        frmYatzy.cmdLiten3.Visible = True
        frmYatzy.cmdStor3.Visible = True
        frmYatzy.cmdKåk3.Visible = True
        frmYatzy.cmdChans3.Visible = True
        frmYatzy.cmdYatzy3.Visible = True
        frmYatzy.cmd2Summa3.Visible = True
        
        frmYatzy.cmdEttor4.Visible = True
        frmYatzy.cmdTvåor4.Visible = True
        frmYatzy.cmdTreor4.Visible = True
        frmYatzy.cmdFyror4.Visible = True
        frmYatzy.cmdFemmor4.Visible = True
        frmYatzy.cmdSexor4.Visible = True
        frmYatzy.cmdSumma4.Visible = True
        frmYatzy.cmdBonus4.Visible = True
        frmYatzy.cmd1Par4.Visible = True
        frmYatzy.cmd2Par4.Visible = True
        frmYatzy.cmdTretal4.Visible = True
        frmYatzy.cmdFyrtal4.Visible = True
        frmYatzy.cmdLiten4.Visible = True
        frmYatzy.cmdStor4.Visible = True
        frmYatzy.cmdKåk4.Visible = True
        frmYatzy.cmdChans4.Visible = True
        frmYatzy.cmdYatzy4.Visible = True
        frmYatzy.cmd2Summa4.Visible = True
        
        
        frmYatzy.cmdEttor5.Visible = True
        frmYatzy.cmdTvåor5.Visible = True
        frmYatzy.cmdTreor5.Visible = True
        frmYatzy.cmdFyror5.Visible = True
        frmYatzy.cmdFemmor5.Visible = True
        frmYatzy.cmdSexor5.Visible = True
        frmYatzy.cmdSumma5.Visible = True
        frmYatzy.cmdBonus5.Visible = True
        frmYatzy.cmd1Par5.Visible = True
        frmYatzy.cmd2Par5.Visible = True
        frmYatzy.cmdTretal5.Visible = True
        frmYatzy.cmdFyrtal5.Visible = True
        frmYatzy.cmdLiten5.Visible = True
        frmYatzy.cmdStor5.Visible = True
        frmYatzy.cmdKåk5.Visible = True
        frmYatzy.cmdChans5.Visible = True
        frmYatzy.cmdYatzy5.Visible = True
        frmYatzy.cmd2Summa5.Visible = True
    End If
    frmYatzy.cmdNästa.Enabled = False
    frmYatzy.cmdKasta.Enabled = True
    frmYatzy.cmdKlara.Enabled = True
End Sub

Private Sub opt1_Click()
    opt1.Value = True
    Spelare = 1
    frmYatzy.mnuEn.Checked = True
    frmYatzy.mnutvå.Checked = False
    frmYatzy.mnuTre.Checked = False
    frmYatzy.mnuFyra.Checked = False
    frmYatzy.mnuFem.Checked = False
    txtEn.Enabled = True
    txtTvå.Enabled = False
    txtTre.Enabled = False
    txtFyra.Enabled = False
    txtFem.Enabled = False
End Sub

Private Sub opt2_Click()
    opt2.Value = True
    Spelare = 2
    frmYatzy.mnuEn.Checked = False
    frmYatzy.mnutvå.Checked = True
    frmYatzy.mnuTre.Checked = False
    frmYatzy.mnuFyra.Checked = False
    frmYatzy.mnuFem.Checked = False
    txtEn.Enabled = True
    txtTvå.Enabled = True
    txtTre.Enabled = False
    txtFyra.Enabled = False
    txtFem.Enabled = False
End Sub

Private Sub opt3_Click()
    opt3.Value = True
    Spelare = 3
    frmYatzy.mnuEn.Checked = False
    frmYatzy.mnutvå.Checked = False
    frmYatzy.mnuTre.Checked = True
    frmYatzy.mnuFyra.Checked = False
    frmYatzy.mnuFem.Checked = False
    txtEn.Enabled = True
    txtTvå.Enabled = True
    txtTre.Enabled = True
    txtFyra.Enabled = False
    txtFem.Enabled = False
End Sub

Private Sub opt4_Click()
    opt4.Value = True
    Spelare = 4
    frmYatzy.mnuEn.Checked = False
    frmYatzy.mnutvå.Checked = False
    frmYatzy.mnuTre.Checked = False
    frmYatzy.mnuFyra.Checked = True
    frmYatzy.mnuFem.Checked = False
    
    txtEn.Enabled = True
    txtTvå.Enabled = True
    txtTre.Enabled = True
    txtFyra.Enabled = True
    txtFem.Enabled = False
End Sub

Private Sub opt5_Click()
    opt5.Value = True
    Spelare = 5
    frmYatzy.mnuEn.Checked = False
    frmYatzy.mnutvå.Checked = False
    frmYatzy.mnuTre.Checked = False
    frmYatzy.mnuFyra.Checked = False
    frmYatzy.mnuFem.Checked = True
    
    txtEn.Enabled = True
    txtTvå.Enabled = True
    txtTre.Enabled = True
    txtFyra.Enabled = True
    txtFem.Enabled = True

End Sub
