VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Proffice Tidrapport© Viktor Gars 1999"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTotalt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtLunch 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Text            =   "60"
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtSluta 
      Height          =   285
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtBorja 
      Height          =   285
      Left            =   240
      MaxLength       =   5
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Räkna"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label Label4 
      Caption         =   "Total Tid"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Lunch (Min)"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Slutar"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Börjar"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TextLen As String

Private Sub Command1_Click()
Dim TotalH As Integer
Dim TotalM As Integer
    
    
    TotalM = DateDiff("n", txtBorja.Text, txtSluta.Text)
    TotalM = TotalM - txtLunch.Text
    Do Until TotalM < 60
        TotalH = TotalH + 1
        TotalM = TotalM - 60
    Loop
    
    TotalM = (TotalM / 60) * 100
    If TotalM < 12.5 Then
        TotalM = 0
    ElseIf (TotalM > 12.5) And (TotalM < 37.5) Then
        TotalM = 25
    ElseIf (TotalM > 37.5) And (TotalM < 62.5) Then
        TotalM = 50
    ElseIf (TotalM > 62.5) And (TotalM < 87.5) Then
        TotalM = 75
    ElseIf TotalM > 87.5 Then
        TotalM = 0
        TotalH = TotalH + 1
    End If

    If Len(Trim(Str(TotalM))) = 1 Then
        txtTotalt.Text = TotalH & ":" & "0" & TotalM
    Else
        txtTotalt.Text = TotalH & ":" & TotalM
    End If
    SendKeys ("{Tab}")
End Sub

Public Sub TextSelected()

Dim i As Integer
Dim oMyTextBox As Object

Set oMyTextBox = Screen.ActiveControl
    If TypeName(oMyTextBox) = "TextBox" Then
        i = Len(oMyTextBox.Text)
        oMyTextBox.SelStart = 0
        oMyTextBox.SelLength = i
    End If

End Sub

Private Sub txtBorja_Change()
    TextLen = txtBorja.Text
    If Len(TextLen) = 2 Then txtBorja.Text = txtBorja.Text & ":"
    SendKeys ("^{END}")
    If Len(TextLen) = 5 Then
        SendKeys ("{Tab}")
    End If
End Sub

Private Sub txtBorja_GotFocus()
    TextSelected
End Sub

Private Sub txtLunch_GotFocus()
    TextSelected
End Sub

Private Sub txtSluta_Change()
    TextLen = txtSluta.Text
    If Len(TextLen) = 2 Then txtSluta.Text = txtSluta.Text & ":"
    SendKeys ("^{END}")
    If Len(TextLen) = 5 Then
        SendKeys ("{Tab}")
    End If
End Sub

Private Sub txtSluta_GotFocus()
    TextSelected
End Sub

Private Sub txtTotalt_GotFocus()
    TextSelected
End Sub
