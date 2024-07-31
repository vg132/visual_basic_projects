VERSION 5.00
Begin VB.Form frmMem 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Info Mem - www.vgsoftware.com"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3360
   Icon            =   "frmMem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   1680
      Top             =   720
   End
   Begin VB.Label lblMemStat 
      Caption         =   "Label7"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblMemStat 
      Caption         =   "Label6"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblMemStat 
      Caption         =   "Label5"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label lblMemStat 
      Caption         =   "Label4"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label lblMemStat 
      Caption         =   "Label3"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label lblMemStat 
      Caption         =   "Label2"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label lblMemStat 
      Caption         =   "Label1"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   3135
   End
End
Attribute VB_Name = "frmMem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Const fmt As String = "###,###,###,###"
Private Const skb As String = " Kbyte"
Private Const nkb As Long = 1024

Private Sub Form_Load()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Private Sub Timer1_Timer()
Dim MS As MEMORYSTATUS
    MS.dwLength = Len(MS)
    GlobalMemoryStatus MS
    lblMemStat(0) = "Memory Load:            " & Format$(MS.dwMemoryLoad, fmt) & " % used"
    lblMemStat(1) = "Total Physical:           " & Format$(MS.dwTotalPhys / nkb, fmt) & skb
    lblMemStat(2) = "Free Physcal:             " & Format$(MS.dwAvailPhys / nkb, fmt) & skb
    lblMemStat(3) = "Page File Size:           " & Format$(MS.dwTotalPageFile / nkb, fmt) & skb
    lblMemStat(4) = "Free Page File Size:   " & Format$(MS.dwAvailPageFile / nkb, fmt) & skb
    lblMemStat(5) = "Total Virtual Memory:  " & Format$(MS.dwTotalVirtual / nkb, fmt) & skb
    lblMemStat(6) = "Free Virtual Memory:   " & Format$(MS.dwAvailVirtual / nkb, fmt) & skb
End Sub
