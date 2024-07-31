Attribute VB_Name = "modMisc"
Option Explicit

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Function ShowOpen(ByVal Filter As String, ByVal hWnd As Long, ByVal InitDir As String, Optional Title As String, Optional DefName As String) As String
Dim OpenFile As OPENFILENAME
Dim lReturn As Long
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = hWnd
    OpenFile.hInstance = App.hInstance
    OpenFile.lpstrFilter = ConvertFilter(Filter)
    OpenFile.nFilterIndex = 1
    If DefName = "" Then
        OpenFile.lpstrFile = Space$(1024) & Chr(0)
    Else
        OpenFile.lpstrFile = DefName & Space$(1024) & Chr(0)
    End If
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile)
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile & Chr(0)
    OpenFile.lpstrInitialDir = InitDir & Chr(0)
    OpenFile.flags = 0
    If Title <> "" Then OpenFile.lpstrTitle = Title
    lReturn = GetOpenFileName(OpenFile)
    If lReturn = 0 Then
        ShowOpen = ""
    Else
        ShowOpen = Left(OpenFile.lpstrFile, InStr(OpenFile.lpstrFile, vbNullChar) - 1)
    End If
End Function

Public Function ShowSave(ByVal Filter As String, ByVal DefExt As String, ByVal hWnd As Long, ByVal InitDir As String, Optional Title As String, Optional DefName As String) As String
Dim OpenFile As OPENFILENAME
Dim lReturn As Long
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = hWnd
    OpenFile.hInstance = App.hInstance
    OpenFile.lpstrFilter = ConvertFilter(Filter)
    OpenFile.nFilterIndex = 1
    If DefName = "" Then
        OpenFile.lpstrFile = Space$(1024) & Chr(0)
    Else
        OpenFile.lpstrFile = DefName & Space$(1024) & Chr(0)
    End If
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile)
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile & Chr(0)
    OpenFile.lpstrInitialDir = InitDir & Chr(0)
    OpenFile.flags = 0
    OpenFile.lpstrDefExt = DefExt
    If Title <> "" Then OpenFile.lpstrTitle = Title
    lReturn = GetSaveFileName(OpenFile)
    If lReturn = 0 Then
        ShowSave = ""
    Else
        ShowSave = Left(OpenFile.lpstrFile, InStr(OpenFile.lpstrFile, vbNullChar) - 1)
    End If
End Function

Private Function ConvertFilter(ByVal Filter As String) As String
Dim X As Long
Dim Read As String
    Read = ""
    X = 10
    Do Until X = 0
        X = InStr(1, Filter, "|")
        If X <> 0 Then
            Read = Read & Mid(Filter, 1, X - 1) & Chr(0)
            Filter = Mid(Filter, X + 1)
        End If
    Loop
    ConvertFilter = Read & Chr(0)
End Function

Public Sub SelectText()
Dim oText As Object
    Set oText = Screen.ActiveControl
    If TypeName(oText) = "TextBox" Then
        oText.SelStart = 0
        oText.SelLength = Len(oText.Text)
    End If
End Sub
