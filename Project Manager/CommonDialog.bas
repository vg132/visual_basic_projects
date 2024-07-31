Attribute VB_Name = "CommonDialog"
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFileName) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OpenFileName) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private lpIDList As Long
Private sBuffer As String
Private szTitle As String
Private tBrowseInfo As BrowseInfo
Private FileDir As String
Private FileNum As Integer

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260&

Private Type BrowseInfo
    hwndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Type OpenFileName
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

Public Enum GetFilePartEnum
    GetExt = 0
    GetFileName = 1
    GetFilePath = 2
End Enum

Public Function ShowOpen(ByVal Filter As String, ByVal hWnd As Long, Optional InitDir As String, Optional Title As String, Optional DefName As String) As String
Dim OpenFile As OpenFileName
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
    If InitDir = "" Then
        OpenFile.lpstrInitialDir = FileDir & Chr(0)
    Else
        OpenFile.lpstrInitialDir = InitDir & Chr(0)
    End If
    OpenFile.flags = &H4
    If Title <> "" Then OpenFile.lpstrTitle = Title
    lReturn = GetOpenFileName(OpenFile)
    If lReturn = 0 Then
        ShowOpen = ""
    Else
        If InitDir = "" Then
            For X = Len(OpenFile.lpstrFile) To 0 Step -1
                If Mid(OpenFile.lpstrFile, X, 1) = "\" Then Exit For
            Next
            FileDir = Mid(OpenFile.lpstrFile, 1, X - 1)
        End If
        ShowOpen = Left(OpenFile.lpstrFile, InStr(OpenFile.lpstrFile, vbNullChar) - 1)
    End If
End Function

Public Function ShowSave(ByVal Filter As String, ByVal DefExt As String, ByVal hWnd As Long, Optional InitDir As String, Optional Title As String, Optional DefName As String) As String
Dim OpenFile As OpenFileName
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
    If InitDir = "" Then
        OpenFile.lpstrInitialDir = FileDir & Chr(0)
    Else
        OpenFile.lpstrInitialDir = InitDir & Chr(0)
    End If
    OpenFile.flags = &H4
    OpenFile.lpstrDefExt = DefExt
    If Title <> "" Then OpenFile.lpstrTitle = Title
    lReturn = GetSaveFileName(OpenFile)
    If lReturn = 0 Then
        ShowSave = ""
    Else
        If InitDir = "" Then
            For X = Len(OpenFile.lpstrFile) To 0 Step -1
                If Mid(OpenFile.lpstrFile, X, 1) = "\" Then Exit For
            Next
            FileDir = Mid(OpenFile.lpstrFile, 1, X - 1)
        End If
        ShowSave = Left(OpenFile.lpstrFile, InStr(OpenFile.lpstrFile, vbNullChar) - 1)
    End If
End Function

Public Function FileExists(ByVal PathName As String) As Boolean
       FileExists = IIf(Dir$(PathName) = "", False, True)
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

Public Function BrowseFolders(ByVal Title As String, ByVal hWnd As Long) As String
    szTitle = Title
    With tBrowseInfo
        .hwndOwner = hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        BrowseFolders = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    End If
End Function

Public Function GetFilePart(ByVal File As String, ByVal Info As GetFilePartEnum) As String
Dim X As Integer
Dim Ext As String
Dim Temp As String
    Select Case Info
    Case 0
        For X = Len(File) To 0 Step -1
            Temp = Mid(File, X, 1)
            If Temp <> "." Then
                Ext = Temp & Ext
            Else
                Exit For
            End If
        Next
        GetFilePart = LCase(Ext)
    Case 1
        For X = Len(File) To 1 Step -1
            If Mid(File, X, 1) = "\" Then Exit For
        Next
        GetFilePart = Mid(File, X + 1)
    Case 2
        For X = Len(File) To 1 Step -1
            If Mid(File, X, 1) = "\" Then Exit For
        Next
        GetFilePart = Mid(File, 1, X - 1)
    End Select
End Function
