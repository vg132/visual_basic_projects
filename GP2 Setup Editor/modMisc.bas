Attribute VB_Name = "modMisc"
Public X As Long
Public Read As String

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Const GWL_STYLE = (-16)
Public Const ES_NUMBER = &H2000&

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

Public Function OFile(ByVal Title As String, Filter As String, Dir As String) As String
Dim OpenFile As OPENFILENAME
Dim lReturn As Long
Dim sFilter As String
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = frmSetup.hwnd
    OpenFile.hInstance = App.hInstance
    OpenFile.lpstrFilter = Filter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = Dir
    OpenFile.lpstrTitle = Title
    OpenFile.flags = 0
    lReturn = GetOpenFileName(OpenFile)
    If lReturn = 0 Then
        OFile = ""
    Else
        OFile = OpenFile.lpstrFile
    End If
End Function

Public Function SFile(ByVal Title As String, Filter As String, Dir As String) As String
Dim SaveFile As OPENFILENAME
Dim lReturn As Long
Dim sFilter As String
    SaveFile.lStructSize = Len(SaveFile)
    SaveFile.hwndOwner = frmSetup.hwnd
    SaveFile.hInstance = App.hInstance
    SaveFile.lpstrFilter = Filter
    SaveFile.nFilterIndex = 1
    SaveFile.lpstrFile = String(257, 0)
    SaveFile.nMaxFile = Len(SaveFile.lpstrFile) - 1
    SaveFile.lpstrFileTitle = SaveFile.lpstrFile
    SaveFile.nMaxFileTitle = SaveFile.nMaxFile
    SaveFile.lpstrInitialDir = Dir
    SaveFile.lpstrTitle = Title
    SaveFile.flags = 0
    lReturn = GetSaveFileName(SaveFile)
    If lReturn = 0 Then
        SFile = ""
    Else
        SFile = SaveFile.lpstrFile
    End If
End Function
