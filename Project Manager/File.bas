Attribute VB_Name = "File"
Option Explicit
Public Java As String
Public Javac As String

Public OpenFileName As String

Public Sub SaveProjectAs(ByVal hWnd As Long)
Dim FileNum As Integer
Dim sValue As String
Dim X As Long
    OpenFileName = ShowSave("Project file (*.ipm)|*.ipm|All files (*.*)|*.*|", "ipm", 0, , "Save Project")
    If OpenFileName = "" Then Exit Sub
    FileNum = FreeFile
    Open OpenFileName For Append As FileNum
    For X = 1 To frmManager.lstFile.ListItems.Count
        sValue = frmManager.lstFile.ListItems.Item(X).Key
        Print #FileNum, sValue
    Next
    Close FileNum
End Sub

Public Sub SaveProject(ByVal hWnd As Long)
Dim X As Integer
Dim FileNum As Integer
Dim sValue As String
    If OpenFileName = "" Then
        SaveProjectAs hWnd
    Else
        Kill (OpenFileName)
        DoEvents
        FileNum = FreeFile
        Open OpenFileName For Append As FileNum
        For X = 1 To frmManager.lstFile.ListItems.Count
            sValue = frmManager.lstFile.ListItems.Item(X).Key
            Print #FileNum, sValue
        Next
        Close FileNum
    End If
End Sub

Public Sub OpenProject(ByVal hWnd As Long)
Dim sTempFileName As String
Dim FileNum As Integer
Dim sLine As String
    sTempFileName = ShowOpen("Project file (*.ipm)|*.ipm|All files (*.*)|*.*|", hWnd, , "Open Project File")
    If sTempFileName = "" Then Exit Sub
    FileNum = FreeFile
    Open sTempFileName For Input As FileNum
    Do While Not EOF(FileNum)
        Line Input #FileNum, sLine
        If GetFilePart(sLine, GetExt) = "java" Then
            frmManager.lstFile.ListItems.Add , sLine, GetFilePart(sLine, GetFileName), "java", "java"
        ElseIf GetFilePart(sLine, GetExt) = "cpp" Then
            frmManager.lstFile.ListItems.Add , sLine, GetFilePart(sLine, GetFileName), "cpp", "cpp"
        ElseIf GetFilePart(sLine, GetExt) = "c" Then
            frmManager.lstFile.ListItems.Add , sLine, GetFilePart(sLine, GetFileName), "c", "c"
        ElseIf GetFilePart(sLine, GetExt) = "txt" Then
            frmManager.lstFile.ListItems.Add , sLine, GetFilePart(sLine, GetFileName), "txt", "txt"
        Else
            frmManager.lstFile.ListItems.Add , sLine, GetFilePart(sLine, GetFileName), "other", "other"
        End If
    Loop
    Close FileNum
End Sub

Public Sub AddFile()
Dim FileName As String
Dim RetVal As Long
Dim FileNum As Integer
    FileName = ShowOpen("All files (*.*)|*.*|Java file (*.java)|*.java|C++ files (*.cpp)|*.cpp|C files (*.c)|*.c|", 0, "", "Add file to Project")
    If FileName = "" Then Exit Sub
    If FileExists(FileName) = False Then
        RetVal = MsgBox("This file was not found." & vbLf & "Do you wan't to create it now?", vbYesNoCancel + vbQuestion, "File not found")
        If (RetVal = vbCancel) Or (RetVal = vbNo) Then
            Exit Sub
        ElseIf RetVal = vbYes Then
            FileNum = FreeFile
            Open FileName For Binary As FileNum
            Close FileNum
        End If
    End If
    If GetFilePart(FileName, GetExt) = "java" Then
        frmManager.lstFile.ListItems.Add , FileName, GetFilePart(FileName, GetFileName), "java", "java"
    ElseIf GetFilePart(FileName, GetExt) = "cpp" Then
        frmManager.lstFile.ListItems.Add , FileName, GetFilePart(FileName, GetFileName), "cpp", "cpp"
    ElseIf GetFilePart(FileName, GetExt) = "c" Then
        frmManager.lstFile.ListItems.Add , FileName, GetFilePart(FileName, GetFileName), "c", "c"
    ElseIf GetFilePart(FileName, GetExt) = "txt" Then
        frmManager.lstFile.ListItems.Add , FileName, GetFilePart(FileName, GetFileName), "txt", "txt"
    Else
        frmManager.lstFile.ListItems.Add , FileName, GetFilePart(FileName, GetFileName), "other", "other"
    End If
End Sub

Public Sub RemoveFile()
Dim RetVal As Long
    RetVal = MsgBox("Do you wan't to remove this file?", vbYesNo + vbQuestion, "Remove file")
    If RetVal = vbYes Then
        frmManager.lstFile.ListItems.Remove (frmManager.lstFile.SelectedItem.Index)
    End If
End Sub
