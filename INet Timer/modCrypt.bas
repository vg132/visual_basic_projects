Attribute VB_Name = "modCrypt"
Private Const base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

Function Crypt(ByVal DecryptedText As String) As String
Dim c1, c2, c3 As Integer
Dim w1 As Integer
Dim w2 As Integer
Dim w3 As Integer
Dim w4 As Integer
Dim n As Integer
Dim Retry As String

    For n = 1 To Len(DecryptedText) Step 3
        c1 = Asc(Mid$(DecryptedText, n, 1))
        c2 = Asc(Mid$(DecryptedText, n + 1, 1) + Chr$(0))
        c3 = Asc(Mid$(DecryptedText, n + 2, 1) + Chr$(0))
        w1 = Int(c1 / 4)
        w2 = (c1 And 3) * 16 + Int(c2 / 16)
        If Len(DecryptedText) >= n + 1 Then w3 = (c2 And 15) * 4 + Int(c3 / 64) Else w3 = -1
        If Len(DecryptedText) >= n + 2 Then w4 = c3 And 63 Else w4 = -1
        Retry = Retry + MimeenCode(w1) + MimeenCode(w2) + MimeenCode(w3) + MimeenCode(w4)
    Next
    Crypt = Chr2Asc(Retry)
End Function

Function DeCrypt(ByVal a As String) As String
Dim w1 As Integer
Dim w2 As Integer
Dim w3 As Integer
Dim w4 As Integer
Dim n As Integer
Dim Retry As String
    a = Asc2Chr(a)
    For n = 1 To Len(a) Step 4
        w1 = MimedeCode(Mid$(a, n, 1))
        w2 = MimedeCode(Mid$(a, n + 1, 1))
        w3 = MimedeCode(Mid$(a, n + 2, 1))
        w4 = MimedeCode(Mid$(a, n + 3, 1))
        If w2 >= 0 Then Retry = Retry + Chr$(((w1 * 4 + Int(w2 / 16)) And 255))
        If w3 >= 0 Then Retry = Retry + Chr$(((w2 * 16 + Int(w3 / 4)) And 255))
        If w4 >= 0 Then Retry = Retry + Chr$(((w3 * 64 + w4) And 255))
    Next
    DeCrypt = Retry
End Function

Private Function MimeenCode(w As Integer) As String
    If w >= 0 Then MimeenCode = Mid$(base64, w + 1, 1) Else MimeenCode = ""
End Function

Private Function MimedeCode(a As String) As Integer
    If Len(a) = 0 Then MimedeCode = -1: Exit Function
    MimedeCode = InStr(base64, a) - 1
End Function

Public Function Asc2Chr(ByVal sText As String) As String
Dim X As Integer
Dim sTemp As String
Dim sRead As String
    sRead = ""
    For X = 1 To Len(sText) Step 3
        sTemp = Mid(sText, X, 3)
        sRead = sRead & Chr(sTemp)
    Next
    Asc2Chr = sRead
End Function

Public Function Chr2Asc(ByVal sText As String) As String
Dim X As Integer
Dim sTemp As String
Dim sRead As String
    sRead = ""
    For X = 1 To Len(sText)
        sTemp = Mid(sText, X, 1)
        sTemp = Asc(sTemp)
        If Len(sTemp) = 3 Then
            sRead = sRead & sTemp
        ElseIf Len(sTemp) = 2 Then
            sRead = sRead & "0" & sTemp
        ElseIf Len(sTemp) = 1 Then
            sRead = sRead & "00" & sTemp
        End If
    Next
    Chr2Asc = sRead
End Function

