Attribute VB_Name = "modMisc"
Option Explicit
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetUserNameAPI Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function InternetGetConnectedStateEx Lib "WinINET.DLL" Alias "InternetGetConnectedStateExA" (lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Long, ByVal dwReserved As Long) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Const INTERNET_CONNECTION_MODEM = &H1&

Public Const HWND_BOTTOM = 1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1

Public Const SWP_DRAWFRAME = &H20
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40

Public INetData As tData
Public oReg As New oReg

Public CheckUser As Boolean
Public CheckUserName As String

Public Type tData
    Price As Double
    OnTime As String
    Tariff As Double
    User As String
    TotPrice As Double
    CountTime As String
    ConName As String
End Type

Public Function ActiveConnection() As Boolean
'*************************************
'Function Name: InternetConnected
'Use: check if the computer is connected to internet.
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-??-??
'*************************************
Dim lpData As Variant
    lpData = 0
    lpData = oReg.GetValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\RemoteAccess", "Remote Connection")
    If Not IsArray(lpData) = True Then Exit Function
    If lpData(0) = 0 Then
        ActiveConnection = False 'False
    Else
        ActiveConnection = True 'True
    End If
End Function

Public Function Sec2Time(ByVal Sec As Long) As String
'*************************************
'Function Name: Sec2Time
'Use: Convert seconds to a time with Min and Hours
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-17
'*************************************

Dim Hour As Long
Dim Min As Long
Dim X As Long
    Do Until Sec < 3600
        Hour = Hour + 1
        Sec = Sec - 3600
    Loop
    Do Until Sec < 60
        Min = Min + 1
        Sec = Sec - 60
    Loop
    If Hour > 0 Then
        Sec2Time = Hour & ":"
    End If
    If Min < 10 Then
        If Sec2Time = "" Then
            Sec2Time = Sec2Time & Min & ":"
        Else
            Sec2Time = Sec2Time & "0" & Min & ":"
        End If
    Else
        Sec2Time = Sec2Time & Min & ":"
    End If
    If Sec < 10 Then
        Sec2Time = Sec2Time & "0" & Sec
    Else
        Sec2Time = Sec2Time & Sec
    End If
End Function

Public Function Cent2Dollar(ByVal lCent As Double) As String
'*************************************
'Function Name: Cent2Dollar
'Use: Convert Cent to dollar and cents
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-17
'*************************************
Dim Dollar As Long
Dim Cent As Long
Dim X As Long
    Do Until lCent < 100
        Dollar = Dollar + 1
        lCent = lCent - 100
    Loop
    If lCent > 9 Then
        Cent2Dollar = Dollar & "." & Round(lCent, 0)
    Else
        Cent2Dollar = Dollar & ".0" & Round(lCent, 0)
    End If
End Function

Public Function GetUserName() As String
'*************************************
'Function Name: GetUser
'Use: Get the username of current user of Windows
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-10-09
'*************************************
Dim sUser As String
Dim sLength As Long
Dim RetVal As Long
    sUser = Space(255)
    sLength = 255
    RetVal = GetUserNameAPI(sUser, sLength)
    GetUserName = Left(sUser, sLength - 1)
End Function

Private Function InternetConnected(lConnectionInfo As Long, sConnectionName As String) As Boolean
'*************************************
'Function Name: InternetConnected
'Use: Se next function
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-??-??
'*************************************

Dim iPosition As Integer
    sConnectionName = String(513, vbNullChar)
    InternetConnected = (InternetGetConnectedStateEx(lConnectionInfo, sConnectionName, 512, 0) <> 0)
    iPosition = InStr(sConnectionName, vbNullChar)
    If iPosition > 0 Then
        sConnectionName = Left(sConnectionName, iPosition - 1)
    ElseIf sConnectionName = String(513, vbNullChar) Then
        sConnectionName = ""
    End If
End Function

Public Function GetConName() As String
'*************************************
'Function Name: InternetConnected
'Use: Get the name of the internet connection the is
'currently used
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-??-??
'*************************************

Dim lConnectionInfo As Long
Dim sConnectionName As String
Dim bConnected As Boolean
Dim X As Long
    bConnected = InternetConnected(lConnectionInfo, sConnectionName)
    If (lConnectionInfo And INTERNET_CONNECTION_MODEM) = INTERNET_CONNECTION_MODEM Then
        GetConName = sConnectionName
    Else
        GetConName = ""
    End If
End Function

Public Function IsLeapYear(Yr As Integer) As Boolean
'*************************************
'Function Name: InternetConnected
'Use: check if its a leepyear this year (366 days insted of 365)
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-??-??
'*************************************
    IsLeapYear = False
    If Yr Mod 4 = 0 Then
        IsLeapYear = True
        If Yr Mod 100 = 0 Then
            If (Yr Mod 400) Then IsLeapYear = False
        End If
    End If
End Function
