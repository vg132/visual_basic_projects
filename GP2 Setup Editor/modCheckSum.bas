Attribute VB_Name = "modCheckSum"
Public Function CheckSum(ByVal Path As String) As String
Dim FileNum As Integer
Dim Read As String
Dim Read2 As String
Dim Check As String
Dim i As Long
Dim X As Long
Dim Count1 As Long
Dim Count2 As Long
Dim Bin As String
Dim O As Long
Dim Tin As Long
Dim File As String
Dim Check2 As Long
Dim Test As Long
Const MaxPower = 16
    FileNum = FreeFile
    Open Path For Binary As FileNum
    Count2 = FileLen(Path) - 5
    File = String(Count2, " ")
    Get #FileNum, 1, File
    Close FileNum

    Read2 = ""
    Check = ""

    Check = 0
    
    Check2 = 0
    For Count1 = 1 To Count2
        Read = Mid(File, Count1, 1)
        Check2 = Check2 + Asc(Read)
        If Check2 > 65535 Then
            Check2 = 0
        End If
        Read = Asc(Read) + Check
        If Read > 65536 Then
            Read = Read - 65536
        End If
        Bin = ""  'Build the desired binary number in this string, bin.
        X = Val(Read) 'Convert decimal string in text1 to long integer
        For i = MaxPower To 0 Step -1
            If X And (2 ^ i) Then   ' Use the logical "AND" operator.
                Bin = Bin + "1"
            Else
                Bin = Bin + "0"
            End If
        Next
        Bin = Mid(Bin, 5, Len(Bin) - 4) & Mid(Bin, 2, 3)
        O = 0
        For Tin = 16 To 1 Step -1
            If Mid(Bin, Tin, 1) = "1" Then
                O = O + 2 ^ (16 - Tin)
            End If
        Next Tin
        Check = O
    Next
'    Read = Hex(Check)
'    Read = Mid(Read, 3) & Mid(Read, 1, 2)
    CheckSum = Check
End Function

