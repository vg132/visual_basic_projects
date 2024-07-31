Attribute VB_Name = "modBlock"
Private blockIndex(1 To 250) As Integer
Private i As Integer

Public Function getBlockIndex() As Integer
    'find first free index
    For i = 1 To 250
        If blockIndex(i) = 0 Then Exit For
    Next
    If i < 251 Then
        'set index to ocupied
        blockIndex(i) = 1
        'return free index
        getBlockIndex = i
    End If
End Function

Public Sub setFreeBlockIndex(index As Integer)
    blockIndex(index) = 0
End Sub

Public Sub resetBlockIndex()
    'Set all index elemets to 0 (0=free)
    For i = 1 To 250
        blockIndex(i) = 0
    Next
End Sub
