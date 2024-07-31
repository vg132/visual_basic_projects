Attribute VB_Name = "Module1"
Public oDB As New DB
Public vArray As Variant
Public X As Long
Public Y As Long
Public Responce

Public Sub InitData()
    Set oDB = New DB
    oDB.LoadDB
End Sub
