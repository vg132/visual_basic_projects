Attribute VB_Name = "modDB"
Option Explicit

Private DBName As String
Public RS As Recordset
Private db As Database

Public Enum GetPriceData
    All = 0
    One = 1
End Enum

Public Sub OpenDataBase()
'*************************************
'Function Name: OpenDataBase
'Use: Open the database
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-18
'*************************************
    If Right(App.Path, 1) = "\" Then
        DBName = "C:\My Documents\Mina Program\Visual Basic\VB Code\Code.mdb"
    Else
        DBName = "C:\My Documents\Mina Program\Visual Basic\VB Code\Code.mdb"
    End If
    Set db = DBEngine.Workspaces(0).OpenDataBase(DBName)
End Sub
Public Sub CloseDataBase()
    db.Close
End Sub

Public Function GetMain() As Variant
'*************************************
'Function Name: GetMain
'Use:
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-10-22
'*************************************
    Set RS = Nothing
    GetDB "SELECT * From Main"
    If RS.EOF = False Then
        GetMain = RS.GetRows(RS.RecordCount + 1)
    End If
End Function

Private Function GetDB(CNString As String)
'*************************************
'Function Name: GetDB
'Use:
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-21
'*************************************
    Set RS = Nothing
    Set RS = db.OpenRecordset(CNString)
    If Not RS.EOF Then
        RS.MoveLast
        RS.MoveFirst
    End If
End Function

Public Sub AddMainItem(sItemName As String)
'*************************************
'Function Name: AddMainItem
'Use:
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-10-22
'*************************************
    GetDB "SELECT * From Main Where Name='" & sItemName & "'"
    If RS.EOF Then
        RS.AddNew
        RS!Name = sItemName
        RS.Update
        GetDB "SELECT ID From Main Where Name='" & sItemName & "'"
        AddNewTabell RS!Id
    Else
        MsgBox "Error"
    End If
End Sub

Public Sub AddNewTabell(sName As String)
'*************************************
'Function Name: AddNewTabell
'Use:
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-10-23
'*************************************
Dim AuTd As TableDef
Dim AuFlds(2) As Field
Dim AuIdx As Index
Dim AuIdxFld As Field
    ' Create new TableDef for Authors table.
    Set AuTd = db.CreateTableDef(sName)
    ' Add fields to MyTableDef.
    Set AuFlds(0) = AuTd.CreateField("ID", dbLong)
    ' Make it a counter field.
    AuFlds(0).Attributes = dbAutoIncrField
    Set AuFlds(1) = AuTd.CreateField("Title", dbText)
    AuFlds(1).Size = 255
    Set AuFlds(2) = AuTd.CreateField("Tip", dbMemo)
    AuTd.Fields.Append AuFlds(0)
    AuTd.Fields.Append AuFlds(1)
    AuTd.Fields.Append AuFlds(2)
    ' Now add an Index.
    Set AuIdx = AuTd.CreateIndex("ID")
    AuIdx.Primary = True
    AuIdx.Unique = True
    Set AuIdxFld = AuIdx.CreateField("ID")
    ' Append Field to Fields collection of Index object.
    AuIdx.Fields.Append AuIdxFld
    ' Append Index to Indexes collection.
    AuTd.Indexes.Append AuIdx
    ' Append TableDef to TableDefs collection.
    db.TableDefs.Append AuTd
End Sub

Public Sub DeleteTabell(sName As String)
'*************************************
'Function Name: DeleteTabell
'Use: Delete a tabell from the database
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-10-23
'*************************************
    GetDB "SELECT * From Main Where ID=" & sName
    If Not RS.EOF Then
        RS.Delete
    End If
    Set RS = Nothing
    sName = ConvertName(sName)
    db.TableDefs.Delete sName
End Sub

Public Function GetTips(ByVal sName As String) As Variant
'*************************************
'Function Name: GetTips
'Use: Load tips from db
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-10-23
'*************************************
    sName = ConvertName(sName)
    GetDB "SELECT Title, ID From " & sName & " Order by Title"
    If Not RS.EOF Then
        GetTips = RS.GetRows(RS.RecordCount + 1)
    End If
End Function

Private Function ConvertName(ByVal sName As String) As String
Dim sTemp As String
Dim X As Long
    sTemp = ""
    For X = 1 To Len(sName)
        If Mid(sName, X, 1) = " " Then
            sTemp = sTemp & "_"
        Else
            sTemp = sTemp & Mid(sName, X, 1)
        End If
    Next
    ConvertName = sTemp
End Function

Public Sub NewItem(sName As String, sTip As String, sDB As String)
    sDB = ConvertName(sDB)
    GetDB "SELECT * From " & sDB
    RS.AddNew
    RS!Title = sName
    RS!Tip = sTip
    RS.Update
End Sub

Public Sub DeleteItem(ByVal lId As Long, ByVal lDB As Long)
    GetDB "SELECT * From " & lDB & " Where ID=" & lId
    If Not RS.EOF Then
        RS.Delete
    End If
End Sub

Public Function GetTip(ByVal lId As Long, ByVal lDB As Long) As String
    GetDB "SELECT Tip From " & lDB & " Where ID=" & lId
    If Not RS.EOF Then
        GetTip = "" & RS!Tip
    End If
End Function

Public Sub SaveTip(ByVal lId As Long, ByVal lDB As Long, sTip As String)
    GetDB "SELECT * From " & lDB & " Where ID=" & lId
    If Not RS.EOF Then
        RS.Edit
        RS!Tip = sTip
        RS.Update
    End If
End Sub

Public Function Find(ByVal sFind As String, ByVal lDB As Long) As Variant
    GetDB "SELECT * From " & lDB & " Where Tip Like '*" & sFind & "*'"
    If Not RS.EOF Then
        Find = RS.GetRows(RS.RecordCount + 1)
    End If
End Function
