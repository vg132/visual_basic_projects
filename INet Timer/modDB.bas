Attribute VB_Name = "modDB"
Option Explicit
Private DatabaseName As String
Public DataRS As Recordset
Private Cn As String
Private TheDatabase As Database

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
        DatabaseName = "C:\My Documents\Mina Program\Visual Basic\INet Timer\INetCounter2.mdb"
        'DatabaseName = App.Path & "INetCounter.mdb"
    Else
        DatabaseName = "C:\My Documents\Mina Program\Visual Basic\INet Timer\INetCounter2.mdb"
        'DatabaseName = App.Path & "\" & "INetCounter.mdb"
    End If
    'Open the database
    Set TheDatabase = DBEngine.Workspaces(0).OpenDataBase(DatabaseName)
End Sub

Public Sub CloseDataBase()
    TheDatabase.Close
End Sub

Public Function GetPrice(GetData As GetPriceData, Optional sName As String) As Variant
'*************************************
'Function Name: GetPrice
'Use: Get the price
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-18
'*************************************
    Set DataRS = Nothing
    If GetData = All Then
        GetDB "SELECT * From One"
        If Not DataRS.EOF Then
            DataRS.MoveLast
            DataRS.MoveFirst
            GetPrice = DataRS.GetRows(DataRS.RecordCount + 1)
        End If
    ElseIf GetData = One Then
        GetDB "SELECT * From One Where Name='" & sName & "'"
        If Not DataRS.EOF Then
            GetPrice = DataRS.GetRows(1)
        End If
    End If
End Function

Public Sub SavePrice(sName As String, dPrice As Double)
'*************************************
'Function Name: SavePrice
'Use: Save a new price
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-22
'*************************************
Dim X As Long
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT * From One Where Name='" & sName & "'")
    If Not DataRS.EOF Then
        DataRS.Edit
        DataRS!Name = sName
        DataRS!Price = dPrice
        DataRS.Update
    Else
        Set DataRS = Nothing
        Set DataRS = TheDatabase.OpenRecordset("SELECT Max(Id) As EndID From One")
        If DataRS!EndId <> vbNull Then
            X = DataRS!EndId
        Else
            X = 0
        End If
        Set DataRS = TheDatabase.OpenRecordset("SELECT * From One Where ID=" & X)
        DataRS.AddNew
        DataRS!Name = sName
        DataRS!Price = dPrice
        DataRS.Update
    End If
End Sub

Public Sub DeletePrice(sName As String)
'*************************************
'Function Name: DeletePrice
'Use: Delete a price from the DB
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-24
'*************************************
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT * From One Where Name='" & sName & "'")
    If Not DataRS.EOF Then
        DataRS.Delete
    End If
End Sub

Public Sub SaveTariff(sName As String, iNr As Integer)
'*************************************
'Function Name: SaveTariff
'Use:
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-21
'*************************************
Dim lID As Long
Dim X As Integer
    With frmTariff
        Set DataRS = Nothing
        Set DataRS = TheDatabase.OpenRecordset("SELECT * From One Where Name='" & sName & "'")
        If Not DataRS.EOF Then
            lID = DataRS!id
            DataRS.Delete
            X = 0
            Do Until X > 10
                GetDB "SELECT * From Two Where Name_ID=" & lID
                If DataRS.RecordCount = 0 Then Exit Do
                DataRS.Delete
            Loop
        End If
        GetDB "SELECT * From One"
        DataRS.AddNew
        DataRS!Name = sName
        DataRS.Update
        GetDB "SELECT * From One Where Name='" & sName & "'"
        lID = DataRS!id
        GetDB "SELECT * From Two"
        For X = 0 To iNr
            DataRS.AddNew
            DataRS!Price = "" & .txtPrice(X).Text
            DataRS!StartTime = .txtTime(X).Text
            If .cboTime(X).Text = "24:00:00" Then
                DataRS!EndTime = "23:59:59"
            Else
                DataRS!EndTime = .cboTime(X).Text
            End If
            DataRS!Nr = X
            DataRS!Name_ID = lID
            DataRS.Update
        Next
    End With
End Sub

Public Function GetTariff(sName As String) As Variant
'*************************************
'Function Name: GetTariff
'Use: Load a tariff
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-21
'*************************************
Dim lID As Long
    GetDB "SELECT * From One Where Name='" & sName & "'"
    If Not DataRS.EOF Then
        lID = DataRS!id
        GetDB "SELECT * From Two Where Name_ID=" & lID
        If Not DataRS.EOF Then
            GetTariff = DataRS.GetRows(DataRS.RecordCount + 1)
        End If
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
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset(CNString)
    If Not DataRS.EOF Then
        DataRS.MoveLast
        DataRS.MoveFirst
    End If
End Function

Public Sub SaveWeek(ByVal sName As String)
'*************************************
'Function Name: SaveWeek
'Use: Save the week settings into the database
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-21
'*************************************
    With frmSetTariff
        GetDB "SELECT * From Four Where Name='" & sName & "'"
        If DataRS.EOF Then
            DataRS.AddNew
            DataRS!Day1 = .cboUse(0).Text
            DataRS!Day2 = .cboUse(1).Text
            DataRS!Day3 = .cboUse(2).Text
            DataRS!Day4 = .cboUse(3).Text
            DataRS!Day5 = .cboUse(4).Text
            DataRS!Day6 = .cboUse(5).Text
            DataRS!Day7 = .cboUse(6).Text
            DataRS!Holiday = .cboUse(7).Text
            If .txtCon.Text <> "" Then DataRS!LogOn = .txtCon.Text
            DataRS!Name = sName
            DataRS.Update
        Else
            DataRS.Edit
            DataRS!Day1 = .cboUse(0).Text
            DataRS!Day2 = .cboUse(1).Text
            DataRS!Day3 = .cboUse(2).Text
            DataRS!Day4 = .cboUse(3).Text
            DataRS!Day5 = .cboUse(4).Text
            DataRS!Day6 = .cboUse(5).Text
            DataRS!Day7 = .cboUse(6).Text
            DataRS!Holiday = .cboUse(7).Text
            DataRS!LogOn = .txtCon.Text
            DataRS!Name = sName
            DataRS.Update
        End If
    End With
End Sub

Public Function GetWeek(ByVal sName As String) As String
'*************************************
'Function Name: GetWeek
'Use: Load the week from the database
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-21
'*************************************
    With frmSetTariff
        GetDB "SELECT * From Four Where Name='" & sName & "'"
        If Not DataRS.EOF Then
            .cboUse(0).Text = DataRS!Day1
            .cboUse(1).Text = DataRS!Day2
            .cboUse(2).Text = DataRS!Day3
            .cboUse(3).Text = DataRS!Day4
            .cboUse(4).Text = DataRS!Day5
            .cboUse(5).Text = DataRS!Day6
            .cboUse(6).Text = DataRS!Day7
            .cboUse(7).Text = DataRS!Holiday
            .txtCon.Text = DataRS!LogOn
        Else
            GetWeek = "No"
        End If
    End With
End Function

Public Function GetCurrentTariff(sDay As String, sName As String) As Variant
'*************************************
'Function Name: GetCurrentPrice
'Use: Get Current tariff
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-26
'*************************************
Dim lID As Long
    GetDB "SELECT " & sDay & " From Four Where Name='" & sName & "'"
    Select Case LCase(sDay)
    Case "day1"
        sDay = DataRS!Day1
    Case "day2"
        sDay = DataRS!Day2
    Case "day3"
        sDay = DataRS!Day3
    Case "day4"
        sDay = DataRS!Day4
    Case "day5"
        sDay = DataRS!Day5
    Case "day6"
        sDay = DataRS!Day6
    Case "day7"
        sDay = DataRS!Day7
    Case "holiday"
        sDay = DataRS!Holiday
    End Select
    GetDB "SELECT * From One Where Name='" & sDay & "'"
    lID = DataRS!id
    GetDB "SELECT * From Two Where Name_ID=" & lID
    If Not DataRS.EOF Then
        DataRS.MoveLast
        DataRS.MoveFirst
        GetCurrentTariff = DataRS.GetRows(DataRS.RecordCount + 1)
    End If
End Function

Public Function AddUser(sUser As String, sPassword As String) As Boolean
'*************************************
'Function Name: AddUser
'Use: Add a user to the database
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-28
'*************************************
    GetDB "SELECT * From Three Where Name='" & sUser & "'"
    If Not DataRS.EOF Then
        AddUser = False
    Else
        DataRS.AddNew
        DataRS!Name = sUser
        If sPassword <> "" Then DataRS!Password = sPassword
        AddUser = True
        DataRS.Update
    End If
    AddNewTabell sUser
End Function

Public Sub AddNewTabell(sName As String)
'*************************************
'Function Name: AddNewTabell
'Use: Add a new Tabell to the Database
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-10-23
'*************************************
Dim AuTd As TableDef
Dim AuFlds(5) As Field
Dim AuIdx As Index
Dim AuIdxFld As Field
    Set AuTd = TheDatabase.CreateTableDef(sName)
    Set AuFlds(0) = AuTd.CreateField("ID", dbLong)
    AuFlds(0).Attributes = dbAutoIncrField
    Set AuFlds(1) = AuTd.CreateField("OnTime", dbText)
    AuFlds(1).Size = 255
    Set AuFlds(2) = AuTd.CreateField("OffTime", dbText)
    AuFlds(2).Size = 255
    Set AuFlds(3) = AuTd.CreateField("Conection", dbText)
    AuFlds(3).Size = 255
    Set AuFlds(4) = AuTd.CreateField("Price", dbText)
    AuFlds(4).Size = 255
    AuTd.Fields.Append AuFlds(0)
    AuTd.Fields.Append AuFlds(1)
    AuTd.Fields.Append AuFlds(2)
    AuTd.Fields.Append AuFlds(3)
    AuTd.Fields.Append AuFlds(4)
    Set AuIdx = AuTd.CreateIndex("ID")
    AuIdx.Primary = True
    AuIdx.Unique = True
    Set AuIdxFld = AuIdx.CreateField("ID")
    AuIdx.Fields.Append AuIdxFld
    AuTd.Indexes.Append AuIdx
    TheDatabase.TableDefs.Append AuTd
End Sub

Public Function GetUser(gType As GetPriceData, Optional ByVal sUser, Optional ByVal lID As Long) As Variant
'*************************************
'Function Name: GetUser
'Use: get user name
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-28
'*************************************
    If gType = All Then
        GetDB "SELECT * From Three"
    ElseIf gType = One Then
        If IsMissing(sUser) Then
            GetDB "SELECT * From Three Where ID=" & lID
        Else
            GetDB "SELECT * From Three Where Name='" & sUser & "'"
        End If
    End If
    If Not DataRS.EOF Then
        GetUser = DataRS.GetRows(DataRS.RecordCount + 1)
    End If
End Function

Public Sub DeleteUser(ByVal lID As Long, ByVal sName As String)
'*************************************
'Function Name: DeleteUser
'Use: Delete a user
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-30
'*************************************
    GetDB "SELECT * From Three Where ID=" & lID
    If Not DataRS.EOF Then
        DataRS.Delete
        TheDatabase.TableDefs.Delete sName
    End If
End Sub

Public Sub SaveLog(ByVal sUser As String)
'*************************************
'Function Name: SaveLog
'Use: Save new data to log
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-10-05
'*************************************
Dim tmpVal As Long
    GetDB "SELECT * From " & sUser
    DataRS.AddNew
    DataRS!OnTime = INetData.OnTime
    DataRS!OffTime = Now
    DataRS!Price = INetData.Price
    DataRS!Conection = frmMain.lblConName.Caption
    DataRS.Update
    Set DataRS = Nothing
End Sub

Public Function GetLog(ByVal sUser As String, Optional FromDate As Date, Optional ToDate As Date) As Variant
'*************************************
'Function Name: GetLog
'Use: load the log for a user
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-10-05
'*************************************
    If FromDate <> 0 Then
        GetDB "SELECT * From " & sUser & " WHERE OnTime Between #" & FromDate & "# And #" & ToDate + 1 & "#;"
    Else
        GetDB "SELECT * From " & sUser
    End If
    If Not DataRS.EOF Then
        GetLog = DataRS.GetRows(DataRS.RecordCount + 1)
    End If
End Function

Public Function GetConFee(sName As String) As Double
'*************************************
'Function Name: GetConFee
'Use: Get the conection fee
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-10-05
'*************************************
    GetDB "SELECT LogOn From Four Where Name='" & sName & "'"
    GetConFee = DataRS!LogOn
End Function

Public Sub DeleteTariff(ByVal sName As String)
Dim lID As Long
Dim X As Long
Dim Rec As Long
    GetDB "SELECT * From One Where Name='" & sName & "'"
    If Not DataRS.EOF Then
        lID = DataRS!id
        DataRS.Delete
        GetDB "SELECT * From Two Where Name_ID=" & lID
        If Not DataRS.EOF Then
            Rec = DataRS.RecordCount - 1
            For X = 0 To Rec
                DataRS.Delete
                DataRS.MoveFirst
            Next
        End If
    End If
End Sub

Public Sub DeleteLogItem(ByVal sName As String, ByVal id As Long)
    GetDB "SELECT * From " & sName & " Where ID=" & id
    If Not DataRS.EOF Then
        DataRS.Delete
    End If
End Sub
