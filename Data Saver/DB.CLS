VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public DatabaseName As String
Public DataRS As Recordset
Public TheDatabase As Database

Public Function GetNr() As Long
    On Error Resume Next
    GetNr = DataRS.RecordCount
End Function

Public Sub LoadDB()
    If Mid(App.Path, Len(App.Path), 1) = "\" Then
        DatabaseName = App.Path & "Data.mdb"
    Else
        DatabaseName = App.Path & "\Data.mdb"
    End If
    'DatabaseName = "g:\mina program\visual basic\Data Saver\data.mdb"
    'Open the database
    Set TheDatabase = DBEngine.Workspaces(0).OpenDatabase(DatabaseName)
End Sub

'Hämta data till treeview
Public Function GetGame() As Variant
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT Title, ID FROM Game WHERE (Game.Title<>'') ORDER BY Game.Title")
    If Not DataRS.EOF Then
        GetGame = DataRS.GetRows(DataRS.RecordCount)
    End If
End Function

Public Function GetFilmName() As Variant
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT Title, ID From Film WHERE (Film.Title<>'') Order by Film.Title")
    If Not DataRS.EOF Then
        GetFilmName = DataRS.GetRows(DataRS.RecordCount)
    End If
End Function

Public Function GetMusicName() As Variant
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT Title, ID From Music WHERE (Music.Title<>'') ORDER BY Title")
    If DataRS.RecordCount > 0 Then
        DataRS.MoveFirst
        If Not DataRS.EOF Then
            GetMusicName = DataRS.GetRows(DataRS.RecordCount)
        End If
    End If
End Function


'Hämta all info när användaren clickar på en node

Public Function GetGameData(ByVal dID As Long) As Variant
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT Title, Type, Maker, Grade, Format, Media, Desc, ID From Game WHERE (Game.ID=" & dID & ")")
    GetGameData = DataRS.GetRows(1)
End Function

Public Function GetMusic(ByVal dID As Long) As Variant
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT ID, Title, Artist, Tracks From Music Where (Music.ID = " & dID & ")")
    GetMusic = DataRS.GetRows(1)
End Function

Public Function GetFilmData(ByVal lID As Long) As Variant
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT Title, ID, Grade, Media, Desc, Year from Film Where (Film.ID=" & lID & ")")
    GetFilmData = DataRS.GetRows(1)
End Function


'Spara data när användaren vill
Public Sub SaveGame(ByVal Title As String, ByVal dType As String, ByVal Maker As String, ByVal Grade As String, ByVal Format As String, ByVal Media As String, ByVal Desc As String, ByVal dID As Long)
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT Title, Type, Maker, Grade, Format, Media, Desc, ID From Game WHERE (Game.ID=" & dID & ")")

    DataRS.Edit
    DataRS!Title = "" & Title
    DataRS!Type = "" & dType
    DataRS!Maker = "" & Maker
    DataRS!Grade = "" & Grade
    DataRS!Format = "" & Format
    DataRS!Media = "" & Media
    DataRS!Desc = "" & Desc
    DataRS.Update
    Set DataRS = Nothing
End Sub

Public Sub SaveFilm(ByVal lID As Long, ByVal Title As String, ByVal Media As String, ByVal Desc As String, ByVal Year As String, ByVal Grade As String)
Dim X As Long
Dim Read As String
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT ID, Title, Media, Desc, Year, Grade From Film Where (Film.ID = " & lID & ")")
    DataRS.Edit

    DataRS!Title = "" & Title
    DataRS!Media = "" & Media
    DataRS!Desc = "" & Desc
    DataRS!Year = "" & Year
    DataRS!Grade = "" & Grade
    DataRS.Update
    Set DataRS = Nothing
End Sub

Public Sub SaveMusic(ByVal lID As Long, ByVal Title As String, ByVal Artist As String, Optional Tracks As String)
Dim X As Long
Dim Read As String
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT ID, Title, Artist, Tracks From Music Where (Music.ID = " & lID & ")")
    DataRS.Edit

    DataRS!Title = "" & Title
    DataRS!Artist = "" & Artist
    DataRS!Tracks = "" & Tracks
    DataRS.Update
    Set DataRS = Nothing
End Sub

'Skapa ett nytt record i databasen
Public Function NewRecord() As Long
    
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT ID, Title From Music WHERE (Music.ID=1)")
    DataRS.AddNew
    DataRS!Title = "MusicGarsVG"
    DataRS.Update
    Set DataRS = Nothing
    
    Set DataRS = TheDatabase.OpenRecordset("SELECT Title, ID From Music WHERE (Music.Title='MusicGarsVG')")
    NewRecord = DataRS!Id
    DataRS.Edit
    DataRS!Title = "New Record"
    DataRS.Update
    Set DataRS = Nothing

End Function

Public Function NewFilm() As Long
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT Title, Media, ID From Film WHERE (Film.ID=1)")

    DataRS.AddNew
    DataRS!Title = "FilmGarsVG"
    DataRS!Media = "VHS"
    DataRS.Update
    Set DataRS = Nothing

    Set DataRS = TheDatabase.OpenRecordset("SELECT Title, ID From Film WHERE (Film.Title='FilmGarsVG')")
    NewFilm = DataRS!Id
    DataRS.Edit
    DataRS!Title = "New Film"
    DataRS.Update
    Set DataRS = Nothing
End Function

Public Function NewGame() As Long
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT Title, Type, Grade, Format, Media, ID From Game WHERE (Game.ID=1)")

    DataRS.AddNew
    DataRS!Title = "GameGarsVG"
    DataRS!Type = "Driving"
    DataRS!Grade = 0
    DataRS!Format = "PC"
    DataRS!Media = "CD"
    DataRS.Update
    Set DataRS = Nothing
    
    Set DataRS = TheDatabase.OpenRecordset("SELECT Title, ID From Game WHERE (Game.Title='GameGarsVG')")
    NewGame = DataRS!Id
    DataRS.Edit
    DataRS!Title = "New Game"
    DataRS.Update
    Set DataRS = Nothing

End Function

'Deleta ett record
Public Sub DeleteGame(ByVal lID)
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT ID From Game WHERE (Game.ID=" & lID & ")")
    DataRS.Delete
    Set DataRS = Nothing
End Sub

Public Sub DeleteFilm(ByVal lID)
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT ID From Film Where (Film.ID=" & lID & ")")
    DataRS.Delete
    Set DataRS = Nothing
End Sub

Public Sub DeleteMusic(ByVal lID)
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT ID From Music WHERE (Music.ID=" & lID & ")")
    DataRS.Delete
    Set DataRS = Nothing
End Sub

Public Function GetAllFilm() As Variant
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("Select ID, Title, Media, Year, Grade from Film Where (Film.Title<>'') Order by Film.Title")
    If DataRS.RecordCount > 0 Then
        GetAllFilm = DataRS.GetRows(DataRS.RecordCount)
    End If
End Function

Public Function GetAllGames() As Variant
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT ID, Title, Type, Media, Format, Grade, Maker FROM Game WHERE (Game.Title<>'') ORDER BY Game.Title")
    If DataRS.RecordCount > 0 Then
        GetAllGames = DataRS.GetRows(DataRS.RecordCount)
    End If
End Function

Public Function GetAllMusic() As Variant
    Set DataRS = Nothing
    Set DataRS = TheDatabase.OpenRecordset("SELECT ID, Title, Artist, Tracks From Music Where (Music.Title<>'')")
    If DataRS.RecordCount > 0 Then
        GetAllMusic = DataRS.GetRows(DataRS.RecordCount)
    End If
End Function
