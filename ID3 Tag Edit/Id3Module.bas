Attribute VB_Name = "Id3Module"
Option Explicit

Public Type Id3
    Title As String * 30
    Artist As String * 30
    Album As String * 30
    sYear  As String * 4
    Comments As String * 30
    Genre As Byte
End Type

Public bInfo As Boolean
Public id3Info As Id3
Public GenreArray() As String

Public Const sGenreMatrix = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" & _
    "Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" & _
    "Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" & _
    "Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" & _
    "House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Space|Meditative|" & _
    "Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" & _
    "Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" & _
    "Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" & _
    "Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" & _
    "Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" & _
    "Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" & _
    "Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" & _
    "Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" & _
    "Ballad|power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Capella|Euro-House|" & _
    "Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|indie|Brit Pop|Negerpunk|Polsk Punk|" & _
    "Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" & _
    "Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"

Public Function GetId3(FileName As String)
Dim TaG As String * 3
    Open FileName For Binary As #1
    Get #1, FileLen(FileName) - 127, TaG
    If TaG = "TAG" Then
        Get #1, FileLen(FileName) - 124, id3Info
        bInfo = True
    Else
        With id3Info
            .Album = ""
            .Artist = ""
            .Comments = ""
            .Genre = 0
            .sYear = ""
            .Title = ""
            bInfo = False
        End With
    End If
    Close #1
End Function

Public Function SaveId3(FileName As String, Mp3Info As Id3)
Dim TaG As String * 3
    Open FileName For Binary As #1
    Get #1, FileLen(FileName) - 127, TaG
    If TaG = "TAG" Then
        Put #1, FileLen(FileName) - 124, Mp3Info
    Else
        Put #1, FileLen(FileName) - 127, "TAG"
        Close #1
        Call SaveId3(FileName, Mp3Info)
    End If
    Close #1
End Function
