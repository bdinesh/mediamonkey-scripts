Sub OnStartup
 'Set UI = SDB.UI

' Add a couple of menu items here:

' Add a submenu to the View menu...
'Set Mnu = UI.AddMenuItemSub( UI.Menu_Tools, -1, 1)
'Mnu.Caption = "Clean Tags"
'Mnu.OnClickFunc = "CleanTags"

' ... and add Statistics item there
'Set Mnu = UI.AddMenuItem( Mnu, 0, 0)
'Mnu.Caption = "&Statistics"     
'Mnu.UseScript = Script.ScriptPath
'Mnu.OnClickFunc = "ShowIt"
'Mnu.Shortcut = "Ctrl+1"
'Mnu.IconIndex = 35
    
 
End Sub

Sub CleanTags()
Dim SongList, JunkText, Song, i, mb, Progress

JunkText = InputBox("Enter the text to be removed from the tags of the below tracks", "Clean Tags - Enter Text")

If JunkText="" Then
	Exit Sub
End If

 ' Get list of selected tracks from MediaMonkey
  Set SongList = SDB.SelectedSongList
  If SongList.count=0 Then
     Set SongList = SDB.AllVisibleSongList
  End If  

  'No songs selected?
  If SongList.count = 0 Then
     mb = MsgBox("No songs were selected. Please select some songs and try again",0,"Error")
     Exit Sub
  End If

  'Set up progress
  Set Progress = SDB.Progress
  Progress.Text = "Changing Tags..."
  Progress.MaxValue = SongList.count

  'Process all selected tracks
  For i=0 To SongList.count-1
    Set Song = SongList.Item(i)
    'correct the tags
	Song.AlbumArtistName = Replace(Song.AlbumArtistName,JunkText,"")
	Song.AlbumName = Replace(Song.AlbumName,JunkText,"")
	Song.ArtistName = Replace(Song.ArtistName,JunkText,"")
	Song.Author = Replace(Song.Author,JunkText,"")
	Song.Band = Replace(Song.Band,JunkText,"")
	Song.Comment = Replace(Song.Comment,JunkText,"")
	Song.Conductor = Replace(Song.Conductor,JunkText,"")
	Song.Copyright = Replace(Song.Copyright,JunkText,"")
	Song.Custom1 = Replace(Song.Custom1,JunkText,"")
	Song.Custom2 = Replace(Song.Custom2,JunkText,"")
	Song.Custom3 = Replace(Song.Custom3,JunkText,"")
	Song.Custom4 = Replace(Song.Custom4,JunkText,"")
	Song.Custom5 = Replace(Song.Custom5,JunkText,"")
	Song.DiscNumberStr = Replace(Song.DiscNumberStr,JunkText,"")
	Song.Encoder = Replace(Song.Encoder,JunkText,"")	
	Song.Genre = Replace(Song.Genre,JunkText,"")
	Song.Grouping = Replace(Song.Grouping,JunkText,"")
	Song.InvolvedPeople = Replace(Song.InvolvedPeople,JunkText,"")
	Song.ISRC = Replace(Song.ISRC,JunkText,"")
	Song.Lyricist = Replace(Song.Lyricist,JunkText,"")
	Song.Lyrics = Replace(Song.Lyrics,JunkText,"")
	Song.Mood = Replace(Song.Mood,JunkText,"")
	Song.MusicComposer = Replace(Song.MusicComposer,JunkText,"")
	Song.Occasion = Replace(Song.Occasion,JunkText,"")
	Song.OriginalArtist = Replace(Song.OriginalArtist,JunkText,"")
	Song.OriginalLyricist = Replace(Song.OriginalLyricist,JunkText,"")
	Song.OriginalTitle = Replace(Song.OriginalTitle,JunkText,"")
	Song.Publisher = Replace(Song.Publisher,JunkText,"")
	Song.Quality = Replace(Song.Quality,JunkText,"")
	Song.RatingString = Replace(Song.RatingString,JunkText,"")
	Song.Tempo = Replace(Song.Tempo,JunkText,"")
	Song.Title = Replace(Song.Title,JunkText,"")
	Song.TrackOrderStr = Replace(Song.TrackOrderStr,JunkText,"")
	Song.Path=Replace(Song.Path,JunkText,"")
	
	Song.WriteTags	
    Song.UpdateDB
    Progress.value = i+1
	
    If Progress.terminate Then
       Exit For
    End if   
  Next
 
  Set Progress =  nothing
End Sub


Sub ReplaceWithSemicolon()
Dim SongList, JunkText, Song, i, mb, Progress

JunkText = InputBox("Enter the text to be replaced with semicolon(;) from the tags of the below tracks", "Replace With Semicolon - Enter Text")

If JunkText="" Then
	Exit Sub
End If

Set SongList = SDB.SelectedSongList
  If SongList.count=0 Then
     Set SongList = SDB.AllVisibleSongList
  End If  

  'No songs selected?
  If SongList.count = 0 Then
     mb = MsgBox("No songs were selected. Please select some songs and try again",0,"Error")
     Exit Sub
  End If
  
  'Set up progress
  Set Progress = SDB.Progress
  Progress.Text = "Replacing Tags with Semicolon..."
  Progress.MaxValue = SongList.count

  'Process all selected tracks
  For i=0 To SongList.count-1
    Set Song = SongList.Item(i)
    'correct the tags
	Song.AlbumArtistName = Replace(Song.AlbumArtistName,JunkText,";")
	Song.ArtistName = Replace(Song.ArtistName,JunkText,";")
	Song.Author = Replace(Song.Author,JunkText,";")
	Song.Band = Replace(Song.Band,JunkText,";")
	Song.Comment = Replace(Song.Comment,JunkText,";")
	Song.Conductor = Replace(Song.Conductor,JunkText,";")
	Song.Encoder = Replace(Song.Encoder,JunkText,";")	
	Song.Genre = Replace(Song.Genre,JunkText,";")
	Song.Grouping = Replace(Song.Grouping,JunkText,";")
	Song.InvolvedPeople = Replace(Song.InvolvedPeople,JunkText,";")
	Song.Lyricist = Replace(Song.Lyricist,JunkText,";")
	Song.Mood = Replace(Song.Mood,JunkText,";")
	Song.MusicComposer = Replace(Song.MusicComposer,JunkText,";")
	Song.Occasion = Replace(Song.Occasion,JunkText,";")
	Song.OriginalArtist = Replace(Song.OriginalArtist,JunkText,";")
	Song.OriginalLyricist = Replace(Song.OriginalLyricist,JunkText,";")
	Song.Publisher = Replace(Song.Publisher,JunkText,";")
	
	Song.WriteTags	
    Song.UpdateDB
    Progress.value = i+1
	
    If Progress.terminate Then
       Exit For
    End if   
  Next
 
  Set Progress =  nothing
End Sub