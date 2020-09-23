<div align="center">

## A Tutorial for Writing to / READING FROM An ASX playlist \(the ones that Windows Media Player uses\)


</div>

### Description

You can use the following code to read from, and write to ASX playlists, the type that can be read and created by Windows Media Player,

and generally most other music programs. This may not be the easyest code in the world to understand, so i would recommend you have a good knowlage of string manipulation before reading this (see my other uploads)

For the following code you work correctly, YOU MUST HAVE, ON THE FORM:

2 x Listboxes, name List1 and List2

1 x Textboxes, name Text1

tHe_cLeanER
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jon Barker](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jon-barker.md)
**Level**          |Intermediate
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Sound/MP3](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/sound-mp3__1-45.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jon-barker-a-tutorial-for-writing-to-reading-from-an-asx-playlist-the-ones-that-windows-me__1-29582/archive/master.zip)





### Source Code

```
'For the following code you work correctly, YOU MUST HAVE, ON THE FORM:
'  2 x Listboxes, name List1 and List2
'  1 x Textboxes, name Text1
'--------------------------------------------------------------------
'
'To use the code, you call it similar to calling another other Sub or Function;
'
'(anywhere in your code)
'Call WriteToASX("C:\windows\newasx.asx") ' THIS WOULD WRITE TO THE FILENAME SPECIFIED
'
'or:
'
'(anywhere in your code)
'Call PullFromASX("C:\windows\newasx.asx") ' THIS WOULD READ FROM THE ASX FILE SPECIFIED
'
'------------------------ Creating an ASX file -----------------------
Sub WriteToASX(Path As String)
On Local Error GoTo WriteError ' ERROR HANDLING
Open Path For Output As #1 ' OPEN A NEW ASX FILE FOR WRITING TO
  Print #1, "<ASX Version = " & Chr(34) & "3.0" & Chr(34) & " >" 'THE FOLLOWING CODE WRITE THE GENERAL HEADER INFO, SO THAT MEDIA PLAYER ACCEPTS IT
  Print #1, ""
  Print #1, "<Param Name = " & Chr(34) & "Name" & Chr(34) & " Value = " & Chr(34) & "Playlist - ASX Format - jBistoGOOD@Hotmail.com" & Chr(34) & " />" 'SET THE FORMAT... AND CREDITS :)
  Print #1, "<Param Name = " & Chr(34) & "AllowShuffle" & Chr(34) & " Value = " & Chr(34) & "yes" & Chr(34) & " />" 'INFO ABOUT THE PLAYLIST (REQUIRED)
  Print #1, ""
  For i = 0 To List1.ListCount - 1 'LOOP WRITING THE FILENAMES UNTIL ALL ARE DONE
  Print #1, "<Entry>"
    Print #1, "<Param Name = " & Chr(34) & "name" & Chr(34) & " Value = " & Chr(34) & List2.List(i) & Chr(34) & " />" 'ONLY THE NAME OF THE SONG ETC, NOT PATH OR FILENAME
    Print #1, "<ref href = " & Chr(34) & List1.List(i) & Chr(34) & " />" 'WRITES THE FULL PATH, AND FILENAME THAT ARE READ TO PLAY THE SONG
    Print #1, "</Entry>"
    Print #1, ""
  Next i
  Print #1, "</asx>" ' CLOSE THE ASX FILE
Close #1 'PHYSICALLY CLOSE THE OPEN FILE
Exit Sub
WriteError: ' AN ERROR HAS OCCURED; TELL THE USER, AND SAY WHAT THE ERROR IS
  MsgBox "There was an error writing to the playlist: " & Err.Description
End Sub
'---------------------------------------------------------------------
'
'This above code will write an ASX file that can be read by Windows Media Player. It uses List1 for the complete filenames for files you wish to be
'in the playlist, and reads from List2 the song names, or whatever you wish the filenames to be identified by.
'
'------------------------ Reading from a created ASX -----------------
Sub PullFromASX(Path As String)
Dim SStart, Getname As Long
Dim Part2 As Long             ' DECLARE THE VARIBLES
Dim FullFilename, CosmeticName As String
On Local Error GoTo errorasx 'BASIC ERROR HANDLING
SStart = 1  'WHERE TO START LOOKING FOR THE SONG
Getname = 1  'A TEMPORY INTEGER VALUE CONTAINING THE START OF THE SONG FILENAME
Open Path For Input As #1
  Text1.Text = Input(LOF(1), 1)  ' OPEN THE EXISTING ASX FILE, AND EXTRACT THE CONTENTS TO TEXT1
Close #1
Do Until Getname = "0"
  Getname = InStr(SStart, Text1.Text, "<ref href = " & Chr(34)) ' FIND THE START OF THE SONG FILENAME
  Part2 = InStr(Getname + 13, Text1.Text, Chr(34)) ' FIND THE END OF THE SONG FILENAME
  FullFilename = Mid(Text1.Text, Getname + 13, Part2 - Getname - 13) ' FIND THE FILE NAME, RELATIVE TO FIRST AND LAST PARTS
  If Getname <> 0 Then List1.AddItem FullFilename ' ADD THE FILENAME TO LIST1
  SStart = Getname + 1 ' SPECIFIES WHERE TO SEARCH FROM (AND UPDATES IT)
Loop ' RESTART THE LOOP
MsgBox "All filenames extracted from the ASX file. Continuing to extract song names."
Call PullCosmeticsASX(Path) ' START THE NEXT STEP
Exit Sub
errorasx: ' AN ERROR HAS OCCURED; TELL THE USER, AND SAY WHAT THE ERROR IS
  MsgBox "There was an error reading from the playlist: " & Err.Description
End Sub
Sub PullCosmeticsASX(Path As String)
Dim SStart, Getname As Long
Dim Part2 As Long             ' DECLARE THE VARIBLES
Dim FullFilename, CosmeticName As String
On Local Error GoTo errorasx 'BASIC ERROR HANDLING
SStart = 1
Getname = 1
Open Path For Input As #1
  Text1.Text = Input(LOF(1), 1)  ' OPEN THE EXISTING ASX FILE, AND EXTRACT THE CONTENTS TO TEXT1
Close #1
Do Until Getname = 0 ' START THE LOOP, LOOKING FOR SONG NAMES
  Getname = InStr(SStart, Text1.Text, "<Param Name = " & Chr(34) & "name" & Chr(34) & " Value = " & Chr(34)) ' FIND THE START OF THE NAME
  Part2 = InStr(Getname + 30, Text1.Text, Chr(34)) ' FIND THE END OF THE NAME
  CosmeticName = Mid(Text1.Text, Getname + 30, Part2 - Getname - 30) ' FIND THE SONG NAME, RELATIVE TO FIRST AND LAST PARTS
  If Getname <> 0 Then List2.AddItem CosmeticName ' ADD THE SONG NAME TO LIST2
  SStart = Getname + 1 ' SPECIFIES WHERE TO SEARCH FROM (AND UPDATES IT)
Loop ' RESTART THE LOOP
MsgBox "Completed reading from playlist."  ' IF NO MORE ENTERIES ARE FOUND THEN YOU ARE FINISHED; EXIT SUB
Exit Sub
errorasx: ' AN ERROR HAS OCCURED; TELL THE USER, AND SAY WHAT THE ERROR IS
  MsgBox "There was an error reading from the playlist: " & Err.Description
End Sub
'---------------------------------------------------------------------------------
'
'The above code will do two things:
'  1. Take all the full filenames of songs etc in the ASX file, and add them to List1
'  2. Take just the filename, or name of song, or whatever is specified in the ASX, and put it into List2
'Text1 just holds the ASX file to read from it, and is only used when directly reading from the ASX file.
'
'---------------------------------------------------------------------------------
Hope this helps you to understand the making of the playlist, and the various string manipulation commands.
If you need a detailed explanation of the InStr command etc, then check out my other tutorials
(click 'other submissions from this person' below)
keep coding!
tHe_cLeanER
jBistoGOOD@Hotmail.com
```

