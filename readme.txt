This is other tip from Claudio Heidel

With the number of Web sites using MP3 sound files (.mp3 files) , 
as a Web developer it may come in handy to have an ASP page that 
could easily read various tidbits of information about a Sound file. 
Fortunately, MP3 files save a plethora of useful information in their headers. 
This information can be read using the FileSystemObject through an ASP page! 

In order to ease the process of reading this information, 
I created a handy class to do the grunt work for me. 
(For more information on using classes with VBScript, be sure to read: 
Using Classes within VBScript.) The class 
I created returns the following information about a MP3 file: 


File              = FileName (Put & Get) 
Size              = FileSize in bytes
AudioTime	  = Expresed in Seconds
AudioVersion      = (1 or 2)
LayerDescription  = Layer I, Layer II or Layer III
CRC               = True / False
BitRate           = Bit Rate expresed in Kbps
SampleRate        = Expresed in Hz (Hertz)
Pading            = Tre/ False
Frames            = Number of Frames
FrameLenght       = Bytes
IsPrivate         = True / False
ChannelMode       = stereo , joint stereo . dual stereo or mono
CopyRight         = True / False
Original          = True / False
Emphasis          = none , 50/15 ms , reserved or CCIT J.17
SongName          = Text (30)
Artist            = Text (30)
Album             = Text (30)
AudioYear         = Text (4)
Genre             = Text (nChars)
Comment           = Text (30)


The use of the MP#Dump class is very simple. 
First off, I recommend that you place it in an include file so 
that you can easily include the class in the files that need 
to read Flash file information. (For more information on server-side includes 
be sure to read: The Low-Down on #include.) 
I chose to name my include file mp3_class.inc. 

To use the class, simply 
put this files join mp3 files in any directory and 
use code like the following: 

<%
'-------------------------------------------------------------
'  Create Date : 24/10/2001 (dd/mm/yyyy)
'  Mod. Date   : 25/10/2001
'  Author      : Claudio Heidel (heidel@f256.com)
'-------------------------------------------------------------

	Dim folderspec
	Dim fso
	Dim f
	Dim f1
	Dim fc
	Dim s

	set myObj  = new mp3dump
	folderspec = Server.MapPath(".")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFolder(folderspec)
	Set fc = f.Files
	For Each f1 in fc
		If LCase(Mid(StrReverse(f1.name), 1,3)) = "3pm" Then
			myObj.File = folderspec & "\" & f1.name

			Response.Write "FileName          = " & myObj.File              & "<br>"
			Response.Write "Size              = " & myObj.Size              & "<br>"
			Response.Write "AudioVersion      = " & myObj.AudioVersion      & "<br>"
			Response.Write "LayerDescription  = " & myObj.LayerDescription  & "<br>"
			Response.Write "CRC               = " & myObj.CRC               & "<br>"
			Response.Write "BitRate           = " & myObj.BitRate           & "<br>"
			Response.Write "SampleRate        = " & myObj.SampleRate        & "<br>"
			Response.Write "Pading            = " & myObj.Pading            & "<br>"
			Response.Write "Frames            = " & myObj.Frames            & "<br>"
			Response.Write "FrameLenght       = " & myObj.FrameLenght       & "<br>"
			Response.Write "IsPrivate         = " & myObj.IsPrivate         & "<br>"
			Response.Write "ChannelMode       = " & myObj.ChannelMode       & "<br>"
			Response.Write "CopyRight         = " & myObj.CopyRight         & "<br>"
			Response.Write "Original          = " & myObj.Original          & "<br>"
			Response.Write "Emphasis          = " & myObj.Emphasis          & "<br>"
			Response.Write "SongName          = " & myObj.SongName          & "<br>"
			Response.Write "Artist            = " & myObj.Artist            & "<br>"
			Response.Write "Album             = " & myObj.Album             & "<br>"
			Response.Write "AudioYear         = " & myObj.AudioYear         & "<br>"
			Response.Write "Genre             = " & myObj.Genre             & "<br>"
			Response.Write "Comment           = " & myObj.Comment           & "<br>"

            Response.Write "<br>"
		End If
	Next

%>

I hope you find this contribution useful and handy! 
If anybody has any questions please email me at 
heidel@f256.com or visit http://www.f256.com 

Happy Programming!
