<!-- #include file="mp3_class.inc" -->
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
			Response.Write "AudioTime         = " & myObj.AudioTime         & "<br>"
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