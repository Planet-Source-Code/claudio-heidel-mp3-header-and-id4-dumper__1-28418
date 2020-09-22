<%
'-------------------------------------------------------------
'  Create Date : 24/10/2001 (dd/mm/yyyy)
'  Mod. Date   : 25/10/2001
'  Author      : Claudio Heidel (heidel@f256.com)
'-------------------------------------------------------------

Class mp3dump


	Dim msize
	Dim mfile
	Dim mheader
	Dim mauxbin
	Dim mframesync
	Dim maudioversion
	Dim mlayerdescription
	Dim mcrc
	Dim mbitrate
	Dim v
	Dim l
	Dim mbarray(15,5)
	Dim msamplerate
	Dim mpading
	Dim mframelenght
	Dim mframes
	Dim mprivate
	Dim mchannelmode
	Dim mextensionmode
	Dim mcopyright
	Dim moriginal
	Dim memphasis
	Dim msongname
	Dim martist
	Dim malbum
	Dim myear
	Dim mcomment
	Dim mgenre
	Dim mgarray(125)
	Dim mtime
	
	Private Sub Class_Initialize()
	
		mbarray(0,0)     = "free"
		mbarray(0,1)     = "free"
		mbarray(0,2)     = "free"
		mbarray(0,3)     = "free"
		mbarray(0,4)     = "free"
		mbarray(0,5)     = "free"
		mbarray(1,0)     = "32"
		mbarray(1,1)     = "32"
		mbarray(1,2)     = "32"
		mbarray(1,3)     = "32"
		mbarray(1,4)     = "8"
		mbarray(1,5)     = "8"
		mbarray(2,0)     = "64"
		mbarray(2,1)     = "48"
		mbarray(2,2)     = "40"
		mbarray(2,3)     = "48"
		mbarray(2,4)     = "16"
		mbarray(2,5)     = "16"
		mbarray(3,0)     = "96"
		mbarray(3,1)     = "56"
		mbarray(3,2)     = "48"
		mbarray(3,3)     = "56"
		mbarray(3,4)     = "24"
		mbarray(3,5)     = "24"
		mbarray(4,0)     = "128"
		mbarray(4,1)     = "64"
		mbarray(4,2)     = "56"
		mbarray(4,3)     = "64"
		mbarray(4,4)     = "32"
		mbarray(4,5)     = "32"
		mbarray(5,0)     = "160"
		mbarray(5,1)     = "80"
		mbarray(5,2)     = "64"
		mbarray(5,3)     = "80"
		mbarray(5,4)     = "40"
		mbarray(5,5)     = "40"
		mbarray(6,0)     = "192"
		mbarray(6,1)     = "96"
		mbarray(6,2)     = "80"
		mbarray(6,3)     = "96"
		mbarray(6,4)     = "48"
		mbarray(6,5)     = "48"
		mbarray(7,0)     = "224"
		mbarray(7,1)     = "112"
		mbarray(7,2)     = "96"
		mbarray(7,3)     = "112"
		mbarray(7,4)     = "56"
		mbarray(7,5)     = "56"
		mbarray(8,0)     = "256"
		mbarray(8,1)     = "128"
		mbarray(8,2)     = "112"
		mbarray(8,3)     = "128"
		mbarray(8,4)     = "64"
		mbarray(8,5)     = "64"
		mbarray(9,0)     = "288"
		mbarray(9,1)     = "160"
		mbarray(9,2)     = "128"
		mbarray(9,3)     = "144"
		mbarray(9,4)     = "80"
		mbarray(9,5)     = "80"
		mbarray(10,0)    = "320"
		mbarray(10,1)    = "192"
		mbarray(10,2)    = "160"
		mbarray(10,3)    = "160"
		mbarray(10,4)    = "96"
		mbarray(10,5)    = "96"
		mbarray(11,0)    = "352"
		mbarray(11,1)    = "224"
		mbarray(11,2)    = "192"
		mbarray(11,3)    = "176"
		mbarray(11,4)    = "112"
		mbarray(11,5)    = "112"
		mbarray(12,0)    = "384"
		mbarray(12,1)    = "256"
		mbarray(12,2)    = "224"
		mbarray(12,3)    = "192"
		mbarray(12,4)    = "128"
		mbarray(12,5)    = "128"
		mbarray(13,0)    = "416"
		mbarray(13,1)    = "320"
		mbarray(13,2)    = "256"
		mbarray(13,3)    = "224"
		mbarray(13,4)    = "144"
		mbarray(13,5)    = "144"
		mbarray(14,0)    = "448"
		mbarray(14,1)    = "384"
		mbarray(14,2)    = "320"
		mbarray(14,3)    = "256"
		mbarray(14,4)    = "160"
		mbarray(14,5)    = "160"
		mbarray(15,0)    = "bad"
		mbarray(15,1)    = "bad"
		mbarray(15,2)    = "bad"
		mbarray(15,3)    = "bad"
		mbarray(15,4)    = "bad"
		mbarray(15,5)    = "bad"
	
	
	
		mgarray(0)   = "Blues"
		mgarray(1)   = "Classic Rock"
		mgarray(2)   = "Country"
		mgarray(3)   = "Dance"
		mgarray(4)   = "Disco"
		mgarray(5)   = "Funk"
		mgarray(6)   = "Grunge"
		mgarray(7)   = "Hip-Hop"
		mgarray(8)   = "Jazz"
		mgarray(9)   = "Metal"
		mgarray(10)  = "New Age"
		mgarray(11)  = "Oldies"
		mgarray(12)  = "Other"
		mgarray(13)  = "Pop"
		mgarray(14)  = "R&B"
		mgarray(15)  = "Rap"
		mgarray(16)  = "Reggae"
		mgarray(17)  = "Rock"
		mgarray(18)  = "Techno"
		mgarray(19)  = "Industrial"
		mgarray(20)  = "Alternative"
		mgarray(21)  = "Ska"
		mgarray(22)  = "Death Metal"
		mgarray(23)  = "Pranks"
		mgarray(24)  = "Soundtrack"
		mgarray(25)  = "Euro-Techno"
		mgarray(26)  = "Ambient"
		mgarray(27)  = "Trip-Hop"
		mgarray(28)  = "Vocal"
		mgarray(29)  = "Jazz+Funk"
		mgarray(30)  = "Fusion"
		mgarray(31)  = "Trance"
		mgarray(32)  = "Classical"
		mgarray(33)  = "Instrumental"
		mgarray(34)  = "Acid"
		mgarray(35)  = "House"
		mgarray(36)  = "Game"
		mgarray(37)  = "Sound Clip"
		mgarray(38)  = "Gospel"
		mgarray(39)  = "Noise"
		mgarray(40)  = "AlternRock"
		mgarray(41)  = "Bass"
		mgarray(42)  = "Soul"
		mgarray(43)  = "Punk"
		mgarray(44)  = "Space"
		mgarray(45)  = "Meditative"
		mgarray(46)  = "Instrumental Pop"
		mgarray(47)  = "Instrumental Rock"
		mgarray(48)  = "Ethnic"
		mgarray(49)  = "Gothic"
		mgarray(50)  = "Darkwave"
		mgarray(51)  = "Techno-Industrial"
		mgarray(52)  = "Electronic"
		mgarray(53)  = "Pop-Folk"
		mgarray(54)  = "Eurodance"
		mgarray(55)  = "Dream"
		mgarray(56)  = "Southern Rock"
		mgarray(57)  = "Comedy"
		mgarray(58)  = "Cult"
		mgarray(59)  = "Gangsta"
		mgarray(60)  = "Top 40"
		mgarray(61)  = "Christian Rap"
		mgarray(62)  = "Pop/Funk"
		mgarray(63)  = "Jungle"
		mgarray(64)  = "Native American"
		mgarray(65)  = "Cabaret"
		mgarray(66)  = "New Wave"
		mgarray(67)  = "Psychadelic"
		mgarray(68)  = "Rave"
		mgarray(69)  = "Showtunes"
		mgarray(70)  = "Trailer"
		mgarray(71)  = "Lo-Fi"
		mgarray(72)  = "Tribal"
		mgarray(73)  = "Acid Punk"
		mgarray(74)  = "Acid Jazz"
		mgarray(75)  = "Polka"
		mgarray(76)  = "Retro"
		mgarray(77)  = "Musical"
		mgarray(78)  = "Rock & Roll"
		mgarray(79)  = "Hard Rock"
		'WinAmp expanded the above with the following:
		mgarray(80)  = "Folk"
		mgarray(81)  = "Folk-Rock"
		mgarray(82)  = "National Folk"
		mgarray(83)  = "Swing"
		mgarray(84)  = "Fast Fusion"
		mgarray(85)  = "Bebob"
		mgarray(86)  = "Latin"
		mgarray(87)  = "Revival"
		mgarray(88)  = "Celtic"
		mgarray(89)  = "Bluegrass"
		mgarray(90)  = "Avantgarde"
		mgarray(91)  = "Gothic Rock"
		mgarray(92)  = "Progressive Rock"
		mgarray(93)  = "Psychedelic Rock"
		mgarray(94)  = "Symphonic Rock"
		mgarray(95)  = "Slow Rock"
		mgarray(96)  = "Big Band"
		mgarray(97)  = "Chorus"
		mgarray(98)  = "Easy Listening"
		mgarray(99)  = "Acoustic"
		mgarray(100) = "Humour"
		mgarray(101) = "Speech"
		mgarray(102) = "Chanson"
		mgarray(103) = "Opera"
		mgarray(104) = "Chamber Music"
		mgarray(105) = "Sonata"
		mgarray(106) = "Symphony"
		mgarray(107) = "Booty Brass"
		mgarray(108) = "Primus"
		mgarray(109) = "Porn Groove"
		mgarray(110) = "Satire"
		mgarray(111) = "Slow Jam"
		mgarray(112) = "Club"
		mgarray(113) = "Tango"
		mgarray(114) = "Samba"
		mgarray(115) = "Folklore"
		mgarray(116) = "Ballad"
		mgarray(117) = "Poweer Ballad"
		mgarray(118) = "Rhytmic Soul"
		mgarray(119) = "Freestyle"
		mgarray(120) = "Duet"
		mgarray(121) = "Punk Rock"
		mgarray(122) = "Drum Solo"
		mgarray(123) = "A Capela"
		mgarray(124) = "Euro-House"
		mgarray(125) = "Dance Hall"
	

	End Sub

	Private Sub Class_Terminate()

	End Sub

	Private Function Readmheader (filename,nbites)
	   Const ForReading = 1, ForWriting = 2, ForAppending = 8
	   Dim fso, f
	   Set fso = CreateObject("Scripting.FileSystemObject")
	   Set f = fso.OpenTextFile(filename, ForReading)
	   Readmheader = f.Read(nbites)
	End Function
	
	Private Function ReadID3 (filename)
	   Const ForReading = 1, ForWriting = 2, ForAppending = 8
	   Dim fso
	   Dim f
	   Set fso = CreateObject("Scripting.FileSystemObject")
	   Set f = fso.OpenTextFile(filename, ForReading)
	   f.Skip(FileSize(filename) - 128)
	   ReadID3 = f.Read(128)
	End Function
	
	Private Function FileSize(filename)
	   Dim fso
	   Dim f1
	   Set fso  = CreateObject("Scripting.FileSystemObject")
	   Set f1   = fso.GetFile(filename)
	   FileSize = f1.size
	   Set fso  = Nothing
	   Set f1   = Nothing
	End Function

	Private Function ToBin(inNumber, OutLenStr )
		Dim binary
		binary = ""
		do while inNumber >= 1
			binary = binary & inNumber mod 2
			inNumber = inNumber \ 2
		loop
		binary = binary & String(OutLenStr - len(binary), "0")
		ToBin = StrReverse(binary)
	End Function

	Private Function Bin2Decimal(inBin)
		Dim counter
		Dim temp
		Dim Value
		inBin = StrReverse(inBin)
		temp = 0
		For counter = 1 to Len(inBin)
			If counter = 1 then
				Value = 1
			Else
				Value = Value  * 2
			End If
			temp = temp + mid(inBin, counter ,1)  *  Value
		Next
		Bin2Decimal = temp
	End Function


	Private Function MP3Dump(fileName)

		
		msize   = FileSize(fileName)
		mheader = Readmheader (fileName,4)
		
		mauxbin = ToBin(asc(mid(mheader,1,1)),8)
		mauxbin = mauxbin & ToBin(asc(mid(mheader,2,1)),8)
		mauxbin = mauxbin & ToBin(asc(mid(mheader,3,1)),8)
		mauxbin = mauxbin & ToBin(asc(mid(mheader,4,1)),8)

		mframesync        = mid(mauxbin,1,11)
		maudioversion     = mid(mauxbin,12,2)

		Select Case maudioversion
		Case "00"
		  maudioversion = 2.5
		  v            = 2
		Case "01"
		  maudioversion = "reserved"
		  v            = 2
		Case "10"
		  maudioversion = 2
		  v            = 2
		Case "11"
		  maudioversion = 1
		  v            = 1
		End Select

		mlayerdescription = mid(mauxbin,14,2)
		Select Case mlayerdescription
		Case "00"
		  mlayerdescription = "reserved"
		  l                 = 1
		Case "01"
		  mlayerdescription = "Layer III"
		  l                 = 3
		Case "10"
		  mlayerdescription = "Layer II"
		  l                 = 2
		Case "11"
		  mlayerdescription = "Layer I"
		  l                 = 1
		End Select

		mcrc = Not CBool(mid(mauxbin,16,1))

		mbitrate = mid(mauxbin,17,4)
		mbitrate = Bin2Decimal(mbitrate)

		if v = 1 then
			mbitrate = mbarray(mbitrate ,l - 1)
		else
		mbitrate = mbarray(mbitrate ,l + 2)
		end if

		msamplerate = mid(mauxbin,21,2)

		Select Case maudioversion
		Case 2.5
		  Select Case msamplerate
		  Case "00"
		    msamplerate = "11025"
		  Case "01"
		    msamplerate = "12000"
		  Case "10"
		    msamplerate = "8000"
		  Case "11"
		    msamplerate = "0"
		  End Select
		Case 2
		  Select Case msamplerate
		  Case "00"
		    msamplerate = "22050"
		  Case "01"
		    msamplerate = "24000"
		  Case "10"
		    msamplerate = "16000"
		  Case "11"
		    msamplerate = "0"
		  End Select
		Case 1
		  Select Case msamplerate
		  Case "00"
		    msamplerate = "44100"
		  Case "01"
		    msamplerate = "48000"
		  Case "10"
		    msamplerate = "32000"
		  Case "11"
		    msamplerate = "0"
		  End Select
		End Select

	        mpading  = mid(mauxbin,23,1)
	        mprivate = CBool(mid(mauxbin,24,1))

	        If l = 1 Then
	          mframelenght = Int((mbitrate * 1000 * 12 / msamplerate + mpading) * 4)
	        Else
		  mframelenght = Int(144 * mbitrate * 1000 / msamplerate )
		End If

		mchannelmode = mid(mauxbin,25,2)
		Select Case mchannelmode
		Case "00"
		  mchannelmode = "stereo"
		Case "01"
		  mchannelmode = "joint stereo"
		Case "10"
		  mchannelmode = "dual stereo"
		Case "11"
		  mchannelmode = "mono"
		End Select


	        mextensionmode = mid(mauxbin,27,2)
		mcopyright     = CBool(mid(mauxbin,29,1))
		moriginal      = CBool(mid(mauxbin,30,1))
		memphasis      = CBool(mid(mauxbin,31,2))

		Select Case memphasis
		Case "00"
		  memphasis = "none"
		Case "01"
		  memphasis = "50/15 ms"
		Case "10"
		  memphasis = "reserved"
		Case "11"
		  memphasis = "CCIT J.17"
		End Select
		
		mframes = Int (msize / (mframelenght / 2))
		
	        If l = 1 Then
	          mtime = Int(((384 / msamplerate) * mframes) / 2)
	        Else
		  mtime = Int(((1152 / msamplerate) * mframes) / 2)
		End If
		
	End Function
	
	
	Private Function ID3Dump (filename)
		Dim inaux
		
		inaux = ReadID3 (fileName)
		
		If UCase(Mid(inaux,1,3)) = "TAG" Then
			msongname  = Mid(inaux,4,30)
			martist    = Mid(inaux,34,30)
			malbum     = Mid(inaux,64,30)
			myear      = Mid(inaux,94,4)
			mcomment   = Mid(inaux,98,30)
			mgenre     = mgarray(Asc(Mid(inaux,128,1)))
	        Else
			msongname  = ""
			martist    = ""
			malbum     = ""
			myear      = ""
			mcomment   = ""
			mgenre     = ""
	        End If
		
	End Function
	
	Public Property Get File()
		File = mfile
	End Property
	
	Public Property Let File(inValue)
		mfile = inValue
		MP3Dump mfile
		ID3Dump mfile
	End Property
	
	Public Property Get FrameSync()
		FrameSync = mframesync
	End Property

	Public Property Get AudioVersion()
		AudioVersion = maudioversion
	End Property

	Public Property Get LayerDescription()
		LayerDescription = mlayerdescription
	End Property

	Public Property Get CRC()
		CRC = mcrc
	End Property

	Public Property Get BitRate()
		BitRate = mbitrate
	End Property

	Public Property Get SampleRate()
		SampleRate = msamplerate
	End Property

	Public Property Get Pading()
		Pading = mpading
	End Property

	Public Property Get FrameLenght()
		FrameLenght = mframelenght
	End Property

	Public Property Get IsPrivate()
		IsPrivate = mprivate
	End Property
	
	Public Property Get ChannelMode()
		ChannelMode = mchannelmode
	End Property

	Public Property Get CopyRight()
		CopyRight = mcopyright
	End Property

	Public Property Get Original()
		Original = moriginal
	End Property

	Public Property Get Emphasis()
		Emphasis = memphasis
	End Property
	
	Public Property Get SongName()
		SongName = msongname  
	End Property
	
	Public Property Get Artist()
		Artist = martist  
	End Property
	
   	Public Property Get Album()
   		album = malbum   
	End Property
	
   	Public Property Get AudioYear()
   		AudioYear = myear   
	End Property

   	Public Property Get Comment()
   		Comment = mcomment
	End Property
	
   	Public Property Get Genre()
   		Genre = mgenre
	End Property
	
   	Public Property Get Size()
   		Size = msize
	End Property
	
   	Public Property Get Frames()
   		Frames = mframes
	End Property
	
   	Public Property Get AudioTime()
   		AudioTime = mtime
	End Property
End Class

%>