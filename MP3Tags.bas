Attribute VB_Name = "mMP3Tags"
Option Explicit

Public Type MyTrackInfo1
    Title As String
    Artist As String
    Album As String
    iYear As String
    Comment As String
    Genre As String
    Track As String
    ID3V1Exist As Boolean
End Type

Public Type MyTrackInfo2
    Title As String
    Artist As String
    Album As String
    iYear As String
    Comment As String
    Genre As String
    Track As String
    ID3V2Exist As Boolean
    HeaderLength As Long
End Type

Public Type FrameData
    fMPEGVersion As String
    fLayer As Long
    fProtection As Long
    fBitRate As Long
    fChMode As Long
    fCopyRight As Long
    fFrameLength As Long
    fOriginal As Long
    fPadding As Long
    fSampleRate As Long
End Type

Public TrackInfo1 As MyTrackInfo1
Public TrackInfo2 As MyTrackInfo2


Public Sub GetHeader1()
'get ID3v1 header if any. Function returns FALSE if there is no tag
Dim st As String * 128
Dim s As String
Dim pos As Long

Open MP3File For Binary As #1
Seek #1, FileLen(MP3File) - 127
Get #1, , st
Close #1
If Left(st, 3) = "TAG" Then
    s = Mid(st, 4, 30)
    pos = InStr(s, Chr(0))
    If pos > 0 Then
        s = Left(s, pos - 1)
    End If
    TrackInfo1.Title = s
    
    s = Mid(st, 34, 30)
    pos = InStr(s, Chr(0))
    If pos > 0 Then
        s = Left(s, pos - 1)
    End If
    TrackInfo1.Artist = s
    
    s = Mid(st, 64, 30)
    pos = InStr(s, Chr(0))
    If pos > 0 Then
        s = Left(s, pos - 1)
    End If
    TrackInfo1.Album = s
    
    TrackInfo1.iYear = Mid(st, 94, 4)
    TrackInfo1.Comment = Mid(st, 98, 29)
    'track
    pos = Asc(Mid(st, 127, 1))
    If pos = 0 Then
        TrackInfo1.Track = vbNullString
    Else
        TrackInfo1.Track = CStr(pos)
    End If
    
    'genre
    TrackInfo1.Genre = Asc(Mid(st, 128, 1))
    
    TrackInfo1.ID3V1Exist = True
Else
    TrackInfo1.ID3V1Exist = False
End If

End Sub

Public Sub GetHeader2()
Dim st As String
Dim s As String
Dim hFlags As Byte
Dim extHeader As Boolean
Dim dat As Long
Dim frameID As String * 4
Dim pos As Long
Dim frameSt As String
Dim i As Long
Dim p As Long
Dim s1 As Long, s2 As Long, s3 As Long, s4 As Long
Dim j As Long
Dim ch As String

Close #1
Open MP3File For Binary As #1
st = String(10, 0)
'retrieve header
Get #1, , st
If Left(st, 3) = "ID3" Then 'tag ID3v2 is here
    'get flags
    hFlags = Val(Mid(st, 6, 1))
    If (hFlags And 64) = 64 Then extHeader = True
    'get header size - ignore most significan bit in every byte
    s1 = Asc(Mid(st, 10, 1))
    s2 = Asc(Mid(st, 9, 1))
    s2 = s2 * 128
    s3 = Asc(Mid(st, 8, 1))
    s3 = s3 * 256# * 64
    s4 = Asc(Mid(st, 7, 1))
    s4 = s4 * 256# * 256# * 32
    dat = s4 + s3 + s2 + s1
    TrackInfo2.HeaderLength = dat
    'looking for tags
    pos = Seek(1)
    Do
        Get #1, , frameID
        'get frame size
        st = String(4, 0)
        Get #1, , st
        If st = Chr(0) + Chr(0) + Chr(0) + Chr(0) Then Exit Do
        dat = 0
        For i = 0 To 3
            p = Asc(Mid(st, i + 1, 1))
            dat = dat + p * 256# ^ (3 - i)
        Next i
        'change position to end of frame header
        pos = Seek(1) + 2
        Seek #1, pos
        'here 'dat' - represents a frame size
        s = ""
        Select Case frameID
        Case "TIT2" 'title
            st = String(dat, 0)
            Get #1, , st
            st = Right(st, dat - 1)
            st = Trim(st)
            TrackInfo2.Title = st
        Case "TPE1" 'artist
            st = String(dat, 0)
            Get #1, , st
            st = Right(st, dat - 1)
            st = Trim(st)
            TrackInfo2.Artist = st
        Case "TPE2" 'artist
            st = String(dat, 0)
            Get #1, , st
            st = Right(st, dat - 1)
            st = Trim(st)
            TrackInfo2.Artist = st
        Case "TALB" 'album
            st = String(dat, 0)
            Get #1, , st
            st = Right(st, dat - 1)
            st = Trim(st)
            TrackInfo2.Album = st
        Case "TYER" 'year
            st = String(dat, 0)
            Get #1, , st
            st = Right(st, dat - 1)
            st = Trim(st)
            TrackInfo2.iYear = st
        Case "TRCK" 'track number
            st = String(dat, 0)
            Get #1, , st
            st = Right(st, dat - 1)
            st = Trim(st)
            TrackInfo2.Track = st
        Case "TCON" 'genre string
            st = String(dat, 0)
            Get #1, , st
            st = Right(st, dat - 1)
            st = Trim(st)
            TrackInfo2.Genre = st
        Case "COMM" 'comments
            st = String(dat, 0)
            Get #1, , st
            st = Right(st, dat - 1)
            st = Trim(st)
            p = InStr(st, Chr(0))
            If p > 0 Then
                st = Left(st, p - 1) & " " & Right(st, Len(st) - p)
            End If
            TrackInfo2.Comment = st
        Case Else 'anything else just skip
        End Select
        'calculate next header position
        pos = pos + dat 'add 10 for frame header size
        If pos >= TrackInfo2.HeaderLength Then Exit Do
        Seek #1, pos
    Loop
    TrackInfo2.ID3V2Exist = True
Else
    TrackInfo2.ID3V2Exist = False
End If
Close #1

End Sub

Public Sub GetFileParameters()
'retrieves data from frame header
Dim pos As Long
Dim st As String
Dim i As Long
Dim s As Long
Dim dd As Long
Dim cf(1) As FrameData
Dim fr_cnt As Long
Dim ch As String
Dim hb(2) As Byte
Dim cFrameData As FrameData

If TrackInfo2.ID3V2Exist = True Then
    pos = TrackInfo2.HeaderLength + 10
Else
    pos = 0
End If
Close #1
Open MP3File For Binary As #1
Seek #1, pos
st = String(1000, 0)
Get #1, , st
'looking for frame's start: in bits this is 1111 1111 111x xxxx
'hope that 1000 bytes would be enough to get a frame
For i = 0 To 999
    'scroll through the string to find frame's start
    ch = Mid(st, i + 1, 2)
    'convert ch to long
    s = 256# * Asc(Left(ch, 1)) + Asc(Right(ch, 1))
    
    If (s And &HFFE0) = 65504 Then 'the same as &hFFE0
        'found start - get the info
        hb(0) = Asc(Mid(st, i + 4, 1))
        hb(1) = Asc(Mid(st, i + 3, 1))
        hb(2) = Asc(Mid(st, i + 2, 1))
        'MPEG Version
        dd = hb(2) And &H18
        Select Case dd
        Case &H18 '11-MPEG Version 1
            cf(fr_cnt).fMPEGVersion = "MPEG Version 1"
        Case &H8  '01-reserved
            'not a frame - look further
            GoTo efor
        Case &H10 '10-MPEG Version 2
            cf(fr_cnt).fMPEGVersion = "MPEG Version 2"
        Case 0    '00-MPEG Version 2.5
            cf(fr_cnt).fMPEGVersion = "MPEG Version 2.5"
        End Select
        'layer
        dd = hb(2) And &H6
        Select Case dd
        Case &H0 '00-reserved
            'not a frame - look further
            GoTo efor
        Case &H4 '10-layer II
            cf(fr_cnt).fLayer = 2
        Case &H2 '01-layer III
            cf(fr_cnt).fLayer = 3
        Case &H6 '11-layer I
            cf(fr_cnt).fLayer = 1
        End Select
        'protection by CRC
        dd = hb(2) And &H1
        If dd = 0 Then
            cf(fr_cnt).fProtection = 1 'protected
        Else
            cf(fr_cnt).fProtection = 0 'not protected
        End If
        'Bitrate Index
        dd = hb(1) And &HF0
        Select Case dd
        Case 0
            cf(fr_cnt).fBitRate = 0 'free for all versions and layers
        Case &H10
            If cf(fr_cnt).fMPEGVersion = "MPEG Version 1" Then
                cf(fr_cnt).fBitRate = 32
            Else
                If cf(fr_cnt).fLayer > 1 Then
                    'for layers 2 and 3
                    cf(fr_cnt).fBitRate = 8
                Else
                    'for layer 1
                    cf(fr_cnt).fBitRate = 32
                End If
            End If
        Case &H20
            If cf(fr_cnt).fMPEGVersion = "MPEG Version 1" Then
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 64
                ElseIf cf(fr_cnt).fLayer = 2 Then
                    cf(fr_cnt).fBitRate = 48
                Else
                    cf(fr_cnt).fBitRate = 40
                End If
            Else
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 48
                Else
                    cf(fr_cnt).fBitRate = 16
                End If
            End If
        Case &H30
            If cf(fr_cnt).fMPEGVersion = "MPEG Version 1" Then
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 96
                ElseIf cf(fr_cnt).fLayer = 2 Then
                    cf(fr_cnt).fBitRate = 56
                Else
                    cf(fr_cnt).fBitRate = 48
                End If
            Else
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 56
                Else
                    cf(fr_cnt).fBitRate = 24
                End If
            End If
        Case &H40
            If cf(fr_cnt).fMPEGVersion = "MPEG Version 1" Then
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 128
                ElseIf cf(fr_cnt).fLayer = 2 Then
                    cf(fr_cnt).fBitRate = 64
                Else
                    cf(fr_cnt).fBitRate = 56
                End If
            Else
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 64
                Else
                    cf(fr_cnt).fBitRate = 32
                End If
            End If
        Case &H50
            If cf(fr_cnt).fMPEGVersion = "MPEG Version 1" Then
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 160
                ElseIf cf(fr_cnt).fLayer = 2 Then
                    cf(fr_cnt).fBitRate = 80
                Else
                    cf(fr_cnt).fBitRate = 64
                End If
            Else
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 80
                Else
                    cf(fr_cnt).fBitRate = 40
                End If
            End If
        Case &H60
            If cf(fr_cnt).fMPEGVersion = "MPEG Version 1" Then
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 192
                ElseIf cf(fr_cnt).fLayer = 2 Then
                    cf(fr_cnt).fBitRate = 96
                Else
                    cf(fr_cnt).fBitRate = 80
                End If
            Else
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 96
                Else
                    cf(fr_cnt).fBitRate = 48
                End If
            End If
        Case &H70
            If cf(fr_cnt).fMPEGVersion = "MPEG Version 1" Then
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 224
                ElseIf cf(fr_cnt).fLayer = 2 Then
                    cf(fr_cnt).fBitRate = 112
                Else
                    cf(fr_cnt).fBitRate = 96
                End If
            Else
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 112
                Else
                    cf(fr_cnt).fBitRate = 56
                End If
            End If
        Case &H80
            If cf(fr_cnt).fMPEGVersion = "MPEG Version 1" Then
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 256
                ElseIf cf(fr_cnt).fLayer = 2 Then
                    cf(fr_cnt).fBitRate = 128
                Else
                    cf(fr_cnt).fBitRate = 112
                End If
            Else
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 128
                Else
                    cf(fr_cnt).fBitRate = 64
                End If
            End If
        Case &H90
            If cf(fr_cnt).fMPEGVersion = "MPEG Version 1" Then
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 288
                ElseIf cf(fr_cnt).fLayer = 2 Then
                    cf(fr_cnt).fBitRate = 160
                Else
                    cf(fr_cnt).fBitRate = 128
                End If
            Else
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 144
                Else
                    cf(fr_cnt).fBitRate = 80
                End If
            End If
        Case &HA0
            If cf(fr_cnt).fMPEGVersion = "MPEG Version 1" Then
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 320
                ElseIf cf(fr_cnt).fLayer = 2 Then
                    cf(fr_cnt).fBitRate = 192
                Else
                    cf(fr_cnt).fBitRate = 160
                End If
            Else
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 160
                Else
                    cf(fr_cnt).fBitRate = 96
                End If
            End If
        Case &HB0
            If cf(fr_cnt).fMPEGVersion = "MPEG Version 1" Then
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 352
                ElseIf cf(fr_cnt).fLayer = 2 Then
                    cf(fr_cnt).fBitRate = 224
                Else
                    cf(fr_cnt).fBitRate = 192
                End If
            Else
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 176
                Else
                    cf(fr_cnt).fBitRate = 112
                End If
            End If
        Case &HC0
            If cf(fr_cnt).fMPEGVersion = "MPEG Version 1" Then
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 384
                ElseIf cf(fr_cnt).fLayer = 2 Then
                    cf(fr_cnt).fBitRate = 256
                Else
                    cf(fr_cnt).fBitRate = 224
                End If
            Else
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 192
                Else
                    cf(fr_cnt).fBitRate = 128
                End If
            End If
        Case &HD0
            If cf(fr_cnt).fMPEGVersion = "MPEG Version 1" Then
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 416
                ElseIf cf(fr_cnt).fLayer = 2 Then
                    cf(fr_cnt).fBitRate = 320
                Else
                    cf(fr_cnt).fBitRate = 256
                End If
            Else
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 224
                Else
                    cf(fr_cnt).fBitRate = 144
                End If
            End If
        Case &HE0
            If cf(fr_cnt).fMPEGVersion = "MPEG Version 1" Then
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 448
                ElseIf cf(fr_cnt).fLayer = 2 Then
                    cf(fr_cnt).fBitRate = 384
                Else
                    cf(fr_cnt).fBitRate = 320
                End If
            Else
                If cf(fr_cnt).fLayer = 1 Then
                    cf(fr_cnt).fBitRate = 256
                Else
                    cf(fr_cnt).fBitRate = 160
                End If
            End If
        Case &HF0
            cf(fr_cnt).fBitRate = -1
            GoTo efor
        End Select
        'if second frame then save data and exit
        If fr_cnt >= 1 Then
            cFrameData.fBitRate = cf(0).fBitRate
            cFrameData.fChMode = cf(0).fChMode
            cFrameData.fCopyRight = cf(0).fCopyRight
            cFrameData.fFrameLength = cf(0).fFrameLength
            cFrameData.fLayer = cf(0).fLayer
            cFrameData.fMPEGVersion = cf(0).fMPEGVersion
            cFrameData.fOriginal = cf(0).fOriginal
            cFrameData.fPadding = cf(0).fPadding
            cFrameData.fProtection = cf(0).fProtection
            cFrameData.fSampleRate = cf(0).fSampleRate
            Exit For
        End If
        'Sampling Rate Frequency
        dd = hb(1) And 12
        Select Case dd
        Case 0 '
            If cf(fr_cnt).fMPEGVersion = "MPEG Version 1" Then
                cf(fr_cnt).fSampleRate = 44100
            ElseIf cf(fr_cnt).fMPEGVersion = "MPEG Version 2" Then
                cf(fr_cnt).fSampleRate = 22050
            ElseIf cf(fr_cnt).fMPEGVersion = "MPEG Version 2.5" Then
                cf(fr_cnt).fSampleRate = 11025
            End If
        Case 4 '
            If cf(fr_cnt).fMPEGVersion = "MPEG Version 1" Then
                cf(fr_cnt).fSampleRate = 48000
            ElseIf cf(fr_cnt).fMPEGVersion = "MPEG Version 2" Then
                cf(fr_cnt).fSampleRate = 24000
            ElseIf cf(fr_cnt).fMPEGVersion = "MPEG Version 2.5" Then
                cf(fr_cnt).fSampleRate = 12000
            End If
        Case 8 '
            If cf(fr_cnt).fMPEGVersion = "MPEG Version 1" Then
                cf(fr_cnt).fSampleRate = 32000
            ElseIf cf(fr_cnt).fMPEGVersion = "MPEG Version 2" Then
                cf(fr_cnt).fSampleRate = 16000
            ElseIf cf(fr_cnt).fMPEGVersion = "MPEG Version 2.5" Then
                cf(fr_cnt).fSampleRate = 8000
            End If
        Case 12 'reserved
        End Select
        'Padding
        dd = hb(1) And 2
        If dd = 0 Then 'frame is not padded
            cf(fr_cnt).fPadding = 0
        Else 'frame is padded
            cf(fr_cnt).fPadding = 1
        End If
        'now can calculate frame length
        If cf(fr_cnt).fLayer = 1 Then
            cf(fr_cnt).fFrameLength = (12 * cf(fr_cnt).fBitRate * 1000 / cf(fr_cnt).fSampleRate + cf(fr_cnt).fPadding) * 4
        Else
            cf(fr_cnt).fFrameLength = 144 * cf(fr_cnt).fBitRate * 1000 / cf(fr_cnt).fSampleRate + cf(fr_cnt).fPadding
        End If
        cf(fr_cnt).fFrameLength = cf(fr_cnt).fFrameLength - 2
        'try to find how many frames in the file - this is valid for CBR
'        Dim tf As Long
'        tf = (MP3FileData.mFileLength - FileTagID3V2.tTagSize - 10 - IIf(FileTagID3V1.tExist, 128, 0)) \ cf(0).fFrameLength
'        'and now time 26 msec per frame
'        MP3FileData.mSeconds = tf * 0.026
'        MP3FileData.mTime = SecondsToTime(MP3FileData.mSeconds)
'        MP3FileData.mBitrate = cf(0).fBitRate
        
        'skip private bit
        'go on with channel mode
        dd = hb(0) And &HC0
        Select Case dd
        Case 0 '
            cf(fr_cnt).fChMode = "Stereo"
        Case &H80 '
            cf(fr_cnt).fChMode = "Single Channel (Mono)"
        Case &H40 '
            cf(fr_cnt).fChMode = "Joint Stereo (Stereo)"
        Case &HC0 '
            cf(fr_cnt).fChMode = "Dual Channel (Stereo)"
        End Select
        'skip Extension Mode
        'go on with Copyright
        dd = hb(0) And 8
        If dd = 0 Then 'not copyrighted
            cf(fr_cnt).fCopyRight = 0
        Else 'copyrighted
            cf(fr_cnt).fCopyRight = 1
        End If
        'original
        dd = hb(0) And 4
        If dd = 0 Then 'copy of original media
            cf(fr_cnt).fOriginal = 0
        Else 'original media
            cf(fr_cnt).fOriginal = 1
        End If
        fr_cnt = fr_cnt + 1
        i = i + cf(0).fFrameLength
    End If
efor:
Next i
Close #1

End Sub

Public Sub SelectTagToDisplay()
'If TrackInfo1.ID3V1Exist Then
'    MP3FileData.mArtist = FileTagID3V1.tArtist
'    MP3FileData.mAlbum = FileTagID3V1.tAlbum
'    MP3FileData.mComments = FileTagID3V1.tComments
'    MP3FileData.mGenreSt = FileTagID3V1.tGenreSt
'    MP3FileData.mGenreNum = FileTagID3V1.tGenreNum
'    MP3FileData.mTitle = FileTagID3V1.tTitle
'    MP3FileData.mTrack = FileTagID3V1.tTrack
'    MP3FileData.mYear = FileTagID3V1.tYear
'ElseIf TrackInfo2.ID3V2Exist Then
'    MP3FileData.mArtist = FileTagID3V2.tArtist
'    MP3FileData.mAlbum = FileTagID3V2.tAlbum
'    MP3FileData.mComments = FileTagID3V2.tComments
'    MP3FileData.mGenreSt = FileTagID3V2.tGenreSt
'    MP3FileData.mGenreNum = FileTagID3V2.tGenreNum
'    MP3FileData.mTitle = FileTagID3V2.tTitle
'    MP3FileData.mTrack = FileTagID3V2.tTrack
'    MP3FileData.mYear = FileTagID3V2.tYear
'Else
'    MP3FileData.mTitle = Right(MP3File, Len(MP3File) - 4)
'End If
'If MP3FileData.mAlbum = "" Then MP3FileData.mAlbum = "<blank>"
'If MP3FileData.mArtist = "" Then MP3FileData.mArtist = "<blank>"
'If MP3FileData.mGenreSt = "" Then MP3FileData.mGenreSt = "<blank>"

End Sub
