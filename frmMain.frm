VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Decoder"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3720
      TabIndex        =   30
      Text            =   "0"
      Top             =   2400
      Width           =   735
   End
   Begin MSComctlLib.Slider slPos 
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   1440
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      Max             =   400
      TickFrequency   =   20
   End
   Begin VB.CommandButton btnPause 
      Caption         =   "Pause"
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton btnBrowseWAV 
      Caption         =   "Browse"
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtWAV 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1080
      Width           =   3615
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtMP3 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4320
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnBrowseMP3 
      Caption         =   "Browse"
      Height          =   285
      Left            =   3840
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblID3V2 
      Caption         =   "Label20"
      Height          =   255
      Left            =   4080
      TabIndex        =   45
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label lblID3V1 
      Caption         =   "Label19"
      Height          =   255
      Left            =   2400
      TabIndex        =   44
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label lblGenre 
      Caption         =   "Label19"
      Height          =   255
      Left            =   3240
      TabIndex        =   43
      Top             =   4080
      Width           =   2815
   End
   Begin VB.Label lblComment 
      Caption         =   "Label19"
      Height          =   255
      Left            =   3240
      TabIndex        =   42
      Top             =   3840
      Width           =   2815
   End
   Begin VB.Label lblYear 
      Caption         =   "Label19"
      Height          =   255
      Left            =   3240
      TabIndex        =   41
      Top             =   3600
      Width           =   2815
   End
   Begin VB.Label lblAlbum 
      Caption         =   "Label19"
      Height          =   255
      Left            =   3240
      TabIndex        =   40
      Top             =   3360
      Width           =   2815
   End
   Begin VB.Label lblTitle 
      Caption         =   "Label19"
      Height          =   255
      Left            =   3240
      TabIndex        =   39
      Top             =   3120
      Width           =   2815
   End
   Begin VB.Label lblArtist 
      Caption         =   "Label19"
      Height          =   255
      Left            =   3240
      TabIndex        =   38
      Top             =   2880
      Width           =   2815
   End
   Begin VB.Label Label18 
      Caption         =   "Genre:"
      Height          =   255
      Left            =   2400
      TabIndex        =   37
      Top             =   4080
      Width           =   800
   End
   Begin VB.Label Label17 
      Caption         =   "Comment:"
      Height          =   255
      Left            =   2400
      TabIndex        =   36
      Top             =   3840
      Width           =   800
   End
   Begin VB.Label Label16 
      Caption         =   "Year:"
      Height          =   255
      Left            =   2400
      TabIndex        =   35
      Top             =   3600
      Width           =   800
   End
   Begin VB.Label Label15 
      Caption         =   "Album:"
      Height          =   255
      Left            =   2400
      TabIndex        =   34
      Top             =   3360
      Width           =   800
   End
   Begin VB.Label Label14 
      Caption         =   "Title:"
      Height          =   255
      Left            =   2400
      TabIndex        =   33
      Top             =   3120
      Width           =   800
   End
   Begin VB.Label Label13 
      Caption         =   "Artist:"
      Height          =   255
      Left            =   2400
      TabIndex        =   32
      Top             =   2880
      Width           =   800
   End
   Begin VB.Label Label12 
      Caption         =   "Frames to look for:"
      Height          =   255
      Left            =   2280
      TabIndex        =   31
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblVBR 
      Caption         =   "Label12"
      Height          =   255
      Left            =   1200
      TabIndex        =   28
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "VBR:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblFrames 
      Caption         =   "Label11"
      Height          =   255
      Left            =   1200
      TabIndex        =   26
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Frames:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblChannels 
      Caption         =   "Label10"
      Height          =   255
      Left            =   1200
      TabIndex        =   24
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Channels:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblMode 
      Caption         =   "Label9"
      Height          =   255
      Left            =   1200
      TabIndex        =   22
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Mode:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblLayer 
      Caption         =   "Label8"
      Height          =   255
      Left            =   1200
      TabIndex        =   20
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Layer:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblLevel 
      Caption         =   "Label7"
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Level:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblFrequency 
      Caption         =   "Label6"
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblState 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Label lblTimeElapsed 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   1980
      Width           =   1095
   End
   Begin VB.Label lblDuration 
      Caption         =   "Label6"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblBitRate 
      Caption         =   "Label6"
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Frequency:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Duration:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Bit Rate:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Destination File (WAV)"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Source File (MP3)"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim prevState As XA_PlayerState

Private Sub btnBrowseMP3_Click()
Dim ret As Long
Dim hAnalyzer As Long
Dim st As String
Dim wd As Long
Dim fr As Long

'close input stream if opened
If ReceivedMsg.mInputState = XA_STATE_OPEN Then
    IssuedComm = XA_MSG_COMMAND_INPUT_CLOSE
    If SendCommand(IssuedComm, 0, 0) = False Then GoTo erh
End If

On Local Error GoTo erh
If ret < 0 Then
    MsgBox "Cannot create analyzer object."
    GoTo erh
End If
cd1.CancelError = True
cd1.Filter = "MP3 Files (play/decode)|*.mp3"
cd1.Flags = cdlOFNFileMustExist + cdlOFNLongNames
cd1.ShowOpen

MP3File = cd1.FileName
txtMP3 = MP3File
MP3FileShort = GetShortName(MP3File, True)
btnBrowseWAV.Enabled = True


'create new analyzer object
ret = xanalyzer_new(hAnalyzer)
Do While hAnalyzer = 0
    DoEvents
Loop
wd = 0
fr = Val(Text1.Text)
st = MP3FileShort & Chr(0)
'GetHeader1
'GetHeader2

'this does not work properly with callback
'initialize track info
'AnInfo.Track.Album = String(30, Chr(0))
'AnInfo.Track.Artist = String(30, Chr(0))
'AnInfo.Track.Comment = String(30, Chr(0))
'AnInfo.Track.iYear = String(4, Chr(0))
'AnInfo.Track.Title = String(30, Chr(0))

ret = xanalyzer_process_file(hAnalyzer, st, 0&, 0&, AnInfo, 0&, wd, fr)
Select Case ret
Case Is >= 0
    'ok
Case XANALYZE_ERROR_INTERNAL
    'internal error
    MsgBox "There was an internal error while tried to analize file" & vbNewLine & MP3File, vbCritical + vbOKOnly, "MP3Player"
    GoTo erh
Case XANALYSE_ERROR_OUT_OF_MEMORY
    'outof memory error
    MsgBox "Not enough memory to analize file" & vbNewLine & MP3File, vbCritical + vbOKOnly, "MP3Player"
    GoTo erh
Case XANALYZE_ERROR_NO_SUCH_FILE
    'no such file error
    MsgBox "No file was found with the name:" & vbNewLine & MP3File, vbCritical + vbOKOnly, "MP3Player"
    GoTo erh
Case XANALYZE_ERROR_CANNOT_OPEN
    'cannot open error
    MsgBox "Cannot open file" & vbNewLine & MP3File, vbCritical + vbOKOnly, "MP3Player"
    GoTo erh
Case XANALYZE_ERROR_STOP_CONDITION
    'stop condition met
    MsgBox "Stop condition was met.", vbInformation + vbOKOnly, "MP3Player"
Case XANALYZE_ERROR_WATCHDOG
    'watchdog reached
    MsgBox "Could not find frame start after looking at " & wd & "bytes.", vbInformation + vbOKOnly, "MP3Player"
    GoTo erh
Case Else
    MsgBox "Unknown error occured while trying to analyze file.", vbCritical + vbOKOnly, "MP3Player"
    GoTo erh
End Select

Select Case AnInfo.StreamType.level
Case 0
    lblLevel = "MPEG2.5"
Case 1
    lblLevel = "MPEG1"
Case 2
    lblLevel = "MPEG2"
End Select

lblLayer = AnInfo.StreamType.layer
lblBitRate = AnInfo.StreamType.bitrate
lblFrequency = AnInfo.StreamType.frequency
lblDuration = ConvertMilliSecToTime(AnInfo.Duration)
Select Case AnInfo.StreamType.mode
Case 0
    lblMode = "Stereo"
Case 1
    lblMode = "Joint-Stereo"
Case 2
    lblMode = "Dual Channel"
Case 3
    lblMode = "Mono"
End Select
lblChannels = AnInfo.StreamType.channels
If fr = 0 Then lblFrames = AnInfo.Frames Else lblFrames = ""
'XANALYZE_REPORT_CHANGING_LEVEL: the MPEG level (MPEG 1 or MPEG 2) is not constant.
'XANALYZE_REPORT_CHANGING_LAYER: the MPEG layer (1, 2 or 3) is not constant.
'XANALYZE_REPORT_CHANGING_BITRATE: the bitrate is not constant.
'XANALYZE_REPORT_CHANGING_FREQUENCY: the sampling frequency is not constant.
'XANALYZE_REPORT_CHANGING_MODE: the MPEG mode is not constant.
'XANALYZE_REPORT_CHANGING_CHANNELS: the number of channels (1 for mono, 2 for stereo) is not constant.
If (AnInfo.Flags And XANALYZE_REPORT_CHANGING_BITRATE) = XANALYZE_REPORT_CHANGING_BITRATE Then
    lblVBR.Caption = "True"
Else
    lblVBR = "False"
End If
If (AnInfo.Flags And XANALYZE_REPORT_HAS_ID3V1_HEADER) = XANALYZE_REPORT_HAS_ID3V1_HEADER Then
    lblID3V1 = "Has ID3V1"
Else
    lblID3V1 = "Does not have ID3V1"
End If
If (AnInfo.Flags And XANALYZE_REPORT_HAS_ID3V2_HEADER) = XANALYZE_REPORT_HAS_ID3V2_HEADER Then
    lblID3V2 = "Has ID3V2"
Else
    lblID3V2 = "Does not have ID3V2"
End If

lblArtist = agGetStringFromPointer(AnInfo.Track.Artist)
lblTitle = agGetStringFromPointer(AnInfo.Track.Title)
lblAlbum = agGetStringFromPointer(AnInfo.Track.Album)
lblYear = agGetStringFromPointer(AnInfo.Track.iYear)
lblComment = agGetStringFromPointer(AnInfo.Track.Comment)
lblGenre = AnInfo.Track.Genre

erh:
    ret = xanalyzer_delete(hAnalyzer)
'open input file
MP3FileShort = GetShortName(MP3File, True)
IssuedComm = XA_MSG_COMMAND_INPUT_OPEN
If SendCommand(IssuedComm, 0, 0) = False Then GoTo erh

End Sub

Private Sub btnBrowseWAV_Click()
On Error GoTo erh
If ReceivedMsg.mOutputState = XA_STATE_OPEN Then
    'close output
    IssuedComm = XA_MSG_COMMAND_OUTPUT_CLOSE
    If SendCommand(IssuedComm, 0, 0) = False Then GoTo erh
End If
cd1.CancelError = True
cd1.Flags = cdlOFNLongNames + cdlOFNOverwritePrompt
cd1.Filter = "WAV Files (*.wav)|*.wav"
cd1.ShowSave

WAVFile = cd1.FileName
If Dir(WAVFile, vbNormal) = "" Then
    'create the file
    Open WAVFile For Binary As #1
    Close #1
End If
WAVFileShort = "wav:" & GetShortName(WAVFile, False)

erh:
txtWAV.Text = WAVFile

End Sub

Private Sub btnPause_Click()
'preventing pressing this button again (when player is paused)
'  is done in message processing routine
IssuedComm = XA_MSG_COMMAND_PAUSE
SendCommand IssuedComm, 0, 0

DoEvents

End Sub

Private Sub btnStart_Click()
On Error GoTo erh
'make sure input file is open
If ReceivedMsg.mInputState = XA_STATE_CLOSED Then
    MsgBox "File to be played is not selected or" & vbNewLine & "could not open the file.", vbCritical + vbOKOnly, "MP3Player"
    Exit Sub
ElseIf ReceivedMsg.mPlayerState = XA_PLAYER_STATE_EOS Then
    'reopen the file
    IssuedComm = XA_MSG_COMMAND_INPUT_OPEN
    If SendCommand(IssuedComm, 0, 0) = False Then GoTo erh
End If
'prevent from pressing start button when player is playing
Select Case ReceivedMsg.mPlayerState
Case XA_PLAYER_STATE_EOS
Case XA_PLAYER_STATE_PAUSED
    'go to selected position
    IssuedComm = XA_MSG_COMMAND_SEEK
    SendCommand IssuedComm, slPos.Value, slPos.Max
    IssuedComm = XA_MSG_COMMAND_PLAY
    If SendCommand(IssuedComm, 0, 0) = False Then GoTo erh
    Exit Sub
Case XA_PLAYER_STATE_PLAYING
    Exit Sub
Case XA_PLAYER_STATE_STOPPED
End Select

'open output file:
'   if output file is empty string will play on speakers
'   if not empty then will encode to specified location
'   if filename preceeded with "wav:" then will create wav header
IssuedComm = XA_MSG_COMMAND_OUTPUT_OPEN
If SendCommand(IssuedComm, 0, 0) = False Then GoTo erh

'set output to stereo - not necessary
IssuedComm = XA_MSG_SET_OUTPUT_CHANNELS
If SendCommand(IssuedComm, XA_OUTPUT_CHANNELS_STEREO, 0) = False Then GoTo erh

'go to selected position
IssuedComm = XA_MSG_COMMAND_SEEK
SendCommand IssuedComm, slPos.Value, slPos.Max
'start playing here
IssuedComm = XA_MSG_COMMAND_PLAY
If SendCommand(IssuedComm, 0, 0) = False Then GoTo erh
Select Case ReceivedMsg.mPlayerState
Case XA_PLAYER_STATE_PAUSED
    lblState = "Paused"
Case XA_PLAYER_STATE_PLAYING
    lblState = "Playing"
Case XA_PLAYER_STATE_STOPPED
    lblState = "Stopped"
End Select
prevState = XA_PLAYER_STATE_PLAYING
Do 'loop while playing
    'all messages are accepted by frmPlay
    'we just show the time and position.
    lblTimeElapsed = IIf(ReceivedMsg.mTimecode.mH = 0, "", ReceivedMsg.mTimecode.mH & ":") & Format(ReceivedMsg.mTimecode.mM, "00") & ":" & Format(ReceivedMsg.mTimecode.mS, "00") & "." & Format(ReceivedMsg.mTimecode.mF, "0")
    If ReceivedMsg.mPlayerState <> XA_PLAYER_STATE_PAUSED Then slPos.Value = ReceivedMsg.mPosition.mOffset
    DoEvents
Loop While (ReceivedMsg.mPlayerState = XA_PLAYER_STATE_PLAYING) Or (ReceivedMsg.mPlayerState = XA_PLAYER_STATE_PAUSED)
'if here then play was stopped or end of stream was reached
slPos.Value = 0
lblTimeElapsed.Caption = "00:00.0"
lblState = "Stopped"
Select Case ReceivedMsg.mPlayerState
Case XA_PLAYER_STATE_PAUSED
    lblState = "Paused"
Case XA_PLAYER_STATE_PLAYING
    lblState = "Playing"
Case XA_PLAYER_STATE_STOPPED
    lblState = "Stopped"
Case XA_PLAYER_STATE_EOS
    lblState = "End Of File"
End Select

'if were decoding then close the output
If WAVFile <> "" Then
    IssuedComm = XA_MSG_COMMAND_OUTPUT_CLOSE
    SendCommand IssuedComm, 0, 0
End If

Exit Sub

erh:
'necessary to stop - on any error
IssuedComm = XA_MSG_COMMAND_STOP
SendCommand IssuedComm, 0, 0
IssuedComm = XA_MSG_COMMAND_EXIT
SendCommand IssuedComm, 0, 0

End Sub

Private Sub btnStop_Click()
'preventing pressing this button again (when player is stopped)
'  is done in message processing routine
IssuedComm = XA_MSG_COMMAND_STOP
SendCommand IssuedComm, 0, 0
'stop it in any case - even if PlayerStop is failed
DoEvents

End Sub

Private Sub Form_Load()
Dim ret As Long

'initialize player
'load form that will be a player
Load frmPlay
gHW = frmPlay.hwnd
'hook this form to process all messages
Hook
'create the player which is hooked to gHW window
ret = player_new(hPlayer, gHW)
'set user data for the window - don't know why, just in case
SetWindowLong gHW, GWL_USERDATA, hPlayer
'set high priority for player
ret = player_set_priority(hPlayer, XA_CONTROL_PRIORITY_HIGH)
'have to preinitialize some vars
ReceivedMsg.mInputState = XA_STATE_CLOSED
ReceivedMsg.mOutputState = XA_STATE_CLOSED
txtMP3.Text = ""
txtWAV.Text = ""
MP3File = ""
WAVFile = ""

'show equalizer
frmEqualizer.Show 0, frmMain

End Sub

Private Sub Form_Unload(Cancel As Integer)
If hPlayer Then
    'necessary to stop
    control_message_send_N hPlayer, XA_MSG_COMMAND_STOP
    control_message_send_N hPlayer, XA_MSG_COMMAND_EXIT
    player_delete (hPlayer)
    Unhook
    Unload frmPlay
End If

End Sub

Private Sub slPos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    'have to pause player
    If ReceivedMsg.mPlayerState = XA_PLAYER_STATE_PLAYING Then
        prevState = XA_PLAYER_STATE_PLAYING
        IssuedComm = XA_MSG_COMMAND_PAUSE
        SendCommand IssuedComm, 0, 0
    Else
        prevState = XA_PLAYER_STATE_STOPPED
    End If
End If

End Sub

Private Sub slPos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    DoEvents
    If prevState = XA_PLAYER_STATE_PLAYING Then
        'start player again
        'go to selected position
        IssuedComm = XA_MSG_COMMAND_SEEK
        SendCommand IssuedComm, slPos.Value, slPos.Max
        'start playing here
        IssuedComm = XA_MSG_COMMAND_PLAY
        SendCommand IssuedComm, 0, 0
        prevState = XA_PLAYER_STATE_PLAYING
    End If
End If

End Sub
