Attribute VB_Name = "modAnalyzer"
Option Explicit

'this info is valid until hAnalyzer is deleted or another file is processed
'   if there is no info - all strings are empty
Public Type AnalyzerTrackInfo
    Title As Long
    Artist As Long
    Album As Long
    iYear As Long
    Comment As Long
    Genre As Byte
End Type

'MPEG info
Public Type AnalyzerMpegInfo
    level As Long       '1 for MPEG1, 2 for MPEG2, 0 for MPEG2.5
    layer As Long       'MPEG Layer (1, 2 or 3)
    bitrate As Long     'BitRate in bits per second
    frequency As Long   'Sampling Frequency in Hertz
    mode As Long        'MPEG mode (0 for stereo, 1 for joint-stereo, 2 for dual-channel, 3 sor single-channel)
    channels As Long    'number of channels (1 for mono, 2 for stereo)
End Type

'this type provides info on frame in callback function
Public Type AnalyzerMpegFrameInfo
    Index As Long               'frame number
    info As AnalyzerMpegInfo    'frame info
    offset As Long              'offset of the frame from the beginning of the file
    nb_samples As Long          'number of bytes in frame
End Type

Public Type AnalyzerInfo
    StreamType As AnalyzerMpegInfo
    Track As AnalyzerTrackInfo
    Frames As Long      'number of frames to be analyzed
    Duration As Long    'duration of the stream in milliseconds
    Flags As Long       'bit flags indicating which, if any, of the parameters of the frames is not constant during the entire stream.
                        'Values are:
                        'XANALYZE_REPORT_CHANGING_LEVEL: the MPEG level (MPEG 1 or MPEG 2) is not constant.
                        'XANALYZE_REPORT_CHANGING_LAYER: the MPEG layer (1, 2 or 3) is not constant.
                        'XANALYZE_REPORT_CHANGING_BITRATE: the bitrate is not constant.
                        'XANALYZE_REPORT_CHANGING_FREQUENCY: the sampling frequency is not constant.
                        'XANALYZE_REPORT_CHANGING_MODE: the MPEG mode is not constant.
                        'XANALYZE_REPORT_CHANGING_CHANNELS: the number of channels (1 for mono, 2 for stereo) is not constant.
End Type


Public Declare Function xanalyzer_new Lib "xanalyze.dll" (hAnalyzer As Long) As Long
Public Declare Function xanalyzer_delete Lib "xanalyze.dll" (ByVal hAnalyzer As Long) As Long
Public Declare Function xanalyzer_process_file Lib "xanalyze.dll" (ByVal hAnalyzer As Long, ByVal fname As String, ByVal cbAnalyzer As Long, ByVal client As Long, AnInfo As AnalyzerInfo, ByVal stop_mask As Long, ByVal watchdog As Long, ByVal max_frames As Long) As Long


Public Const XANALYZE_REPORT_CHANGING_LEVEL = &H1
Public Const XANALYZE_REPORT_CHANGING_LAYER = &H2
Public Const XANALYZE_REPORT_CHANGING_BITRATE = &H4
Public Const XANALYZE_REPORT_CHANGING_FREQUENCY = &H8
Public Const XANALYZE_REPORT_CHANGING_MODE = &H10
Public Const XANALYZE_REPORT_CHANGING_CHANNELS = &H20
Public Const XANALYZE_REPORT_HAS_META_DATA = &H40
Public Const XANALYZE_REPORT_HAS_ID3V1_HEADER = &H80
Public Const XANALYZE_REPORT_HAS_ID3V2_HEADER = &H100
Public Const XANALYZE_REPORT_ANY_CHANGE_MASK = &H3F
Public Const XANALYZE_ERROR_INTERNAL = -2 'an internal error has occurred.
Public Const XANALYSE_ERROR_OUT_OF_MEMORY = -3 ' not enough memory to complete the operation.
Public Const XANALYZE_ERROR_NO_SUCH_FILE = -4 ' the requested file does not exist.
Public Const XANALYZE_ERROR_CANNOT_OPEN = -5 ' the requested file cannot be openned.
Public Const XANALYZE_ERROR_STOP_CONDITION = -8 ' a stop condition has occurred.
Public Const XANALYZE_ERROR_WATCHDOG = -9 ' a watchdog condition has occurred.


Public CurFrame As Long
Public AvBitRate As Long
Public FirstBitRate As Long
Public rVBR As Boolean
Public NumberFrames As Long
Public AnInfo As AnalyzerInfo

Public Function ProcessFile(ByVal an As Long, ByVal fn As String, ByVal cbfun As Long, ByVal watchdog As Long, ByVal max_fr As Long) As Long
Dim AnInfo As AnalyzerInfo
Dim ret As Long

ret = xanalyzer_process_file(an, fn, cbfun, 0, AnInfo, 0, 0, 0)
ProcessFile = ret

End Function

Public Function cbProcessFile(reserved As Long, AnFInfo As AnalyzerMpegFrameInfo) As Long
NumberFrames = NumberFrames + 1
CurFrame = AnFInfo.Index
CurFrame = AnFInfo.info.channels
CurFrame = AnFInfo.info.frequency
CurFrame = AnFInfo.info.layer
CurFrame = AnFInfo.info.level
CurFrame = AnFInfo.info.mode
CurFrame = AnFInfo.nb_samples
CurFrame = AnFInfo.offset
CurFrame = AnInfo.Frames

FirstBitRate = AnFInfo.info.bitrate
AvBitRate = AvBitRate + FirstBitRate
frmMain.lblFrames = NumberFrames
cbProcessFile = -1
'reserved = 0

End Function

Public Function ConvertMilliSecToTime(tt As Long) As String
Dim tim As String
Dim tem As Long
Dim t1 As Long

tem = tt
t1 = tem \ 3600000
If t1 = 0 Then
    tim = ""
Else
    tim = t1 & ":"
End If
    
tem = tt - t1 * 3600000
t1 = tem \ 60000
tim = tim & Format(t1, "00") & ":"

tem = tem - t1 * 60000
t1 = tem \ 1000
tim = tim & Format(t1, "00") & "."

tem = tem - t1 * 1000
t1 = tem \ 100
tim = tim & Format(t1, "0")
ConvertMilliSecToTime = tim

End Function
