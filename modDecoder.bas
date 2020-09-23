Attribute VB_Name = "modDecoder"
Option Explicit

Public Const XA_DECODER_EQUALIZER_NB_BANDS = 31

'MESSAGE CODES
Public Enum XA_MessageCode
    XA_MSG_UNKNOWN = 0
    
    'commands to decoder
    XA_MSG_COMMAND_EXIT = 1
    XA_MSG_COMMAND_SYNC = 2
    XA_MSG_COMMAND_PING = 3
    XA_MSG_COMMAND_PLAY = 4
    XA_MSG_COMMAND_PAUSE = 5
    XA_MSG_COMMAND_STOP = 6
    XA_MSG_COMMAND_SEEK = 7
    XA_MSG_COMMAND_INPUT_OPEN = 8
    XA_MSG_COMMAND_INPUT_CLOSE = 9
    XA_MSG_COMMAND_INPUT_SEND_MESSAGE = 10
    XA_MSG_COMMAND_INPUT_ADD_FILTER = 11
    XA_MSG_COMMAND_INPUT_REMOVE_FILTER = 12
    XA_MSG_COMMAND_INPUT_FILTERS_LIST = 13
    XA_MSG_COMMAND_INPUT_MODULE_REGISTER = 14
    XA_MSG_COMMAND_INPUT_MODULE_QUERY = 15
    XA_MSG_COMMAND_INPUT_MODULES_LIST = 16
    XA_MSG_COMMAND_OUTPUT_OPEN = 17
    XA_MSG_COMMAND_OUTPUT_CLOSE = 18
    XA_MSG_COMMAND_OUTPUT_SEND_MESSAGE = 19
    XA_MSG_COMMAND_OUTPUT_MUTE = 20
    XA_MSG_COMMAND_OUTPUT_UNMUTE = 21
    XA_MSG_COMMAND_OUTPUT_RESET = 22
    XA_MSG_COMMAND_OUTPUT_DRAIN = 23
    XA_MSG_COMMAND_OUTPUT_ADD_FILTER = 24
    XA_MSG_COMMAND_OUTPUT_REMOVE_FILTER = 25
    XA_MSG_COMMAND_OUTPUT_FILTERS_LIST = 26
    XA_MSG_COMMAND_OUTPUT_MODULE_REGISTER = 27
    XA_MSG_COMMAND_OUTPUT_MODULE_QUERY = 28
    XA_MSG_COMMAND_OUTPUT_MODULES_LIST = 29
    XA_MSG_COMMAND_CODEC_SEND_MESSAGE = 30
    XA_MSG_COMMAND_CODEC_MODULE_REGISTER = 31
    XA_MSG_COMMAND_CODEC_MODULE_QUERY = 32
    XA_MSG_COMMAND_CODEC_MODULES_LIST = 33
    XA_MSG_SET_PLAYER_MODE = 34
    XA_MSG_GET_PLAYER_MODE = 35
    XA_MSG_SET_PLAYER_ENVIRONMENT_INTEGER = 36
    XA_MSG_GET_PLAYER_ENVIRONMENT_INTEGER = 37
    XA_MSG_SET_PLAYER_ENVIRONMENT_STRING = 38
    XA_MSG_GET_PLAYER_ENVIRONMENT_STRING = 39
    XA_MSG_UNSET_PLAYER_ENVIRONMENT = 40
    XA_MSG_SET_INPUT_NAME = 41
    XA_MSG_GET_INPUT_NAME = 42
    XA_MSG_SET_INPUT_MODULE = 43
    XA_MSG_GET_INPUT_MODULE = 44
    XA_MSG_SET_INPUT_POSITION_RANGE = 45
    XA_MSG_GET_INPUT_POSITION_RANGE = 46
    XA_MSG_SET_INPUT_TIMECODE_GRANULARITY = 47
    XA_MSG_GET_INPUT_TIMECODE_GRANULARITY = 48
    XA_MSG_SET_OUTPUT_NAME = 49
    XA_MSG_GET_OUTPUT_NAME = 50
    XA_MSG_SET_OUTPUT_MODULE = 51
    XA_MSG_GET_OUTPUT_MODULE = 52
    XA_MSG_SET_OUTPUT_POSITION_RANGE = 53
    XA_MSG_GET_OUTPUT_POSITION_RANGE = 54
    XA_MSG_SET_OUTPUT_TIMECODE_GRANULARITY = 55
    XA_MSG_GET_OUTPUT_TIMECODE_GRANULARITY = 56
    XA_MSG_SET_OUTPUT_VOLUME = 57
    XA_MSG_GET_OUTPUT_VOLUME = 58
    XA_MSG_SET_OUTPUT_CHANNELS = 59
    XA_MSG_GET_OUTPUT_CHANNELS = 60
    XA_MSG_SET_OUTPUT_PORTS = 61
    XA_MSG_GET_OUTPUT_PORTS = 62
    XA_MSG_SET_CODEC_QUALITY = 63
    XA_MSG_GET_CODEC_QUALITY = 64
    XA_MSG_SET_CODEC_EQUALIZER = 65
    XA_MSG_GET_CODEC_EQUALIZER = 66
    XA_MSG_SET_CODEC_MODULE = 67
    XA_MSG_GET_CODEC_MODULE = 68
    XA_MSG_SET_NOTIFICATION_MASK = 69
    XA_MSG_GET_NOTIFICATION_MASK = 70
    XA_MSG_SET_DEBUG_LEVEL = 71
    XA_MSG_GET_DEBUG_LEVEL = 72
    
    'notifications from decoder
    XA_MSG_NOTIFY_READY = 73
    XA_MSG_NOTIFY_ACK = 74
    XA_MSG_NOTIFY_NACK = 75
    XA_MSG_NOTIFY_PONG = 76
    XA_MSG_NOTIFY_EXITED = 77
    XA_MSG_NOTIFY_PLAYER_STATE = 78
    XA_MSG_NOTIFY_PLAYER_MODE = 79
    XA_MSG_NOTIFY_PLAYER_ENVIRONMENT_INTEGER = 80
    XA_MSG_NOTIFY_PLAYER_ENVIRONMENT_STRING = 81
    XA_MSG_NOTIFY_INPUT_STATE = 82
    XA_MSG_NOTIFY_INPUT_NAME = 83
    XA_MSG_NOTIFY_INPUT_CAPS = 84
    XA_MSG_NOTIFY_INPUT_POSITION = 85
    XA_MSG_NOTIFY_INPUT_POSITION_RANGE = 86
    XA_MSG_NOTIFY_INPUT_TIMECODE = 87
    XA_MSG_NOTIFY_INPUT_TIMECODE_GRANULARITY = 88
    XA_MSG_NOTIFY_INPUT_MODULE = 89
    XA_MSG_NOTIFY_INPUT_MODULE_INFO = 90
    XA_MSG_NOTIFY_INPUT_DEVICE_INFO = 91
    XA_MSG_NOTIFY_INPUT_FILTER_INFO = 92
    XA_MSG_NOTIFY_OUTPUT_STATE = 93
    XA_MSG_NOTIFY_OUTPUT_NAME = 94
    XA_MSG_NOTIFY_OUTPUT_CAPS = 95
    XA_MSG_NOTIFY_OUTPUT_POSITION = 96
    XA_MSG_NOTIFY_OUTPUT_POSITION_RANGE = 97
    XA_MSG_NOTIFY_OUTPUT_TIMECODE = 98
    XA_MSG_NOTIFY_OUTPUT_TIMECODE_GRANULARITY = 99
    XA_MSG_NOTIFY_OUTPUT_VOLUME = 100
    XA_MSG_NOTIFY_OUTPUT_BALANCE = 101
    XA_MSG_NOTIFY_OUTPUT_PCM_LEVEL = 102
    XA_MSG_NOTIFY_OUTPUT_MASTER_LEVEL = 103
    XA_MSG_NOTIFY_OUTPUT_CHANNELS = 104
    XA_MSG_NOTIFY_OUTPUT_PORTS = 105
    XA_MSG_NOTIFY_OUTPUT_MODULE = 106
    XA_MSG_NOTIFY_OUTPUT_MODULE_INFO = 107
    XA_MSG_NOTIFY_OUTPUT_DEVICE_INFO = 108
    XA_MSG_NOTIFY_OUTPUT_FILTER_INFO = 109
    XA_MSG_NOTIFY_STREAM_MIME_TYPE = 110
    XA_MSG_NOTIFY_STREAM_DURATION = 111
    XA_MSG_NOTIFY_STREAM_PARAMETERS = 112
    XA_MSG_NOTIFY_STREAM_PROPERTIES = 113
    XA_MSG_NOTIFY_CODEC_QUALITY = 114
    XA_MSG_NOTIFY_CODEC_EQUALIZER = 115
    XA_MSG_NOTIFY_CODEC_MODULE = 116
    XA_MSG_NOTIFY_CODEC_MODULE_INFO = 117
    XA_MSG_NOTIFY_CODEC_DEVICE_INFO = 118
    XA_MSG_NOTIFY_NOTIFICATION_MASK = 119
    XA_MSG_NOTIFY_DEBUG_LEVEL = 120
    XA_MSG_NOTIFY_PROGRESS = 121
    XA_MSG_NOTIFY_DEBUG = 122
    XA_MSG_NOTIFY_ERROR = 123
    XA_MSG_NOTIFY_PRIVATE_DATA = 124
    
    'commands to timesync
    XA_MSG_COMMAND_FEEDBACK_HANDLER_MODULE_REGISTER = 125
    XA_MSG_COMMAND_FEEDBACK_HANDLER_MODULE_QUERY = 126
    XA_MSG_COMMAND_FEEDBACK_HANDLER_MODULES_LIST = 127
    XA_MSG_COMMAND_FEEDBACK_HANDLER_EXIT = 128
    XA_MSG_COMMAND_FEEDBACK_HANDLER_START = 129
    XA_MSG_COMMAND_FEEDBACK_HANDLER_STOP = 130
    XA_MSG_COMMAND_FEEDBACK_HANDLER_PAUSE = 131
    XA_MSG_COMMAND_FEEDBACK_HANDLER_RESTART = 132
    XA_MSG_COMMAND_FEEDBACK_HANDLER_FLUSH = 133
    XA_MSG_COMMAND_FEEDBACK_HANDLER_SEND_MESSAGE = 134
    XA_MSG_COMMAND_FEEDBACK_HANDLER_QUEUE_AUDIO_EVENT = 135
    XA_MSG_COMMAND_FEEDBACK_HANDLER_QUEUE_TAG_EVENT = 136
    XA_MSG_COMMAND_FEEDBACK_HANDLER_QUEUE_TIMECODE_EVENT = 137
    XA_MSG_COMMAND_FEEDBACK_HANDLER_QUEUE_POSITION_EVENT = 138
    XA_MSG_SET_FEEDBACK_AUDIO_EVENT_RATE = 139
    XA_MSG_GET_FEEDBACK_AUDIO_EVENT_RATE = 140
    XA_MSG_SET_FEEDBACK_HANDLER_NAME = 141
    XA_MSG_GET_FEEDBACK_HANDLER_NAME = 142
    XA_MSG_SET_FEEDBACK_HANDLER_MODULE = 143
    XA_MSG_GET_FEEDBACK_HANDLER_MODULE = 144
    XA_MSG_SET_FEEDBACK_HANDLER_ENVIRONMENT_INTEGER = 145
    XA_MSG_GET_FEEDBACK_HANDLER_ENVIRONMENT_INTEGER = 146
    XA_MSG_SET_FEEDBACK_HANDLER_ENVIRONMENT_STRING = 147
    XA_MSG_GET_FEEDBACK_HANDLER_ENVIRONMENT_STRING = 148
    XA_MSG_UNSET_FEEDBACK_HANDLER_ENVIRONMENT = 149
    
    'notifications from timesync
    XA_MSG_NOTIFY_FEEDBACK_AUDIO_EVENT_RATE = 150
    XA_MSG_NOTIFY_FEEDBACK_HANDLER_STATE = 151
    XA_MSG_NOTIFY_FEEDBACK_HANDLER_MODULE = 152
    XA_MSG_NOTIFY_FEEDBACK_HANDLER_MODULE_INFO = 153
    XA_MSG_NOTIFY_FEEDBACK_HANDLER_NAME = 154
    XA_MSG_NOTIFY_FEEDBACK_HANDLER_INFO = 155
    XA_MSG_NOTIFY_FEEDBACK_HANDLER_ENVIRONMENT_INTEGER = 156
    XA_MSG_NOTIFY_FEEDBACK_HANDLER_ENVIRONMENT_STRING = 157
    XA_MSG_NOTIFY_FEEDBACK_AUDIO_EVENT = 158
    XA_MSG_NOTIFY_FEEDBACK_TAG_EVENT = 159
    
    'sentinel
    XA_MSG_LAST = 160
End Enum

'input state
Public Enum InputState
    XA_INPUT_STATE_OPEN = 0
    XA_INPUT_STATE_CLOSED = 1
End Enum

'output state
Public Enum OutputState
    XA_OUTPUT_STATE_OPEN = 0
    XA_OUTPUT_STATE_CLOSED = 1
End Enum


'OUTPUT CHANNELS
Public Enum XA_OutputChannels
    XA_OUTPUT_CHANNELS_STEREO = 0
    XA_OUTPUT_CHANNELS_MONO_LEFT = 1
    XA_OUTPUT_CHANNELS_MONO_RIGHT = 2
    XA_OUTPUT_CHANNELS_MONO_MIX = 3
End Enum


Public Type XA_InputStreamInfo
    mChanged As Long             '0 if the stream information has not changed since the last decoded frame, or non zero if it has
    mLevel As Long               'MPEG syntax level (1 for MPEG1, 2 for MPEG2, 0 for MPEG2.5)
    mLayer As Long               'MPEG layer (1, 2 or 3)
    mBitrate As Long             'MPEG bitrate (in bits per second)
    mFrequency As Long           'MPEG sampling frequency (in Hz)
    mMode As Long                'MPEG mode (0 for stereo, 1 for joint-stereo, 2 for dual-channel, 3 for mono)
    mDuration As Long            'estimated stream duration (in milliseconds)
End Type

Public Type XA_TimeCode
    mH As Long       'hours
    mM As Long       'minutes
    mS As Long       'seconds
    mF As Long       'fractures in 100th of second
End Type

Public Type XA_StatusInfo
    mFrame As Long               'current frame number
    mPosition As Single          'Value between 0.0 and 1.0 giving the relative position in the stream
    mInfo As XA_InputStreamInfo  'input stream structure
    mTimecode As XA_TimeCode     'time code structure
End Type

Public Type XA_AbsoluteTime
    mSeconds As Long
    mMicroseconds As Long
End Type

Public Type XA_EnvironmentInfo
    mName As String
    mInteger As Long
    mString As String
End Type

Public Type XA_TimecodeInfo
    mH As Byte
    mM As Byte
    mS As Byte
    mF As Byte
End Type

Public Type XA_NackInfo
    mCommand As Byte
    mCode As Integer
End Type

Public Type XA_VolumeInfo
    mMasterLevel As Byte
    mPCMLevel As Byte
    mBalance As Byte
End Type

Public Type XA_PositionInfo
    mOffset As Long
    mRange As Long
End Type

Public Type XA_ModuleInfo
    nID As Byte
    mNBDevices As Byte
    mName As String
    mDescription As String
End Type

Public Type XA_FilterInfo
    nID As Integer
    mName As String
End Type

Public Type XA_DeviceInfo
    mModuleID As Byte
    mIndex As Byte
    mFlags As Byte
    mName As String
    mDescription As String
End Type

Public Type XA_StreamParameters
    mFrequency As Long
    mBitrate As Integer
    mNBChannels As Byte
End Type

Public Type XA_ModuleMessage
    mType As Integer
    mSize As Long
    mData As Long
End Type

Public Type XA_TagEvent
    mWhen As XA_AbsoluteTime
    mTag As Long
End Type

Public Type XA_AudioEvent
    mWhen As XA_AbsoluteTime
    mSamplingFrequency As Long
    mNBChannels As Integer
    mNBSamples As Integer
    mSamples As String
End Type

Public Type XA_TimecodeEvent
    mWhen As XA_AbsoluteTime
    mTimecode As XA_TimecodeInfo
End Type

Public Type XA_PositionEvent
    mWhen As XA_AbsoluteTime
    mPosition As XA_PositionInfo
End Type

Public Type XA_ProgressInfo
    mSource As Byte
    mCode As Byte
    mValue As Integer
    mMessage As String
End Type

Public Type XA_DebugInfo
    mSource As Byte
    mLevel As Byte
    mMessage As String
End Type

Public Type XA_ErrorInfo
    mSource As Byte
    mCode As Integer
    mMessage As String
End Type

Public Type XA_PrivateData
    mSource As Byte
    mType As Integer
    mData As Long
    mSize As Long
End Type

Public Type XA_EqualizerInfo
    eleft(XA_DECODER_EQUALIZER_NB_BANDS) As Byte
    eright(XA_DECODER_EQUALIZER_NB_BANDS) As Byte
End Type

Public Enum XA_PropertyType
    XA_PROPERTY_TYPE_STRING
    XA_PROPERTY_TYPE_INTEGER
End Enum

Public Type XA_PropertyValue
    mInteger As Long
    mString As String
End Type

Public Type XA_Property
    mName As String
    mType As XA_PropertyType
    mValue As XA_PropertyValue
End Type

Public Type XA_PropertyList
    mNBProperties As Long
    mProperties As XA_Property
End Type

'INPUT/OUTPUT STATE
Public Enum XA_InputOutputState
    XA_STATE_OPEN = 0
    XA_STATE_CLOSED = 1
End Enum

'PLAYER STATE
Public Enum XA_PlayerState
    XA_PLAYER_STATE_STOPPED = 0
    XA_PLAYER_STATE_PLAYING = 1
    XA_PLAYER_STATE_PAUSED = 2
    XA_PLAYER_STATE_EOS = 3
End Enum

Public Type XA_Message
    mCode As XA_MessageCode
    'data structure follows
    mBuffer As String
    mName As String
    mString As String
    mMimeType As String
    mModuleID As Integer
    mMode As Long
    mChannels As Byte
    mQuality As Byte
    mDuration As Long
    mRange As Long
    mGranularity As Long
    mCaps As Long
    mPorts As Byte
    mAck As Byte
    mTag As Long
    mDebugLevel As Byte
    mNotificationMask As Long
    mRate As Byte
    mNack As XA_NackInfo
    mVolume As XA_VolumeInfo
    mPosition As XA_PositionInfo
    mEqualizer As XA_EqualizerInfo
    mModuleinfo As XA_ModuleInfo
    mFilterInfo As XA_FilterInfo
    mDeviceInfo As XA_DeviceInfo
    mStreamParameters As XA_StreamParameters
    mEnvironmentInfo As XA_EnvironmentInfo
    mTimecode As XA_TimecodeInfo
    mModuleMessage As XA_ModuleMessage
    mTagEvent As XA_TagEvent
    mAudioEvent As XA_AudioEvent
    mTimecodeEvent As XA_TimecodeEvent
    mPositionEvent As XA_PositionEvent
    mProperties As XA_PropertyList
    mWhen As XA_AbsoluteTime
    mProgress As XA_ProgressInfo
    mDebug As XA_DebugInfo
    mError As XA_ErrorInfo
    mPrivateData As XA_PrivateData
    mInputState As XA_InputOutputState
    mOutputState As XA_InputOutputState
    mPlayerState As XA_PlayerState
    mInputStreamInfo As XA_InputStreamInfo
    'end of data structure
End Type



'open new player if par=0 then msges will be sent to player, otherwise to specified window handle hWnd
Public Declare Function player_new Lib "xaudio.dll" (hPlayer As Long, par As Long) As Long
'delete specified player
Public Declare Function player_delete Lib "xaudio.dll" (ByVal hPlayer As Long) As Long
Public Declare Function player_set_priority Lib "xaudio.dll" (ByVal hPlayer As Long, ByVal par As Long) As Long
Public Declare Function player_get_priority Lib "xaudio.dll" (ByVal hPlayer As Long) As Long
Public Declare Function control_message_send_S Lib "xaudio.dll" (ByVal hPlayer As Long, ByVal msg_code As Long, ByVal msg_str As String) As Long
Public Declare Function control_message_send_N Lib "xaudio.dll" (ByVal hPlayer As Long, ByVal msg_code As Long) As Long
Public Declare Function control_message_send_I Lib "xaudio.dll" (ByVal hPlayer As Long, ByVal msg_code As Long, ByVal dat As Long) As Long
'Public Declare Function control_message_get Lib "xaudio.dll" (hPlayer As Long, status As XA_Message) As Long
'Public Declare Function control_message_wait Lib "xaudio.dll" (ByVal hPlayer As Long, status As XA_Message, ByVal TimeOut As Long) As Long
'Public Declare Function control_message_sprint Lib "xaudio.dll" (strBuff As String, status As XA_Message) As Long
'Public Declare Function xaudio_get_version Lib "xaudio.dll" (ByVal cc As Long) As Long
Public Declare Function control_message_send_IPI Lib "xaudio.dll" (ByVal hPlayer As Long, ByVal msg_code As Long, ByVal dat1 As Long, ptr As Any, ByVal dat2 As Long) As Long
Public Declare Function control_message_send_II Lib "xaudio.dll" (ByVal hPlayer As Long, ByVal msg_code As Long, ByVal dat1 As Long, ByVal dat2 As Long) As Long
Public Declare Function control_message_send_P Lib "xaudio.dll" (ByVal hPlayer As Long, ByVal msg_code As Long, dat As Any) As Long
Public Declare Function control_message_send_III Lib "xaudio.dll" (ByVal hPlayer As Long, ByVal msg_code As Long, ByVal dat1 As Long, ByVal dat2 As Long, ByVal dat3 As Long) As Long



'FEEDBACK HANDLER STATE
Public Const XA_FEEDBACK_HANDLER_STATE_STARTED = 0
Public Const XA_FEEDBACK_HANDLER_STATE_STOPPED = 1


'ERROR CODES
Public Const XA_SUCCESS = 0
Public Const XA_FAILURE = -1

'Priorities
Public Const XA_CONTROL_PRIORITY_LOWEST = 0
Public Const XA_CONTROL_PRIORITY_LOW = 1
Public Const XA_CONTROL_PRIORITY_NORMAL = 2
Public Const XA_CONTROL_PRIORITY_HIGH = 3
Public Const XA_CONTROL_PRIORITY_HIGHEST = 4

'general error codes
Public Const XA_ERROR_BASE_GENERAL = -100
Public Const XA_ERROR_OUT_OF_MEMORY = XA_ERROR_BASE_GENERAL - 0
Public Const XA_ERROR_OUT_OF_RESOURCES = XA_ERROR_BASE_GENERAL - 1
Public Const XA_ERROR_INVALID_PARAMETERS = XA_ERROR_BASE_GENERAL - 2
Public Const XA_ERROR_INTERNAL = XA_ERROR_BASE_GENERAL - 3
Public Const XA_ERROR_TIMEOUT = XA_ERROR_BASE_GENERAL - 4
Public Const XA_ERROR_VERSION_EXPIRED = XA_ERROR_BASE_GENERAL - 5
Public Const XA_ERROR_VERSION_MISMATCH = XA_ERROR_BASE_GENERAL - 6

'network error codes
Public Const XA_ERROR_BASE_NETWORK = -200
Public Const XA_ERROR_CONNECT_TIMEOUT = XA_ERROR_BASE_NETWORK - 0
Public Const XA_ERROR_CONNECT_FAILED = XA_ERROR_BASE_NETWORK - 1
Public Const XA_ERROR_CONNECTION_REFUSED = XA_ERROR_BASE_NETWORK - 2
Public Const XA_ERROR_ACCEPT_FAILED = XA_ERROR_BASE_NETWORK - 3
Public Const XA_ERROR_LISTEN_FAILED = XA_ERROR_BASE_NETWORK - 4
Public Const XA_ERROR_SOCKET_FAILED = XA_ERROR_BASE_NETWORK - 5
Public Const XA_ERROR_SOCKET_CLOSED = XA_ERROR_BASE_NETWORK - 6
Public Const XA_ERROR_BIND_FAILED = XA_ERROR_BASE_NETWORK - 7
Public Const XA_ERROR_HOST_UNKNOWN = XA_ERROR_BASE_NETWORK - 8
Public Const XA_ERROR_HTTP_INVALID_REPLY = XA_ERROR_BASE_NETWORK - 9
Public Const XA_ERROR_HTTP_ERROR_REPLY = XA_ERROR_BASE_NETWORK - 10
Public Const XA_ERROR_HTTP_FAILURE = XA_ERROR_BASE_NETWORK - 11
Public Const XA_ERROR_FTP_INVALID_REPLY = XA_ERROR_BASE_NETWORK - 12
Public Const XA_ERROR_FTP_ERROR_REPLY = XA_ERROR_BASE_NETWORK - 13
Public Const XA_ERROR_FTP_FAILURE = XA_ERROR_BASE_NETWORK - 14

'control error codes
Public Const XA_ERROR_BASE_CONTROL = -300
Public Const XA_ERROR_PIPE_FAILED = XA_ERROR_BASE_CONTROL - 0
Public Const XA_ERROR_FORK_FAILED = XA_ERROR_BASE_CONTROL - 1
Public Const XA_ERROR_SELECT_FAILED = XA_ERROR_BASE_CONTROL - 2
Public Const XA_ERROR_PIPE_CLOSED = XA_ERROR_BASE_CONTROL - 3
Public Const XA_ERROR_PIPE_READ_FAILED = XA_ERROR_BASE_CONTROL - 4
Public Const XA_ERROR_PIPE_WRITE_FAILED = XA_ERROR_BASE_CONTROL - 5
Public Const XA_ERROR_INVALID_MESSAGE = XA_ERROR_BASE_CONTROL - 6
Public Const XA_ERROR_CIRQ_FULL = XA_ERROR_BASE_CONTROL - 7
Public Const XA_ERROR_POST_FAILED = XA_ERROR_BASE_CONTROL - 8

'url error codes
Public Const XA_ERROR_BASE_URL = -400
Public Const XA_ERROR_URL_UNSUPPORTED_SCHEME = XA_ERROR_BASE_URL - 0
Public Const XA_ERROR_URL_INVALID_SYNTAX = XA_ERROR_BASE_URL - 1

'i/o error codes
Public Const XA_ERROR_BASE_IO = -500
Public Const XA_ERROR_OPEN_FAILED = XA_ERROR_BASE_IO - 0
Public Const XA_ERROR_CLOSE_FAILED = XA_ERROR_BASE_IO - 1
Public Const XA_ERROR_READ_FAILED = XA_ERROR_BASE_IO - 2
Public Const XA_ERROR_WRITE_FAILED = XA_ERROR_BASE_IO - 3
Public Const XA_ERROR_PERMISSION_DENIED = XA_ERROR_BASE_IO - 4
Public Const XA_ERROR_NO_DEVICE = XA_ERROR_BASE_IO - 5
Public Const XA_ERROR_IOCTL_FAILED = XA_ERROR_BASE_IO - 6
Public Const XA_ERROR_MODULE_NOT_FOUND = XA_ERROR_BASE_IO - 7
Public Const XA_ERROR_UNSUPPORTED_INPUT = XA_ERROR_BASE_IO - 8
Public Const XA_ERROR_UNSUPPORTED_OUTPUT = XA_ERROR_BASE_IO - 9
Public Const XA_ERROR_UNSUPPORTED_FORMAT = XA_ERROR_BASE_IO - 10
Public Const XA_ERROR_DEVICE_BUSY = XA_ERROR_BASE_IO - 11
Public Const XA_ERROR_NO_SUCH_DEVICE = XA_ERROR_BASE_IO - 12
Public Const XA_ERROR_NO_SUCH_FILE = XA_ERROR_BASE_IO - 13
Public Const XA_ERROR_INPUT_EOS = XA_ERROR_BASE_IO - 14

'codec error codes
Public Const XA_ERROR_BASE_CODEC = -600
Public Const XA_ERROR_NO_CODEC = XA_ERROR_BASE_CODEC - 0

'bitstream error codes
Public Const XA_ERROR_BASE_BITSTREAM = -700
Public Const XA_ERROR_INVALID_FRAME = XA_ERROR_BASE_BITSTREAM - 0

'dynamic linking error codes
Public Const XA_ERROR_BASE_DYNLINK = -800
Public Const XA_ERROR_DLL_NOT_FOUND = XA_ERROR_BASE_DYNLINK - 0
Public Const XA_ERROR_SYMBOL_NOT_FOUND = XA_ERROR_BASE_DYNLINK - 1

'environment variables / porperties  error codes
Public Const XA_ERROR_BASE_ENVIRONMENT = -900
Public Const XA_ERROR_NO_SUCH_ENVIRONMENT = XA_ERROR_BASE_ENVIRONMENT - 0
Public Const XA_ERROR_NO_SUCH_PROPERTY = XA_ERROR_BASE_ENVIRONMENT - 0
Public Const XA_ERROR_ENVIRONMENT_TYPE_MISMATCH = XA_ERROR_BASE_ENVIRONMENT - 1
Public Const XA_ERROR_PROPERTY_TYPE_MISMATCH = XA_ERROR_BASE_ENVIRONMENT - 1

'modules
Public Const XA_ERROR_BASE_MODULES = -1000
Public Const XA_ERROR_NO_SUCH_INTERFACE = XA_ERROR_BASE_MODULES - 0

Public Const OT_COMMAND = 1000&

Public Function SendCommand(cmnd As Long, ByVal wpar As Long, ByVal lpar As Long) As Boolean
Dim StartTime As Long

SendMessage gHW, cmnd, wpar, lpar

StartTime = GetTickCount
Do
    DoEvents
    If GetTickCount - StartTime > OT_COMMAND Then
        SendCommand = False
        Exit Function
    End If
    If ReceivedMsg.mAck = IssuedComm Then Exit Do
    If ReceivedMsg.mError.mCode > 0 Then
        MsgBox ReceivedMsg.mError.mMessage, vbCritical + vbOKOnly, "MP3Player"
        Exit Function
    End If
Loop
SendCommand = True

End Function
