Attribute VB_Name = "modSubs"
Option Explicit

Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)

'**********************************
'**  Function Declarations:
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessagePointer Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Any, lParam As Any) As Long
Public Declare Function agGetStringFromPointer Lib "apigid32.dll" Alias "agGetStringFromLPSTR" (ByVal ptr As Long) As String
Public Declare Sub agCopyData Lib "apigid32.dll" (ByVal source As Long, dest As Any, ByVal nCount As Long)
Public Declare Function agGetAddressForObject Lib "apigid32.dll" (obj As Any) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

'**********************************
'**  Public variables:
Public hPlayer As Long          'handle to the created player
Public lpPrevWndProc As Long    'address of default Windows procedure
Public gHW As Long              'handle to the window that processes player's messages
Public MP3File As String        'name of opened MP3 file
Public MP3FileShort As String   'short name of MP3
Public WAVFile As String        'name of WAV file to be saved
Public WAVFileShort As String   'short name
Public IssuedComm As Long       'command to send to player

'*************************************
'**  Public var for received messages
'this structure is filled out every time you send a command
'   and during the play/decode. You can retrive a value from
'   this structure to make sure command was executed successfully
Public ReceivedMsg As XA_Message

Public Function GetShortName(fn As String, exist_check As Boolean) As String
Dim ret As Long
Dim st As String

If exist_check Then
    If Dir(fn, vbNormal) = "" Then
        MsgBox "File " & fn & vbNewLine & "does not exist.", vbCritical + vbOKOnly, "MP3Player"
        GetShortName = ""
        Exit Function
    End If
End If
st = String(255, 0)
ret = GetShortPathName(fn, st, 255)
GetShortName = Left(st, ret)

End Function

Public Sub Hook()
'hook the window. Return value is an address of default Windows procedure
lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)

End Sub

Public Sub Unhook()
'Must unhook before exiting the program
Dim temp As Long
temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)

End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'this will take care of sending appropriate command to player or execute
'   default Windows procedure
Dim ret As Long

'Debug.Print hw, uMsg, wParam, lParam, hPlayer

If GetWindowLong(gHW, GWL_USERDATA) = hPlayer Then
    'Go ahead - execute command
    If uMsg = IssuedComm Then
        'clear receivedMsg structure
        ReceivedMsg.mAck = 0
        ReceivedMsg.mNack.mCode = 0
        ReceivedMsg.mNack.mCommand = 0
        'execute command here
        Select Case uMsg
        Case XA_MSG_COMMAND_INPUT_OPEN 'OK
            'we can send an address of the string - but i am just too lazy
            ret = control_message_send_S(hPlayer, uMsg, MP3FileShort)
        Case XA_MSG_COMMAND_INPUT_CLOSE 'OK
            ret = control_message_send_N(hPlayer, uMsg)
        Case XA_MSG_COMMAND_OUTPUT_OPEN 'OK
            'we can send an address of the string - but i am just too lazy
            ret = control_message_send_S(hPlayer, uMsg, WAVFileShort)
        Case XA_MSG_COMMAND_OUTPUT_CLOSE 'OK
            ret = control_message_send_N(hPlayer, uMsg)
        Case XA_MSG_COMMAND_OUTPUT_RESET
            ret = control_message_send_N(hPlayer, uMsg)
        Case XA_MSG_COMMAND_INPUT_SEND_MESSAGE
            'do not use it - EVER!!!!!
        Case XA_MSG_COMMAND_PLAY ' start playing - if applicable. OK
            If ReceivedMsg.mPlayerState = XA_PLAYER_STATE_PLAYING Then Exit Function
            ret = control_message_send_N(hPlayer, uMsg)
        Case XA_MSG_COMMAND_STOP ' stop playing. OK
            If ReceivedMsg.mPlayerState = XA_PLAYER_STATE_EOS Or ReceivedMsg.mPlayerState = XA_PLAYER_STATE_STOPPED Then Exit Function
            ret = control_message_send_N(hPlayer, uMsg)
        Case XA_MSG_COMMAND_EXIT ' close input and output and exit.
            ret = control_message_send_N(hPlayer, uMsg)
        Case XA_MSG_COMMAND_PAUSE 'OK
            'execute only if playing
            If ReceivedMsg.mPlayerState = XA_PLAYER_STATE_PLAYING Then
                ret = control_message_send_N(hPlayer, uMsg)
            Else
                Exit Function
            End If
        Case XA_MSG_COMMAND_SEEK
            'wParam is an offset, lParam is a range
            ret = control_message_send_II(hPlayer, uMsg, wParam, lParam)
        Case XA_MSG_COMMAND_SYNC 'DON'T KNOW HOW TO USE
            'do not use
        Case XA_MSG_SET_INPUT_POSITION_RANGE 'OK
            'by default range=400
            ret = control_message_send_I(hPlayer, uMsg, wParam)
        Case XA_MSG_GET_INPUT_POSITION_RANGE 'OK
            ret = control_message_send_N(hPlayer, uMsg)
        Case XA_MSG_COMMAND_OUTPUT_MUTE 'OK
            'it does not really mute the output
            '   it just sets the volume of wav to 0
            ret = control_message_send_N(hPlayer, uMsg)
        Case XA_MSG_COMMAND_OUTPUT_UNMUTE 'OK
            'recall value of volume, when it was muted
            ret = control_message_send_N(hPlayer, uMsg)
        Case XA_MSG_SET_INPUT_TIMECODE_GRANULARITY 'DOES NOT WORK
            'DOES NOT WORK, BUT DOES NOT HURT EITHER IF GREATER THAN 100
            'by default granularity is 100. It suppose to take care of
            '   the speed of sending update messages
            ret = control_message_send_I(hPlayer, uMsg, wParam)
        Case XA_MSG_GET_INPUT_TIMECODE_GRANULARITY 'OK
            ret = control_message_send_N(hPlayer, uMsg)
        Case XA_MSG_SET_OUTPUT_VOLUME 'OK
            'wparam is divided in three parts:
            '   upper 16 bits - master volume 0-100
            '   bits 8-15 - wav volume 0-100
            '   bits 0-7 - balance left/right 0-100
            Dim mast As Long
            Dim wavvol As Long
            Dim bal As Long
            mast = (wParam) \ (256& * 256& * 256&)
            wavvol = (wParam - mast * 256& * 256& * 256&) \ (256& * 256)
            bal = (wParam - mast * 256& * 256& * 256& - wavvol * 256& * 256&) \ 256&
            If mast > 100 Then mast = 100
            If wavvol > 100 Then wavvol = 100
            If bal > 100 Then bal = 100
            If mast < 0 Then mast = 0
            If wavvol < 0 Then wavvol = 0
            If bal < 0 Then bal = 0
            ret = control_message_send_III(hPlayer, uMsg, bal, wavvol, mast)
        Case XA_MSG_GET_OUTPUT_VOLUME 'OK
            'send back separate data for master, wav and balance
            ret = control_message_send_N(hPlayer, uMsg)
        Case XA_MSG_SET_OUTPUT_CHANNELS 'OK
            'accepts only values from 0 to 4
            'no needs to use this command
            If (wParam >= 0) And (wParam < 5) Then
                ret = control_message_send_I(hPlayer, uMsg, wParam)
            End If
        Case XA_MSG_GET_OUTPUT_CHANNELS 'OK
            ret = control_message_send_N(hPlayer, uMsg)
        Case XA_MSG_SET_CODEC_EQUALIZER 'OK
            'wParam is an address of XA_equalizer structure
            ret = control_message_send_I(hPlayer, uMsg, wParam)
            Debug.Print "Equalizer Set"
        Case XA_MSG_GET_CODEC_EQUALIZER 'OK
            ret = control_message_send_N(hPlayer, uMsg)
        End Select
    ElseIf uMsg > 36864 Then
        ProcessReturn uMsg - 36864, wParam, lParam
        
    'EXECUTE DEFAULT PROCEDURE FOR THE WINDOW
    Else
        WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
    End If
Else
    'message was not send by player redirect it to default Windows procedure
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
    Exit Function
End If

End Function

Public Function ProcessReturn(code As Long, wParam As Long, lParam As Long) As Boolean
'Debug.Print code, wParam, lParam
Select Case code
Case XA_MSG_NOTIFY_ACK 'OK
    'acknowledged - wParam is a code of command that was acknowledged
    ReceivedMsg.mAck = wParam
'    Debug.Print "Ack "; wParam; lParam
Case XA_MSG_NOTIFY_NACK 'OK
    'not acknowledged - wParam - code to be not acjknowledged, lParam - error
    ReceivedMsg.mNack.mCode = lParam
    ReceivedMsg.mNack.mCommand = wParam
'    Debug.Print "Not Ack"; wParam; lParam
Case XA_MSG_NOTIFY_READY 'OK - notify when player is created
    'player ready
Case XA_MSG_NOTIFY_PLAYER_STATE 'OK - state like Playing, stopped, paused, E(nd)O(f)S(tream)
    'returns player state - wParam - code of player state
    ReceivedMsg.mPlayerState = wParam
Case XA_MSG_NOTIFY_INPUT_TIMECODE 'OK - during play/decode player sends this info every second
'   does not make sense to use fraction as it is always 0
'    ReceivedMsg.mTimecode.mF = CSng(wParam And &HFF000000) / (256& * 256& * 256&)
    ReceivedMsg.mTimecode.mS = (wParam And 16711680) / (256& * 256&)
    ReceivedMsg.mTimecode.mM = (wParam And 65280) / 256&
    ReceivedMsg.mTimecode.mH = wParam And 255
Case XA_MSG_NOTIFY_DEBUG 'OK
    'if something happened you will be notified to debug
    ReceivedMsg.mDebug.mMessage = agGetStringFromPointer(lParam)
    ReceivedMsg.mDebug.mSource = wParam Mod (256& * 256&)
    ReceivedMsg.mDebug.mLevel = (wParam And &HFFFF0000) \ (256& * 256&)
    Debug.Print "Debug Received "; wParam; agGetStringFromPointer(lParam)
Case XA_MSG_NOTIFY_INPUT_STATE 'OK
    'sent every time state of input is changed
    ReceivedMsg.mInputState = wParam
Case XA_MSG_NOTIFY_OUTPUT_STATE
    'sent every time state of output is changed
    ReceivedMsg.mOutputState = wParam
Case XA_MSG_NOTIFY_ERROR 'OK
    'difficult to debug
    Debug.Print "Error occured"
    ReceivedMsg.mError.mCode = wParam \ (256& * 256&)
    ReceivedMsg.mError.mSource = wParam Mod (256& * 256&)
    ReceivedMsg.mError.mMessage = agGetStringFromPointer(lParam)
Case XA_MSG_NOTIFY_PONG
    'did not try
    ReceivedMsg.mTag = wParam
Case XA_MSG_NOTIFY_PLAYER_MODE 'OK
    'notifies when mode is changed
    ReceivedMsg.mMode = wParam
Case XA_MSG_NOTIFY_INPUT_NAME 'OK
    'does not know how to use it
    ReceivedMsg.mName = agGetStringFromPointer(lParam)
Case XA_MSG_NOTIFY_INPUT_CAPS 'OK
    'Notifies of the player's current input capabilities.
    'This is typically useful to be notified wether it is possible to seek into an input stream (an disabling a seek bar for instance, if it is not).
    'The flag XA_DECODER_INPUT_SEEKABLE indicates wether it is possible to seek or not.
    ReceivedMsg.mCaps = wParam
Case XA_MSG_NOTIFY_INPUT_POSITION 'OK
    'notifies when position of input is changed
    ReceivedMsg.mPosition.mOffset = wParam
    ReceivedMsg.mPosition.mRange = lParam
Case XA_MSG_NOTIFY_INPUT_POSITION_RANGE 'OK
    'notifies when range is changed - by default the range is 400
    ReceivedMsg.mPosition.mRange = lParam
Case XA_MSG_NOTIFY_INPUT_TIMECODE_GRANULARITY 'OK
    ReceivedMsg.mGranularity = lParam
Case XA_MSG_NOTIFY_STREAM_DURATION 'OK
    'shows it when stream is VBR - estimated duration of stream
    ReceivedMsg.mDuration = wParam
Case XA_MSG_NOTIFY_STREAM_PARAMETERS 'OK
    'shows it when stream is VBR after each chunk of data
    '   and on opening the input file
    ReceivedMsg.mStreamParameters.mNBChannels = CByte(wParam \ (256& * 256&))
    ReceivedMsg.mStreamParameters.mBitrate = CInt(wParam And (255& * 256& + 255&)) 'in kB/s
    ReceivedMsg.mStreamParameters.mFrequency = lParam 'in Hertz
Case XA_MSG_NOTIFY_STREAM_PROPERTIES
    'it fires but don't know how to use
Case XA_MSG_NOTIFY_STREAM_MIME_TYPE
    ReceivedMsg.mMimeType = agGetStringFromPointer(lParam)
Case XA_MSG_NOTIFY_INPUT_MODULE
    'did not try and do not know how to use
    ReceivedMsg.mModuleID = wParam
Case XA_MSG_NOTIFY_INPUT_MODULE_INFO
    'did not try and do not know how to use
    ReceivedMsg.mModuleinfo.nID = CByte(wParam And &HFF)
    ReceivedMsg.mModuleinfo.mNBDevices = CByte((wParam And &HFF00) / 256)
    ReceivedMsg.mModuleinfo.mName = agGetStringFromPointer(lParam)
Case XA_MSG_NOTIFY_OUTPUT_NAME
    'did not try and do not know how to use
    ReceivedMsg.mName = agGetStringFromPointer(lParam)
Case XA_MSG_NOTIFY_OUTPUT_CAPS
    'did not try
    ReceivedMsg.mCaps = wParam
Case XA_MSG_NOTIFY_OUTPUT_VOLUME 'this does not fire ever
    ReceivedMsg.mVolume.mMasterLevel = CByte(wParam And &HFF)
    ReceivedMsg.mVolume.mPCMLevel = CByte((wParam And &HFF00) / 256&)
    ReceivedMsg.mVolume.mBalance = CByte((wParam And &HFF0000) / (256& * 256&))
Case XA_MSG_NOTIFY_OUTPUT_BALANCE 'OK
    ReceivedMsg.mVolume.mBalance = wParam
Case XA_MSG_NOTIFY_OUTPUT_PCM_LEVEL 'OK
    ReceivedMsg.mVolume.mPCMLevel = wParam
Case XA_MSG_NOTIFY_OUTPUT_MASTER_LEVEL 'OK
    ReceivedMsg.mVolume.mMasterLevel = wParam
Case XA_MSG_NOTIFY_OUTPUT_CHANNELS 'OK
    ReceivedMsg.mChannels = CByte(wParam)
Case XA_MSG_NOTIFY_OUTPUT_MODULE_INFO
    'did not try and do not know how to use
    ReceivedMsg.mModuleinfo.nID = CByte(wParam And &HFF)
    ReceivedMsg.mModuleinfo.mNBDevices = CByte((wParam And &HFF00) / 256)
    ReceivedMsg.mModuleinfo.mName = agGetStringFromPointer(lParam)
Case XA_MSG_NOTIFY_CODEC_EQUALIZER 'DOES NOT WORK VERY GOOD
    'I have no idea why I should add 4 to the address
    agCopyData lParam + 4, ReceivedMsg.mEqualizer, Len(ReceivedMsg.mEqualizer)
'    Dim i As Long
'    For i = 0 To 31
'        Debug.Print i, ReceivedMsg.mEqualizer.eleft(i), ReceivedMsg.mEqualizer.eright(i)
'    Next i
Case XA_MSG_NOTIFY_NOTIFICATION_MASK
    'did not try
    ReceivedMsg.mNotificationMask = wParam
Case XA_MSG_NOTIFY_PROGRESS
    'did not try
    ReceivedMsg.mProgress.mSource = CByte(wParam And &HFF)
    ReceivedMsg.mProgress.mCode = CByte((wParam And &HFF00) / 256&)
    ReceivedMsg.mProgress.mValue = CInt((wParam And &HFFFF0000) / (256& * 256&))
    ReceivedMsg.mProgress.mMessage = agGetStringFromPointer(lParam)
End Select

End Function

