Attribute VB_Name = "modEncoder"
Option Explicit

Public Type MP3_CONFIG
    dwSampleRate As Long    '48000, 44100, 32000 allowed
    byMode As Byte          'can be 0-stereo, 1-joint stereo, 2-dual channel, 3-mono
    wBitRate As Integer     '32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256 and 320 allowed
    bPrivate As Long        'true-private, false-not private
    bCRC As Long
    bCopyright As Long
    bOriginal As Long
End Type
Private ddd As Long

Public Enum LAME_QUALITY_PRESET
    LQP_NOPRESET& = -1

    ' STANDARD QUALITY PRESETS
    LQP_NORMAL_QUALITY& = 0
    LQP_LOW_QUALITY&
    LQP_HIGH_QUALITY&
    LQP_VOICE_QUALITY&

    LQP_PHONE& = 1000
    LQP_SW& = 2000
    LQP_AM& = 3000
    LQP_FM& = 4000
    LQP_VOICE& = 5000
    LQP_RADIO& = 6000
    LQP_TAPE& = 7000
    LQP_HIFI& = 8000
    LQP_CD& = 9000
    LQP_STUDIO& = 10000
End Enum

Public Type LHV1_CONFIG
    'structure info
    dwStructureVersion As Long  'current version 1
    dwStructureSize As Long     'current size 331 bytes
    'basic encoder settings
    dwSampleRate As Long        'sample rate of input file
    dwResampleRate As Long      'downsampple rate 0-encoder decides
    nMode As Long               '0-stereo, 1-joint stereo, 2-dual channel, 3-mono
    dwBitRate As Long           'CBR bitrate, VBR min rate
    dwMaxBitRate As Long        'CBR ignored, VBR max bit rate
    nPreset As Long
    dwMpegVersion As Long       'for future use
    dwPsyModel As Long          'for future, set to 0
    dwEmphasis As Long          'for future, set to 0
    ' BIT STREAM SETTINGS
    bPrivate As Long            'set private bit
    bCRC As Long                'insert CRC
    bCopyright As Long          'set copyright bit
    bOriginal As Long           'set original bit
    ' VBR STUFF
    bWriteVBRHeader As Long     'write Xing VBR Header
    bEnableVBR As Long          'use VBR encoding
    nVBRQuality As Long         'VBR quality 0 - 9
    dwVBRAbr_bps As Long        'use ABR instead of nVBRQuality
    bNoRes As Long              'disable bit reservoir
    
    btReserved(1 To 255 - 2 * Len(ddd) - 1) As Byte    'future use, set to 0
End Type
Dim bb

Public Type AAC_CONFIG
    dwSampleRate As Long
    byMode As Byte
    wBitRate As Integer
    byEncodingMethod As Byte
End Type

Public Type tFORMAT
    mp3 As MP3_CONFIG
'    lhv1 As LHV1_CONFIG
    aac As AAC_CONFIG
End Type

Public Type BE_CONFIG
    dwConfig As Long
    fform As tFORMAT
End Type

Public Declare Function beInitStream Lib "bladeenc.dll" (pbeConfig As BE_CONFIG, dwSamples As Long, dwBufferSize As Long, phbeStream As Long) As Long
Public Declare Function beEncodeChunk Lib "bladeenc.dll" (ByVal hbeStream As Long, ByVal nSamples As Long, pSamples As Integer, pOutput As Byte, pdwOutput As Long) As Long
Public Declare Function beDeinitStream Lib "bladeenc.dll" (ByVal hbeStream As Long, pOutput As Byte, pdwOutput As Long) As Long


