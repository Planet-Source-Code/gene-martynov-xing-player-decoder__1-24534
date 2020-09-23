VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEqualizer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Equalizer"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   154
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   328
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   90
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   315
      TabIndex        =   16
      Top             =   0
      Width           =   4725
      Begin VB.Line linV 
         BorderColor     =   &H00008000&
         Index           =   6
         X1              =   136
         X2              =   136
         Y1              =   8
         Y2              =   40
      End
      Begin VB.Line linV 
         BorderColor     =   &H00008000&
         Index           =   5
         X1              =   120
         X2              =   120
         Y1              =   8
         Y2              =   40
      End
      Begin VB.Line linV 
         BorderColor     =   &H00008000&
         Index           =   4
         X1              =   104
         X2              =   104
         Y1              =   8
         Y2              =   40
      End
      Begin VB.Line linV 
         BorderColor     =   &H00008000&
         Index           =   3
         X1              =   88
         X2              =   88
         Y1              =   8
         Y2              =   40
      End
      Begin VB.Line linV 
         BorderColor     =   &H00008000&
         Index           =   2
         X1              =   72
         X2              =   72
         Y1              =   8
         Y2              =   40
      End
      Begin VB.Line linV 
         BorderColor     =   &H00008000&
         Index           =   1
         X1              =   56
         X2              =   56
         Y1              =   8
         Y2              =   40
      End
      Begin VB.Line linV 
         BorderColor     =   &H00008000&
         Index           =   0
         X1              =   40
         X2              =   40
         Y1              =   8
         Y2              =   40
      End
      Begin VB.Line lin0 
         BorderColor     =   &H00008000&
         X1              =   8
         X2              =   280
         Y1              =   24
         Y2              =   24
      End
   End
   Begin MSComctlLib.Slider slEqualizer 
      Height          =   1095
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1931
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   10
      Min             =   -128
      Max             =   127
      TickFrequency   =   25
   End
   Begin MSComctlLib.Slider slEqualizer 
      Height          =   1095
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1931
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   10
      Min             =   -128
      Max             =   127
      TickFrequency   =   25
   End
   Begin MSComctlLib.Slider slEqualizer 
      Height          =   1095
      Index           =   2
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1931
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   10
      Min             =   -128
      Max             =   127
      TickFrequency   =   25
   End
   Begin MSComctlLib.Slider slEqualizer 
      Height          =   1095
      Index           =   3
      Left            =   2040
      TabIndex        =   6
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1931
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   10
      Min             =   -128
      Max             =   127
      TickFrequency   =   25
   End
   Begin MSComctlLib.Slider slEqualizer 
      Height          =   1095
      Index           =   4
      Left            =   2640
      TabIndex        =   8
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1931
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   10
      Min             =   -128
      Max             =   127
      TickFrequency   =   25
   End
   Begin MSComctlLib.Slider slEqualizer 
      Height          =   1095
      Index           =   5
      Left            =   3240
      TabIndex        =   10
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1931
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   10
      Min             =   -128
      Max             =   127
      TickFrequency   =   25
   End
   Begin MSComctlLib.Slider slEqualizer 
      Height          =   1095
      Index           =   6
      Left            =   3840
      TabIndex        =   12
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1931
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   10
      Min             =   -128
      Max             =   127
      TickFrequency   =   25
   End
   Begin MSComctlLib.Slider slEqualizer 
      Height          =   1095
      Index           =   7
      Left            =   4440
      TabIndex        =   14
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1931
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   10
      Min             =   -128
      Max             =   127
      TickFrequency   =   25
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   8
      X2              =   320
      Y1              =   83
      Y2              =   83
   End
   Begin VB.Label lblBand 
      Alignment       =   2  'Center
      Caption         =   "16kHz"
      Height          =   255
      Index           =   7
      Left            =   4320
      TabIndex        =   15
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblBand 
      Alignment       =   2  'Center
      Caption         =   "8kHz"
      Height          =   255
      Index           =   6
      Left            =   3720
      TabIndex        =   13
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblBand 
      Alignment       =   2  'Center
      Caption         =   "4kHz"
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   11
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblBand 
      Alignment       =   2  'Center
      Caption         =   "2kHz"
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   9
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblBand 
      Alignment       =   2  'Center
      Caption         =   "1kHz"
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   7
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblBand 
      Alignment       =   2  'Center
      Caption         =   "500Hz"
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   5
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblBand 
      Alignment       =   2  'Center
      Caption         =   "256Hz"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblBand 
      Alignment       =   2  'Center
      Caption         =   "128Hz"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   495
   End
End
Attribute VB_Name = "frmEqualizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CalculatePicture()
Dim ph As Long 'picture height
Dim bw As Long 'distance between two bands
Dim i As Long, j As Long
Dim y As Single
Dim Y2 As Long, Y1 As Long
Const PI = 3.1416
Dim A As Single
Dim sh As Single
Dim phase As Single
Dim c As Long

ph = Picture1.Height - 1
bw = Picture1.Width / 7

Picture1.Cls
For i = 0 To 6
    Y1 = slEqualizer(i).Value
    Y2 = slEqualizer(i + 1).Value
    If Y1 > Y2 Then
        phase = 0
        A = (Y1 - Y2) / 2
        sh = Y1 - A
    Else
        phase = PI
        A = (Y2 - Y1) / 2
        sh = Y1 + A
    End If
    For j = 0 To bw - 1
        y = A * Cos(PI * j / bw + phase) + sh
        Select Case Abs(y)
        Case Is > 120
            c = RGB(255, 0, 0)
        Case Is > 110
            c = RGB(192, 64, 64)
        Case Is > 100
            c = RGB(128, 96, 96)
        Case Is > 90
            c = RGB(64, 128, 128)
        Case Else
            c = RGB(0, 128, 128)
        End Select
        y = (127 - y) * ph / 255
        Picture1.PSet (i * bw + j, y), c
    Next j
Next i

End Sub

Private Sub Form_Load()
Dim i As Long

For i = 0 To 6
    linV(i).Y1 = 0
    linV(i).Y2 = Picture1.Height
    linV(i).X1 = (i + 1) * Picture1.Width / 8
    linV(i).X2 = linV(i).X1
Next i
lin0.X1 = 0
lin0.X2 = Picture1.ScaleWidth
lin0.Y1 = (Picture1.ScaleHeight) / 2
lin0.Y2 = lin0.Y1
frmEqualizer.Top = frmMain.Top + frmMain.Height
frmEqualizer.Left = frmMain.Left
CalculatePicture

End Sub

Private Sub slEqualizer_Change(Index As Integer)
Dim dat As Integer

dat = (slEqualizer(Index).Value)
'the deal is:
'   -1 is the same as 255, -2 is 254 and so on
If dat < 0 Then
    dat = 256 + dat
End If
Select Case Index
Case 0
    EqData.eleft(0) = dat
    EqData.eright(0) = dat
Case 1
    EqData.eleft(1) = dat
    EqData.eright(1) = dat
Case 2
    EqData.eleft(2) = dat
    EqData.eright(2) = dat
    EqData.eleft(3) = dat
    EqData.eright(3) = dat
Case 3
    EqData.eleft(4) = dat
    EqData.eright(4) = dat
    EqData.eleft(5) = dat
    EqData.eright(5) = dat
    EqData.eleft(6) = dat
    EqData.eright(6) = dat
Case 4
    EqData.eleft(7) = dat
    EqData.eright(7) = dat
    EqData.eleft(8) = dat
    EqData.eright(8) = dat
    EqData.eleft(9) = dat
    EqData.eright(9) = dat
    EqData.eleft(10) = dat
    EqData.eright(10) = dat
Case 5
    EqData.eleft(11) = dat
    EqData.eright(11) = dat
    EqData.eleft(12) = dat
    EqData.eright(12) = dat
    EqData.eleft(13) = dat
    EqData.eright(13) = dat
    EqData.eleft(14) = dat
    EqData.eright(14) = dat
    EqData.eleft(15) = dat
    EqData.eright(15) = dat
    EqData.eleft(16) = dat
    EqData.eright(16) = dat
    EqData.eleft(17) = dat
    EqData.eright(17) = dat
Case 6
    EqData.eleft(18) = dat
    EqData.eright(18) = dat
    EqData.eleft(19) = dat
    EqData.eright(19) = dat
    EqData.eleft(20) = dat
    EqData.eright(20) = dat
    EqData.eleft(21) = dat
    EqData.eright(21) = dat
    EqData.eleft(22) = dat
    EqData.eright(22) = dat
    EqData.eleft(23) = dat
    EqData.eright(23) = dat
    EqData.eleft(24) = dat
    EqData.eright(24) = dat
Case 7
    EqData.eleft(25) = dat
    EqData.eright(25) = dat
    EqData.eleft(26) = dat
    EqData.eright(26) = dat
    EqData.eleft(27) = dat
    EqData.eright(27) = dat
    EqData.eleft(28) = dat
    EqData.eright(28) = dat
    EqData.eleft(29) = dat
    EqData.eright(29) = dat
    EqData.eleft(30) = dat
    EqData.eright(30) = dat
    EqData.eleft(31) = dat
    EqData.eright(31) = dat
End Select
IssuedComm = XA_MSG_SET_CODEC_EQUALIZER
SendCommand IssuedComm, agGetAddressForObject(EqData), 0
'IssuedComm = XA_MSG_GET_CODEC_EQUALIZER
'SendCommand IssuedComm, 0, 0
CalculatePicture

End Sub

