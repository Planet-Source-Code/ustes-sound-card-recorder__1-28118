VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "SoundCard Recorder"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "WAV Options"
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   4215
      Begin VB.OptionButton optFreq 
         Caption         =   "44,100 kHz, 16 Bit, Stereo"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   3735
      End
      Begin VB.OptionButton optFreq 
         Caption         =   "44,100 kHz, 16 Bit, Stereo"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   3735
      End
      Begin VB.OptionButton optFreq 
         Caption         =   "44,100 kHz, 16 Bit, Stereo"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   3735
      End
      Begin VB.OptionButton optFreq 
         Caption         =   "44,100 kHz, 16 Bit, Stereo"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   3735
      End
   End
   Begin MSComDlg.CommonDialog cdg1 
      Left            =   3240
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "WAV Files (*.wav)|*.wav"
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "&Record"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "WAV Info"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   4215
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "bytes"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   16
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "File Size"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2760
         TabIndex        =   15
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Time Recording"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   5
         Top             =   480
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close "
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saving wav file name"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   330
         Left            =   3480
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtFileName 
         Height          =   330
         Left            =   240
         TabIndex        =   1
         Text            =   "d:\radio.wav"
         Top             =   360
         Width           =   3120
      End
      Begin VB.Label Label1 
         Height          =   225
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   2865
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'autor :pramod kumar (tkpramodkumar@yahoo.com)
'after a lot of searching for recording sound and save to wav file
'I found two samples in net
'one is Sabu's example to record sound and play
'second is Damjan's save CD trcks to WAV
'it is combination of this two samples
'Thanx to Sabu's http://www.vbsquare.com/graphics/tip451.html
'Thanx to Damjan http://www.planet-source-code.com/xq/ASP/txtCodeId.2091/lngWId.1/qx/vb/scripts/ShowCode.htm
'(check it with different volume level in Master volume ctrl.)
Private Declare Function mciSendString Lib "winmm.dll" _
                                   Alias "mciSendStringA" _
                                   (ByVal lpstrCommand As String, _
                                   ByVal lpstrReturnString As String, _
                                   ByVal uReturnLength As Long, _
                                   ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" _
                                   Alias "mciGetErrorStringA" _
                                   (ByVal dwError As Long, _
                                   ByVal lpstrBuffer As String, _
                                   ByVal uLength As Long) As Long
Dim RecdTime As Boolean
Private sBits As String
Private sBytes As String
Private sSample As String
Private lSeconds As Long
Private start As Long
Private Function RecordSound(FileName As String) As Boolean
'sound aliased by recsound
    cmdRecord.Enabled = False
    Dim Result&
    Dim errormsg%
    Dim ReturnString As String * 1024
    Dim ErrorString As String * 1024
    Dim mssg As String * 255
    Dim i As Long
    
    Result& = mciSendString("open new Type waveaudio Alias recsound", ReturnString, Len(ReturnString), 0)
    If Not Result& = 0 Then
        errormsg% = mciGetErrorString(Result&, ErrorString, 1024)
        MsgBox ErrorString, 0, "Error"
    End If
    Result& = mciSendString("set recsound time format ms bitspersample " & CInt(sBits) & " channels 2 bytespersec 22500  samplespersec " & sSample, ReturnString, 1024, 0)
    If Not Result& = 0 Then
        errormsg% = mciGetErrorString(Result&, ErrorString, 1024)
        MsgBox ErrorString, 0, "Error"
    End If
    Result& = mciSendString("record  recsound", ReturnString, Len(ReturnString), 0)
    If Not Result& = 0 Then
        errormsg% = mciGetErrorString(Result&, ErrorString, 1024)
        MsgBox ErrorString, 0, "Error"
    End If
    RecdTime = True
    start = Timer
    Do Until Not RecdTime
        WaveStatus
        DoEvents
    Loop
    Result& = mciSendString("save recsound " & FileName, ReturnString, Len(ReturnString), 0)
    If Not Result& = 0 Then
        errormsg% = mciGetErrorString(Result&, ErrorString, 1024)
        MsgBox ErrorString, 0, "Error"
    End If
    Result& = mciSendString("close recsound", ReturnString, 1024, 0)
    If Not Result& = 0 Then
        errormsg% = mciGetErrorString(Result&, ErrorString, 1024)
        MsgBox ErrorString, 0, "Error"
    End If

End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRecord_Click()
If txtFileName <> "" Then
    If Dir(txtFileName) <> "" Then
        Kill (txtFileName)
    End If
    cmdStop.Enabled = True
    lSeconds = 1
    Call RecordSound(txtFileName)
End If
End Sub

Private Sub cmdStop_Click()
If cmdRecord.Enabled = False Then
    RecdTime = False
End If
cmdRecord.Enabled = True
cmdStop.Enabled = False

End Sub

Private Sub Command1_Click()
    cdg1.ShowSave
    txtFileName.Text = cdg1.FileName
End Sub

Private Sub Form_Load()
cmdStop.Enabled = False

optFreq(0).Caption = "44,100 kHz, 16 Bit, Stereo"
optFreq(1).Caption = "44,100 kHz,  8 Bit, Stereo"
optFreq(2).Caption = "48,000 kHz, 16 Bit, Stereo"
optFreq(3).Caption = "48,000 kHz,  8 Bit, Stereo"

sBits = "16"
sBytes = "172000"
sSample = "44100"

optFreq(0).Value = True


txtFileName.Text = App.Path & "\Recorded.wav"
End Sub

Private Sub optFreq_Click(Index As Integer)
    Select Case Index
        Case 0
            sBits = "16"
            sSample = "44100"

        Case 1
            sBits = "8"
            sSample = "44100"
        
        Case 2
            sBits = "16"
            sSample = "48000"

        Case 3
            sBits = "8"
            sSample = "48000"

    End Select
End Sub

Private Sub WaveStatus()
        
    Dim mssg As String * 255
    Dim i As Long
    Dim elapsed As Long
    Dim intSec As Integer
    Dim sngMin As Single
    Dim TotalTime As String
   
    elapsed = Timer - start
    
    If elapsed < 60 Then
        TotalTime = "00:" & Format(elapsed, "00")
    Else
        intSec = elapsed Mod 60
        sngMin = elapsed \ 60
        TotalTime = Format(sngMin, "00") & ":" & Format(intSec, "00")
    End If
    
    lblTime.Caption = TotalTime

    i = mciSendString("set recsound time format bytes", 0&, 0, 0)
    If i <> 0 Then RecdTime = False
    
    i = mciSendString("status recsound length", mssg, 255, 0)
    If i <> 0 Then RecdTime = False
    
    mssg = CStr(CLng(mssg) / 1024)
    lblSize.Caption = Format(Str(mssg), "######.00") & " kb"
        
End Sub
