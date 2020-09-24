VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   Caption         =   "send"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4275
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   389
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   285
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   1635
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   4110
      Width           =   4185
   End
   Begin VB.ListBox lstACMStatus 
      BackColor       =   &H8000000F&
      Height          =   1560
      IntegralHeight  =   0   'False
      ItemData        =   "Main.frx":0442
      Left            =   45
      List            =   "Main.frx":0444
      TabIndex        =   7
      Top             =   2475
      Width           =   4185
   End
   Begin VB.TextBox txtWaveLength 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   900
      TabIndex        =   5
      Text            =   "8192"
      Top             =   360
      Width           =   3330
   End
   Begin VB.TextBox txtHost 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   900
      TabIndex        =   4
      Text            =   "127.0.0.1"
      Top             =   45
      Width           =   3330
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "&Record"
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   675
      Width           =   2055
   End
   Begin VB.ListBox lstStatus 
      BackColor       =   &H8000000F&
      Height          =   1020
      IntegralHeight  =   0   'False
      Left            =   45
      TabIndex        =   1
      Top             =   1215
      Width           =   4185
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   3825
      Top             =   765
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   285
      Left            =   45
      TabIndex        =   0
      Top             =   675
      Width           =   2055
   End
   Begin VB.Label lbl 
      Caption         =   "ACM Status:"
      Height          =   255
      Index           =   3
      Left            =   45
      TabIndex        =   9
      Top             =   2250
      Width           =   1395
   End
   Begin VB.Label lbl 
      Caption         =   "Connection Status:"
      Height          =   255
      Index           =   2
      Left            =   45
      TabIndex        =   8
      Top             =   990
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Buffer:"
      Height          =   255
      Index           =   1
      Left            =   45
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Host:"
      Height          =   255
      Index           =   0
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bRecord As Boolean, bSent As Boolean, hWaveIn As Long
Dim wave() As Byte, waveout() As Byte, dwWaveLength As Long
Dim InFormat As WAVEFORMATEX, OutFormat As WAVEFORMATEXBIG
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Sub cmdConnect_Click()
    ws.Close
    ws.Connect txtHost.Text, 4299
End Sub
Private Sub cmdRecord_Click()
    If (Not bRecord) Then
        bRecord = True
        cmdRecord.Caption = "&Stop"
        startRecord
    Else
        bRecord = False
        cmdRecord.Caption = "&Record"
    End If
End Sub
Private Sub Form_Load()
    txtWaveLength_LostFocus
    'sound format structures
    InFormat.wFormatTag = WAVE_FORMAT_PCM
    InFormat.nSamplesPerSec = 8000
    InFormat.wBitsPerSample = 16
    InFormat.nChannels = 1
    InFormat.nBlockAlign = InFormat.wBitsPerSample * InFormat.nChannels / 8
    InFormat.nAvgBytesPerSec = InFormat.nBlockAlign * InFormat.nSamplesPerSec
    'for PCM do not fill out InFormat.cbSize
    'hgkhgkhgkhgkhgkhg
    OutFormat.wFormatTag = WAVE_FORMAT_PCM
    OutFormat.nSamplesPerSec = 8000
    OutFormat.wBitsPerSample = 16
    OutFormat.nChannels = 1
    OutFormat.nBlockAlign = InFormat.wBitsPerSample * InFormat.nChannels / 8
    OutFormat.nAvgBytesPerSec = InFormat.nBlockAlign * InFormat.nSamplesPerSec
    msacm32.ChooseFormat InFormat, OutFormat
    addACMStatus ("InFormat:")
    addACMStatus ("InFormat.wFormatTag=" & InFormat.wFormatTag)
    addACMStatus ("InFormat.nSamplesPerSec=" & InFormat.nSamplesPerSec)
    addACMStatus ("InFormat.wBitsPerSample=" & InFormat.wBitsPerSample)
    addACMStatus ("InFormat.nChannels=" & InFormat.nChannels)
    addACMStatus ("InFormat.nBlockAlign=" & InFormat.nBlockAlign)
    addACMStatus ("InFormat.nAvgBytesPerSec=" & InFormat.nAvgBytesPerSec)
    addACMStatus ("InFormat.cbSize=" & InFormat.cbSize)
    addACMStatus ("OutFormat:")
    addACMStatus ("OutFormat.wFormatTag=" & OutFormat.wFormatTag)
    addACMStatus ("OutFormat.nSamplesPerSec=" & OutFormat.nSamplesPerSec)
    addACMStatus ("OutFormat.wBitsPerSample=" & OutFormat.wBitsPerSample)
    addACMStatus ("OutFormat.nChannels=" & OutFormat.nChannels)
    addACMStatus ("OutFormat.nBlockAlign=" & OutFormat.nBlockAlign)
    addACMStatus ("OutFormat.nAvgBytesPerSec=" & OutFormat.nAvgBytesPerSec)
    addACMStatus ("OutFormat.cbSize=" & OutFormat.cbSize)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    waveInReset hWaveIn
    waveInClose hWaveIn
End Sub
Private Sub startRecord()
    Dim wvhdr As WAVEHDR, i As Integer
    addStatus "Start recording."
    waveInOpen hWaveIn, WAVE_MAPPER, InFormat, 0, 0, 0
    wvhdr.lpData = VarPtr(wave(0))
    wvhdr.dwBufferLength = dwWaveLength
    While (bRecord)
        waveInPrepareHeader hWaveIn, wvhdr, Len(wvhdr)
        waveInAddBuffer hWaveIn, wvhdr, Len(wvhdr)
        waveInStart hWaveIn
        If (wvhdr.dwBytesRecorded > 0) Then
            ReDim waveout(0 To (wvhdr.dwBytesRecorded - 1)) As Byte
            For i = 0 To (wvhdr.dwBytesRecorded - 1)
                waveout(i) = wave(i)
            Next i
            bSent = False
            msacm32.ConvertFormat waveout(), UBound(waveout()) + 1, VarPtr(InFormat), VarPtr(OutFormat)
            ws.SendData waveout
            While (Not bSent)
                DoEvents
            Wend
        End If
        DoEvents
        waveInStop hWaveIn
        waveInUnprepareHeader hWaveIn, wvhdr, Len(wvhdr)
    Wend
    waveInReset hWaveIn
    waveInClose hWaveIn
    addStatus "Stop recording."
    ws.Close
    cmdConnect_Click
End Sub
Private Sub addStatus(ByRef strLine As String)
    lstStatus.AddItem strLine
    lstStatus.TopIndex = lstStatus.NewIndex
End Sub
Private Sub lstACMStatus_Click()
    Text1.Text = Text1.Text & lstACMStatus.Text & vbNewLine
End Sub
Private Sub txtWaveLength_LostFocus()
    dwWaveLength = CLng(txtWaveLength.Text)
    ReDim wave(0 To (dwWaveLength - 1)) As Byte
End Sub
Private Sub ws_Connect()
    Dim output() As Byte
    addStatus "Connected to " & ws.RemoteHostIP & ":" & ws.RemotePort
    ReDim output(0 To (Len(InFormat) - 1)) As Byte
    CopyMemory VarPtr(output(0)), VarPtr(InFormat), Len(InFormat)
    ws.SendData output
    ReDim output(0 To (Len(OutFormat) - 1)) As Byte
    CopyMemory VarPtr(output(0)), VarPtr(OutFormat), Len(OutFormat)
    ws.SendData output
End Sub
Private Sub ws_Close()
    ws.Close
    addStatus "Socket closed"
End Sub
Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    addStatus "Winsock Error: " & Description
End Sub
Private Sub ws_SendComplete()
    bSent = True
    addStatus "Send complete."
End Sub
Private Sub ws_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    bSent = False
    addStatus "Sent " & bytesSent & ", remaining: " & bytesRemaining
End Sub
Public Sub addACMStatus(ByRef strLine As String)
    lstACMStatus.AddItem strLine
    lstACMStatus.TopIndex = lstACMStatus.NewIndex
End Sub
