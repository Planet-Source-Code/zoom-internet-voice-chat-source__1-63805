VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   Caption         =   "recv"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4470
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
   ScaleHeight     =   254
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   298
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstACMStatus 
      BackColor       =   &H8000000F&
      Height          =   1605
      IntegralHeight  =   0   'False
      Left            =   45
      TabIndex        =   4
      Top             =   2160
      Width           =   4365
   End
   Begin VB.TextBox txtHost 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   900
      TabIndex        =   3
      Text            =   "hp"
      Top             =   45
      Width           =   3510
   End
   Begin VB.ListBox lstStatus 
      BackColor       =   &H8000000F&
      Height          =   1020
      IntegralHeight  =   0   'False
      Left            =   45
      TabIndex        =   1
      Top             =   900
      Width           =   4365
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "&Listen"
      Height          =   285
      Left            =   45
      TabIndex        =   0
      Top             =   360
      Width           =   4395
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   4005
      Top             =   540
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lbl 
      Caption         =   "ACM Status:"
      Height          =   255
      Index           =   2
      Left            =   45
      TabIndex        =   6
      Top             =   1935
      Width           =   1395
   End
   Begin VB.Label lbl 
      Caption         =   "Connection Status:"
      Height          =   255
      Index           =   1
      Left            =   45
      TabIndex        =   5
      Top             =   675
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Host:"
      Height          =   255
      Index           =   0
      Left            =   45
      TabIndex        =   2
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
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Dim hWaveOut As Long
Dim InFormat As WAVEFORMATEX, OutFormat As WAVEFORMATEXBIG
Private Sub Form_Load()
    waveOutOpen hWaveOut, WAVE_MAPPER, InFormat, 0, 0, 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    waveOutClose hWaveOut
End Sub
Private Sub cmdListen_Click()
    ws.Close
    ws.Bind 4299, txtHost.Text
    ws.Listen
    addStatus "Listening on " & ws.LocalIP & ":" & ws.LocalPort
End Sub
Private Sub ws_Connect()
    addStatus "Connected to " & ws.RemoteHostIP & ":" & ws.RemotePort
End Sub
Private Sub ws_ConnectionRequest(ByVal requestID As Long)
    ws.Close
    ws.Accept requestID
    addStatus "Accepted connection."
End Sub
Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    Dim bData() As Byte, wvhdr As WAVEHDR, i As Long
    If (bytesTotal = (Len(InFormat) + Len(OutFormat))) Then
        ReDim bData(0 To (Len(InFormat) - 1)) As Byte
        ws.GetData bData, vbByte, Len(InFormat)
        CopyMemory VarPtr(InFormat), VarPtr(bData(0)), Len(InFormat)
        ReDim bData(0 To (Len(OutFormat) - 1)) As Byte
        ws.GetData bData, vbByte, Len(OutFormat)
        CopyMemory VarPtr(OutFormat), VarPtr(bData(0)), Len(OutFormat)
        addACMStatus ("")
        addACMStatus ("InFormat:")
        addACMStatus ("InFormat.wFormatTag=" & InFormat.wFormatTag)
        addACMStatus ("InFormat.nSamplesPerSec=" & InFormat.nSamplesPerSec)
        addACMStatus ("InFormat.wBitsPerSample=" & InFormat.wBitsPerSample)
        addACMStatus ("InFormat.nChannels=" & InFormat.nChannels)
        addACMStatus ("InFormat.nBlockAlign=" & InFormat.nBlockAlign)
        addACMStatus ("InFormat.nAvgBytesPerSec=" & InFormat.nAvgBytesPerSec)
        addACMStatus ("InFormat.cbSize=" & InFormat.cbSize)
        addACMStatus ("")
        addACMStatus ("OutFormat:")
        addACMStatus ("OutFormat.wFormatTag=" & OutFormat.wFormatTag)
        addACMStatus ("OutFormat.nSamplesPerSec=" & OutFormat.nSamplesPerSec)
        addACMStatus ("OutFormat.wBitsPerSample=" & OutFormat.wBitsPerSample)
        addACMStatus ("OutFormat.nChannels=" & OutFormat.nChannels)
        addACMStatus ("OutFormat.nBlockAlign=" & OutFormat.nBlockAlign)
        addACMStatus ("OutFormat.nAvgBytesPerSec=" & OutFormat.nAvgBytesPerSec)
        addACMStatus ("OutFormat.cbSize=" & OutFormat.cbSize)
        waveOutOpen hWaveOut, WAVE_MAPPER, InFormat, 0, 0, 0
    Else
        addStatus "Received " & bytesTotal & " bytes."
        ReDim bData(0 To (bytesTotal - 1)) As Byte
        ws.GetData bData, vbByte
        bytesTotal = msacm32.ConvertFormat(bData, bytesTotal, VarPtr(OutFormat), VarPtr(InFormat))
        addStatus "Uncompressed " & bytesTotal & " bytes."
        wvhdr.lpData = VarPtr(bData(0))
        wvhdr.dwBufferLength = bytesTotal
        waveOutPrepareHeader hWaveOut, wvhdr, Len(wvhdr)
        waveOutWrite hWaveOut, wvhdr, Len(wvhdr)
        While ((wvhdr.dwFlags And WHDR_DONE) <> WHDR_DONE)
            DoEvents
            If (ws.BytesReceived > 0) Then ws_DataArrival ws.BytesReceived
        Wend
        waveOutUnprepareHeader hWaveOut, wvhdr, Len(wvhdr)
    End If
End Sub
Private Sub ws_Close()
    ws.Close
    addStatus "Socket closed."
    cmdListen_Click
End Sub
Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    addStatus "Winsock Error: " & Description
End Sub
Private Sub addStatus(ByRef strLine As String)
    lstStatus.AddItem strLine
    lstStatus.TopIndex = lstStatus.NewIndex
End Sub
Public Sub addACMStatus(ByRef strLine As String)
    lstACMStatus.AddItem strLine
    lstACMStatus.TopIndex = lstACMStatus.NewIndex
End Sub
