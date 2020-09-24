Attribute VB_Name = "mmsystem"
Option Explicit
Public Const MAXPNAMELEN = 32  '  max product name length (including NULL)
Public Const WAVE_ALLOWSYNC = &H2
Public Const WAVE_FORMAT_1M08 = &H1              '  11.025 kHz, Mono,   8-bit
Public Const WAVE_FORMAT_1M16 = &H4              '  11.025 kHz, Mono,   16-bit
Public Const WAVE_FORMAT_1S08 = &H2              '  11.025 kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_1S16 = &H8              '  11.025 kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_2M08 = &H10             '  22.05  kHz, Mono,   8-bit
Public Const WAVE_FORMAT_2M16 = &H40             '  22.05  kHz, Mono,   16-bit
Public Const WAVE_FORMAT_2S08 = &H20             '  22.05  kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_2S16 = &H80             '  22.05  kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_4M08 = &H100            '  44.1   kHz, Mono,   8-bit
Public Const WAVE_FORMAT_4M16 = &H400            '  44.1   kHz, Mono,   16-bit
Public Const WAVE_FORMAT_4S08 = &H200            '  44.1   kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_4S16 = &H800            '  44.1   kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_DIRECT = &H8
Public Const WAVE_FORMAT_DSPGROUP_TRUESPEECH = &H22 ' DSP Group Wave Format
Public Const WAVE_FORMAT_PCM = 1  '  Needed in resource files so outside #ifndef RC_INVOKED
Public Const WAVE_FORMAT_QUERY = &H1
Public Const WAVE_FORMAT_DIRECT_QUERY = (WAVE_FORMAT_QUERY Or WAVE_FORMAT_DIRECT)
Public Const WAVE_INVALIDFORMAT = &H0              '  invalid format
Public Const WAVE_MAPPED = &H4
Public Const WAVE_MAPPER = -1&
Public Const WAVE_VALID = &H3         '  ;Internal
Public Const WAVECAPS_LRVOLUME = &H8         '  separate left-right volume control
Public Const WAVECAPS_PITCH = &H1         '  supports pitch control
Public Const WAVECAPS_PLAYBACKRATE = &H2         '  supports playback rate control
Public Const WAVECAPS_SYNC = &H10
Public Const WAVECAPS_VOLUME = &H4         '  supports volume control
Public Const WAVERR_BASE = 32
Public Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)    '  unsupported wave format
Public Const WAVERR_LASTERROR = (WAVERR_BASE + 3)    '  last error in range
Public Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)    '  still something playing
Public Const WAVERR_SYNC = (WAVERR_BASE + 3)    '  device is synchronous
Public Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)    '  header not prepared
Public Const WHDR_BEGINLOOP = &H4         '  loop start block
Public Const WHDR_DONE = &H1         '  done bit
Public Const WHDR_ENDLOOP = &H8         '  loop end block
Public Const WHDR_INQUEUE = &H10        '  reserved for driver
Public Const WHDR_PREPARED = &H2         '  set if this header has been prepared
Public Const WHDR_VALID = &H1F        '  valid flags      / ;Internal /
Public Type MMTIME
    wType As Long
    u As Long
End Type
Public Type WAVEFORMAT
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
End Type
Public Type WAVEFORMATEX
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    cbSize As Integer
End Type
'in case the the codec information is too big for the regular structure
Public Type WAVEFORMATEXBIG
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    cbSize As Integer
    additional(64) As Byte
End Type
Public Type WAVEHDR
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    Reserved As Long
End Type
Public Type WAVEINCAPS
    wMid As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
    dwFormats As Long
    wChannels As Integer
End Type
Public Type WAVEOUTCAPS
    wMid As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
    dwFormats As Long
    wChannels As Integer
    dwSupport As Long
End Type
Public Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Public Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
Public Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Public Declare Function waveInGetID Lib "winmm.dll" (ByVal hWaveIn As Long, lpuDeviceID As Long) As Long
Public Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function waveInGetPosition Lib "winmm.dll" (ByVal hWaveIn As Long, lpInfo As MMTIME, ByVal uSize As Long) As Long
Public Declare Function waveInMessage Lib "winmm.dll" (ByVal hWaveIn As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Public Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMATEX, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Public Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Public Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Public Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveOutBreakLoop Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEOUTCAPS, ByVal uSize As Long) As Long
Public Declare Function waveOutGetErrorText Lib "winmm.dll" Alias "waveOutGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Public Declare Function waveOutGetID Lib "winmm.dll" (ByVal hWaveOut As Long, lpuDeviceID As Long) As Long
Public Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function waveOutGetPitch Lib "winmm.dll" (ByVal hWaveOut As Long, lpdwPitch As Long) As Long
Public Declare Function waveOutGetPlaybackRate Lib "winmm.dll" (ByVal hWaveOut As Long, lpdwRate As Long) As Long
Public Declare Function waveOutGetPosition Lib "winmm.dll" (ByVal hWaveOut As Long, lpInfo As MMTIME, ByVal uSize As Long) As Long
Public Declare Function waveOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Public Declare Function waveOutMessage Lib "winmm.dll" (ByVal hWaveOut As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Public Declare Function waveOutOpen Lib "winmm.dll" (lphWaveOut As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMATEX, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function waveOutPause Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutRestart Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutSetPitch Lib "winmm.dll" (ByVal hWaveOut As Long, ByVal dwPitch As Long) As Long
Public Declare Function waveOutSetPlaybackRate Lib "winmm.dll" (ByVal hWaveOut As Long, ByVal dwRate As Long) As Long
Public Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Public Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
