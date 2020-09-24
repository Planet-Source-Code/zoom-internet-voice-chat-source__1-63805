Attribute VB_Name = "msacm32"
Option Explicit
Private Const MAXEXTRABYTES = 50 '3
'Microsoft Audio Compression Manager APIs
Private Const ACMERR_BASE = 512
Private Const ACMERR_NOTPOSSIBLE = (ACMERR_BASE + 0)
Private Const ACMERR_BUSY = (ACMERR_BASE + 1)
Private Const ACMERR_UNPREPARED = (ACMERR_BASE + 2)
Private Const ACMERR_CANCELED = (ACMERR_BASE + 3)
Private Const ACM_DRIVERENUMF_DISABLED = &H80000000
Private Const ACM_DRIVERENUMF_NOLOCAL = &H40000000
Private Const ACM_FORMATDETAILSF_FORMAT = &H1
Private Const ACM_FORMATDETAILSF_INDEX = &H0
Private Const ACM_FORMATDETAILSF_QUERYMASK = &HF
Private Const ACM_FORMATENUMF_CONVERT = &H100000
Private Const ACM_FORMATENUMF_HARDWARE = &H400000
Private Const ACM_FORMATENUMF_INPUT = &H800000
Private Const ACM_FORMATENUMF_NCHANNELS = &H20000
Private Const ACM_FORMATENUMF_NSAMPLESPERSEC = &H40000
Private Const ACM_FORMATENUMF_OUTPUT = &H1000000
Private Const ACM_FORMATENUMF_SUGGEST = &H200000
Private Const ACM_FORMATENUMF_WBITSPERSAMPLE = &H80000
Private Const ACM_FORMATENUMF_WFORMATTAG = &H10000
Private Const ACM_FORMATSUGGESTF_WFORMATTAG = &H10000
Private Const ACM_FORMATSUGGESTF_NCHANNELS = &H20000
Private Const ACM_FORMATSUGGESTF_NSAMPLESPERSEC = &H40000
Private Const ACM_FORMATSUGGESTF_TYPEMASK = &HFF0000
Private Const ACM_FORMATSUGGESTF_WBITSPERSAMPLE = &H80000
Private Const ACM_FORMATTAGDETAILSF_FORMATTAG = &H1
Private Const ACM_FORMATTAGDETAILSF_INDEX = &H0
Private Const ACM_FORMATTAGDETAILSF_LARGESTSIZE = &H2
Private Const ACM_FORMATTAGDETAILSF_QUERYMASK = &HF
Private Const ACM_METRIC_COUNT_CODECS = &H2
Private Const ACM_METRIC_COUNT_CONVERTERS = &H3
Private Const ACM_METRIC_COUNT_DRIVERS = &H1
Private Const ACM_METRIC_COUNT_FILTERS = &H4
Private Const ACM_METRIC_MAX_SIZE_FILTER = &H33
Private Const ACM_METRIC_MAX_SIZE_FORMAT = &H32
Private Const ACM_STREAMCONVERTF_BLOCKALIGN = &H4&
Private Const ACM_STREAMCONVERTF_START = &H10&
Private Const ACM_STREAMCONVERTF_END = &H20&
Private Const ACM_STREAMOPENF_QUERY = &H1&
Private Const ACM_STREAMOPENF_ASYNC = &H2&
Private Const ACM_STREAMOPENF_NONREALTIME = &H4&
Private Const ACM_STREAMSIZEF_SOURCE = &H0&
Private Const ACM_STREAMSIZEF_DESTINATION = &H1&
Private Const ACM_STREAMSIZEF_QUERYMASK = &HF&
Private Const ACMDRIVERDETAILS_COPYRIGHT_CHARS = &H50
Private Const ACMDRIVERDETAILS_FEATURES_CHARS = &H200
Private Const ACMDRIVERDETAILS_LICENSING_CHARS = &H80
Private Const ACMDRIVERDETAILS_LONGNAME_CHARS = &H80
Private Const ACMDRIVERDETAILS_SHORTNAME_CHARS = &H20
Private Const ACMDRIVERDETAILS_SUPPORTF_ASYNC = &H10
Private Const ACMDRIVERDETAILS_SUPPORTF_CODEC = &H1
Private Const ACMDRIVERDETAILS_SUPPORTF_CONVERTER = &H2
Private Const ACMDRIVERDETAILS_SUPPORTF_DISABLED = &H80000000
Private Const ACMDRIVERDETAILS_SUPPORTF_FILTER = &H4
Private Const ACMDRIVERDETAILS_SUPPORTF_HARDWARE = &H8
Private Const ACMDRIVERDETAILS_SUPPORTF_LOCAL = &H40000000
Private Const ACMFORMATCHOOSE_STYLEF_CONTEXTHELP = &H80
Private Const ACMFORMATCHOOSE_STYLEF_ENABLEHOOK = &H8
Private Const ACMFORMATCHOOSE_STYLEF_ENABLETEMPLATE = &H10
Private Const ACMFORMATCHOOSE_STYLEF_ENABLETEMPLATEHANDLE = &H20
Private Const ACMFORMATCHOOSE_STYLEF_INITTOWFXSTRUCT = &H40
Private Const ACMFORMATCHOOSE_STYLEF_SHOWHELP = &H4
Private Const ACMFORMATDETAILS_FORMAT_CHARS = &H80
Private Const ACMFORMATTAGDETAILS_FORMATTAG_CHARS = &H30
Private Const ACMSTREAMHEADER_STATUSF_DONE = &H10000
Private Const ACMSTREAMHEADER_STATUSF_PREPARED = &H20000
Private Const ACMSTREAMHEADER_STATUSF_INQUEUE = &H100000
'Microsoft Audio Compression Manager APIs
Private Declare Function acmDriverDetails Lib "msacm32.dll" Alias "acmDriverDetailsA" (ByVal hadid As Long, ByRef padd As acmDriverDetailsA, ByVal fdwDetails As Long) As Long
Private Declare Function acmDriverEnum Lib "msacm32.dll" (ByVal fnCallback As Long, ByRef dwInstance As Long, ByVal fdwEnum As Long) As Long
Private Declare Function acmDriverOpen Lib "msacm32.dll" (ByRef phad As Long, ByVal hadid As Long, ByVal fdwOpen As Long) As Long
Private Declare Function acmDriverClose Lib "msacm32.dll" (ByVal had As Long, ByVal fdwClose As Long) As Long
Private Declare Function acmFormatChoose Lib "msacm32.dll" Alias "acmFormatChooseA" (ByRef pafmtc As ACMFORMATCHOOSEA) As Long
Private Declare Function acmFormatDetails Lib "msacm32.dll" Alias "acmFormatDetailsA" (ByVal had As Long, ByRef pafd As acmFormatDetailsA, ByVal fdwDetails As Long) As Long
Private Declare Function acmFormatEnum Lib "msacm32.dll" Alias "acmFormatEnumA" (ByVal had As Long, ByRef pafd As acmFormatDetailsA, ByVal fnCallback As Long, ByRef dwInstance As Long, ByVal fdwEnum As Long) As Long
Private Declare Function acmFormatSuggest Lib "msacm32.dll" (ByVal had As Long, ByRef pwfxSrc As WAVEFORMATEX, ByRef pwfxDst As WAVEFORMATEX, ByVal cbwfxDst As Long, ByVal fdwSuggest As Long) As Long
Private Declare Function acmFormatTagDetails Lib "msacm32.dll" Alias "acmFormatTagDetailsA" (ByVal had As Long, ByRef paftd As ACMFORMATTAGDETAILSA, ByVal fdwDetails As Long) As Long
Private Declare Function acmGetVersion Lib "msacm32.dll" () As Long
Private Declare Function acmMetrics Lib "msacm32.dll" (ByVal hao As Long, ByVal uMetric As Long, ByVal pMetric As Long) As Long
Private Declare Function acmStreamOpen Lib "msacm32" (hAS As Long, ByVal hADrv As Long, ByVal wfxSrc As Long, ByVal wfxDst As Long, ByVal wFltr As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Private Declare Function acmStreamClose Lib "msacm32" (ByVal hAS As Long, ByVal dwClose As Long) As Long
Private Declare Function acmStreamPrepareHeader Lib "msacm32" (ByVal hAS As Long, hASHdr As ACMSTREAMHEADER, ByVal dwPrepare As Long) As Long
Private Declare Function acmStreamUnprepareHeader Lib "msacm32" (ByVal hAS As Long, hASHdr As ACMSTREAMHEADER, ByVal dwUnPrepare As Long) As Long
Private Declare Function acmStreamConvert Lib "msacm32" (ByVal hAS As Long, hASHdr As ACMSTREAMHEADER, ByVal dwConvert As Long) As Long
Private Declare Function acmStreamReset Lib "msacm32" (ByVal hAS As Long, ByVal dwReset As Long) As Long
Private Declare Function acmStreamSize Lib "msacm32" (ByVal hAS As Long, ByVal cbInput As Long, dwOutBytes As Long, ByVal dwSize As Long) As Long
Private Type FIND_DRIVER_INFO
    hadid As Long
    wFormatTag As Long
End Type
Private Type acmDriverDetailsA ' 920 bytes
    cbStruct    As Long
    fccType     As Long
    fccComp     As Long
    wMid        As Integer
    wPid        As Integer
    vdwACM      As Long
    vdwDriver   As Long
    fdwSupport  As Long
    cFormatTags As Long
    cFilterTags As Long
    hIcon       As Long
    szShortName As String * ACMDRIVERDETAILS_SHORTNAME_CHARS
    szLongName As String * ACMDRIVERDETAILS_LONGNAME_CHARS
    szCopyright As String * ACMDRIVERDETAILS_COPYRIGHT_CHARS
    szLicensing As String * ACMDRIVERDETAILS_LICENSING_CHARS
    szFeatures As String * ACMDRIVERDETAILS_FEATURES_CHARS
End Type
Private Type ACMFORMATCHOOSEA
    cbStruct As Long
    fdwStyle As Long
    hWndOwner As Long
    pwfx As Long 'WaveFormatEx
    cbwfx As Long
    pszTitle As Long 'String
    szFormatTag(0 To ACMFORMATTAGDETAILS_FORMATTAG_CHARS - 1) As Byte
    szFormat(0 To ACMFORMATDETAILS_FORMAT_CHARS - 1) As Byte
    pszName As Long ' String
    cchName As Long
    fdwEnum As Long
    pwfxEnum As Long ' WaveFormatEx
    hInstance As Long
    pszTemplateName As Long 'String
    lCustData As Long
    pfnHook As Long 'ACMFORMATCHOOSEHOOKPROC
End Type
Private Type acmFormatDetailsA
    cbStruct As Long
    dwFormatIndex As Long
    dwFormatTag As Long
    fdwSupport As Long
    pwfx As Long            '    LPWAVEFORMATEX pwfx;
    cbwfx As Long
    szFormat As String * ACMFORMATDETAILS_FORMAT_CHARS
End Type
Private Type ACMFORMATTAGDETAILSA
    cbStruct As Long
    dwFormatTagIndex As Long
    dwFormatTag As Long
    cbFormatSize As Long
    fdwSupport As Long
    cStandardFormats As Long
    szFormatTag As String * ACMFORMATTAGDETAILS_FORMATTAG_CHARS
End Type
Private Type ACMSTREAMHEADER            ' ACM STREAM HEADER TYPE]
    cbStruct As Long            ' Size of header in bytes
    dwStatus As Long            ' Conversion status buffer
    dwUser As Long              ' 32 bits of user data specified by application
    pbSrc As Long               ' Source data buffer pointer
    cbSrcLength As Long         ' Source data buffer size in bytes
    cbSrcLengthUsed As Long     ' Source data buffer size used in bytes
    dwSrcUser As Long           ' 32 bits of user data specified by application
    pbDst As Long               ' Dest data buffer pointer
    cbDstLength As Long         ' Dest data buffer size in bytes
    cbDstLengthUsed As Long     ' Dest data buffer size used in bytes
    dwDstUser As Long           ' 32 bits of user data specified by application
    dwReservedDriver(0 To 9) As Long ' Reserved and should not be used
End Type
'Calls window to choose compression format
Public Function ChooseFormat(ByRef infx As WAVEFORMATEX, ByRef outfx As WAVEFORMATEXBIG)
    Dim ACMFC As ACMFORMATCHOOSEA
    Dim nMaxSize As Long
    'Returned value is the size of the largest WAVEFORMATEX structure.
    acmMetrics 0, ACM_METRIC_MAX_SIZE_FORMAT, VarPtr(nMaxSize)
    ACMFC.cbStruct = LenB(ACMFC)
    ACMFC.hWndOwner = frmMain.hWnd
    ACMFC.fdwStyle = ACMFORMATCHOOSE_STYLEF_INITTOWFXSTRUCT
    ACMFC.pwfx = VarPtr(outfx)
    ACMFC.cbwfx = nMaxSize
    ACMFC.fdwEnum = ACM_FORMATENUMF_CONVERT
    ACMFC.pwfxEnum = VarPtr(infx)
    'ACMFC.cbStruct = 232
    'ACMFC.hWndOwner = 1967030
    'ACMFC.fdwStyle = 64
    'ACMFC.pwfx = 2238668
    'ACMFC.fdwEnum = 1048576
    'ACMFC.pwfxEnum = 2238648
    frmMain.addACMStatus ("")
    frmMain.addACMStatus ("Fomrat Choose:")
    frmMain.addACMStatus ("acmfc.cbStruct=" & ACMFC.cbStruct)
    frmMain.addACMStatus ("acmfc.hWndOwner=" & ACMFC.hWndOwner)
    frmMain.addACMStatus ("acmfc.fdwStyle=" & ACMFC.fdwStyle)
    frmMain.addACMStatus ("acmfc.pwfx=" & ACMFC.pwfx)
    frmMain.addACMStatus ("acmfc.fdwEnum=" & ACMFC.fdwEnum)
    frmMain.addACMStatus ("acmfc.pwfxEnum=" & ACMFC.pwfxEnum)
    frmMain.addACMStatus ("Choose format return: " & acmFormatChoose(ACMFC))
End Function
'Converts sound from one format to another
Public Function ConvertFormat(ByRef bWavBufIn() As Byte, ByVal nBufSizeIn As Long, ByRef InFormat As Long, ByRef OutFormat As Long) As Long
    Dim ACMStream As Long, result As Long
    Dim bWavBufOut() As Byte
    Dim nBufSizeOut As Long
    Dim StreamHeader As ACMSTREAMHEADER
    result = acmStreamOpen(ACMStream, ByVal 0&, InFormat, OutFormat, ByVal 0&, ByVal 0&, ByVal 0&, ACM_STREAMOPENF_NONREALTIME)
    result = acmStreamSize(ACMStream, nBufSizeIn, nBufSizeOut, ACM_STREAMSIZEF_SOURCE)
    ReDim bWavBufOut(0 To nBufSizeOut - 1)
    With StreamHeader
        .cbStruct = Len(StreamHeader)
        .pbSrc = VarPtr(bWavBufIn(0))
        .cbSrcLength = nBufSizeIn
        .pbDst = VarPtr(bWavBufOut(0))
        .cbDstLength = nBufSizeOut
    End With
    result = acmStreamPrepareHeader(ACMStream, StreamHeader, 0)
    result = acmStreamConvert(ACMStream, StreamHeader, ACM_STREAMCONVERTF_START Or ACM_STREAMCONVERTF_END)
    result = acmStreamUnprepareHeader(ACMStream, StreamHeader, 0)
    result = acmStreamClose(ACMStream, 0)
    ReDim bWavBufIn(nBufSizeOut)
    bWavBufIn = bWavBufOut
    ConvertFormat = nBufSizeOut
End Function
