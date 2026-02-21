VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InternetURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'inspired by this thread at ActiveVB:
'http://foren.activevb.de/forum/vb-classic/thread-416377/beitrag-416380/Re-feststellen-ob-mein-Rechner/

'Hinweis:
'WinINet unterstützt keine Serverimplementierungen. Darüber hinaus sollte es nicht von einem Dienst verwendet werden.
'Verwenden Sie für Serverimplementierungen oder Dienste Microsoft Windows HTTP Services (WinHTTP).
 
'https://learn.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetopenw
Private Declare Function InternetOpenW Lib "wininet" (ByVal lpszAgent As LongPtr, ByVal dwAccessType As Long, ByVal lpszProxy As LongPtr, ByVal lpszProxyBypass As LongPtr, ByVal dwFlags As Long) As LongPtr 'HANDLE

'InternetOpenW:dwAccessType
Private Const INTERNET_OPEN_TYPE_PRECONFIG As Long = 0
Private Const INTERNET_OPEN_TYPE_DIRECT    As Long = 1
Private Const INTERNET_OPEN_TYPE_PROXY     As Long = 3
Private Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY  As Long = 4

'InternetOpenW:dwFlags
Private Const INTERNET_FLAG_NONE       As Long = &H0&
Private Const INTERNET_FLAG_ASYNC      As Long = &H10000000
Private Const INTERNET_FLAG_FROM_CACHE As Long = &H1000000
Private Const INTERNET_FLAG_OFFLINE    As Long = INTERNET_FLAG_FROM_CACHE

'https://learn.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetopenurlw
Private Declare Function InternetOpenUrlW Lib "wininet" (ByVal hInternet As LongPtr, ByVal lpszUrl As LongPtr, ByVal lpszHeaders As LongPtr, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As LongPtr 'HANDLE

'InternetOpenUrlW:dwFlags
Private Const INTERNET_FLAG_NEED_FILE                As Long = &H10&
Private Const INTERNET_FLAG_PRAGMA_NOCACHE           As Long = &H100&
Private Const INTERNET_FLAG_NO_UI                    As Long = &H200&
Private Const INTERNET_FLAG_HYPERLINK                As Long = &H400&
Private Const INTERNET_FLAG_RESYNCHRONIZE            As Long = &H800&
Private Const INTERNET_FLAG_IGNORE_CERT_CN_INVALID   As Long = &H1000&
Private Const INTERNET_FLAG_IGNORE_CERT_DATE_INVALID As Long = &H2000&
Private Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS As Long = &H4000&
Private Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP  As Long = &H8000&
Private Const INTERNET_FLAG_NO_AUTH                  As Long = &H40000
Private Const INTERNET_FLAG_NO_COOKIES               As Long = &H80000
Private Const INTERNET_FLAG_NO_AUTO_REDIRECT         As Long = &H200000
Private Const INTERNET_FLAG_KEEP_CONNECTION          As Long = &H400000
Private Const INTERNET_FLAG_SECURE                   As Long = &H800000
Private Const INTERNET_FLAG_NO_CACHE_WRITE           As Long = &H4000000
Private Const INTERNET_FLAG_PASSIVE                  As Long = &H8000000
Private Const INTERNET_FLAG_EXISTING_CONNECT         As Long = &H20000000
Private Const INTERNET_FLAG_RAW_DATA                 As Long = &H40000000
Private Const INTERNET_FLAG_RELOAD                   As Long = &H80000000

'https://learn.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetreadfile
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal dwNumberOfBytesToRead As Long, ByRef out_NumberOfBytesRead As Long) As Long 'BOOL

Private Type INTERNET_BUFFERSW ' INTERNET_BUFFERSW, *LPINTERNET_BUFFERSW
    dwStructSize    As Long 'DWORD
    NextBUFFER      As LongPtr 'struct _INTERNET_BUFFERSW *
    lpcszHeader     As LongPtr 'LPCWSTR
    dwHeadersLength As Long    'DWORD
    dwHeadersTotal  As Long    'DWORD
    lpvBuffer       As LongPtr 'LPVOID
    dwBufferLength  As Long    'DWORD
    dwBufferTotal   As Long    'DWORD
    dwOffsetLow     As Long    'DWORD
    dwOffsetHigh    As Long    'DWORD
End Type

'https://learn.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetreadfileexw
Private Declare Function InternetReadFileExW Lib "wininet" (ByRef hFile As LongPtr, ByRef lpBuffersOut As INTERNET_BUFFERSW, ByVal dwFlags As Long, ByVal dwContext As Long) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetclosehandle
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInternet As LongPtr) As Long 'BOOL

Private m_hInet  As LongPtr
Private m_hURL   As LongPtr
Private m_sURL   As String
Private m_BufSiz As Long

Private Sub Class_Initialize()
    Dim INET    As String:    INET = "INET"
    Dim PROXY   As String:   PROXY = vbNullString
    Dim PRXYBYP As String: PRXYBYP = vbNullString
    m_hInet = InternetOpenW(StrPtr(INET), INTERNET_OPEN_TYPE_DIRECT, StrPtr(PROXY), StrPtr(PRXYBYP), INTERNET_FLAG_NONE)
End Sub

Friend Sub New_(sURL As String, Optional bufferSize As Long = 16777216) '1048576) '32768) '2048)
    m_sURL = sURL: m_BufSiz = bufferSize
    Dim Headers As String: Headers = vbNullString
    m_hURL = InternetOpenUrlW(m_hInet, StrPtr(m_sURL), StrPtr(Headers), 0&, INTERNET_FLAG_DONT_CACHE Or INTERNET_FLAG_RELOAD, 0&)
End Sub
'Copy to MNew:
'Public Function InternetURL(ByVal sURL As String, Optional bufferSize As Long = 2048) As InternetURL
'    Set InternetURL = New InternetURL: InternetURL.New_ sURL, bufferSize
'End Function

Public Function Read() As String
    Dim s_ges As String, s As String
    Dim nBytesRead As Long: nBytesRead = m_BufSiz
    Dim bufferSize As Long: bufferSize = nBytesRead
    Dim nBytes_ges As Long
    Do
        nBytesRead = pRead(bufferSize, s)
        s_ges = s_ges & s
        nBytes_ges = nBytes_ges + nBytesRead
        If nBytesRead < bufferSize Then Exit Do
        bufferSize = bufferSize * 2
    Loop
    Read = Left(s_ges, nBytes_ges)
    Debug.Print nBytes_ges
End Function

Private Function pRead(ByVal len_in As Long, Buffer_inout As String) As Long
    Buffer_inout = Space(len_in)
    Dim rv As Long: rv = InternetReadFile(m_hURL, StrPtr(Buffer_inout), len_in, pRead)
    Buffer_inout = StrConv(Buffer_inout, vbUnicode)
End Function

Private Sub Class_Terminate()
    If InternetCloseHandle(m_hFile) = 0 Then MsgBox "URL-File-handle could not be closed"
    If InternetCloseHandle(m_hInet) = 0 Then MsgBox "Internet-handle could not be closed"
End Sub


'Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
'Private Declare Function CreateFileW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, Optional ByVal dwFlagsAndAttributes As Long, Optional ByVal hTemplateFile As Long) As Long
'Private Declare Function GetQueueStatus Lib "user32.dll" (ByVal Flags As Long) As Long
'Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInternet As Long) As Long
'Private Declare Function InternetOpenW Lib "wininet.dll" (ByVal lpszAgent As Long, ByVal dwAccessType As Long, ByVal lpszProxyName As Long, ByVal lpszProxyBypass As Long, ByVal dwFlags As Long) As Long
'Private Declare Function InternetOpenUrlW Lib "wininet.dll" (ByVal hInternet As Long, ByVal lpszUrl As Long, ByVal lpszHeaders As Long, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
'Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal dwNumberOfBytesToRead As Long, ByRef lpdwNumberOfBytesRead As Long) As Long
'Private Declare Function SysReAllocStringLen Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long, Optional ByVal Length As Long) As Long
'Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, Optional ByRef lpNumberOfBytesWritten As Long, Optional ByVal lpOverlapped As Long) As Long
'
''Downloads the remote file specified by the sURL argument to the local file pointed
''by the sFileName parameter. The optional Chunk parameter determines the number
''of bytes to be downloaded at a time. Bigger chunks download faster while smaller
''ones enables the GUI to be more responsive. Returns the total number of bytes
''successfully written to disk. Maximum download size of 2,047.99 MB only.
'
'Public Function DownloadURL2File(ByRef sURL As String, ByRef sFileName As String, Optional ByVal Chunk As Long = 1024&) As Long
'    Const INTERNET_OPEN_TYPE_DIRECT = 1&, INTERNET_FLAG_DONT_CACHE = &H4000000, INTERNET_FLAG_RELOAD = &H80000000
'    Const GENERIC_WRITE = &H40000000, FILE_SHARE_NONE = 0&, CREATE_ALWAYS = 2&, QS_ALLINPUT = &H4FF&
'    Const INVALID_HANDLE_VALUE = -1&, ERROR_INSUFFICIENT_BUFFER = &H7A&
'    Dim hInternet As Long, hURL As Long, hFile As Long, nBytesRead As Long, nBytesWritten As Long
'    Dim bSuccess As Boolean, sBuffer_Ptr As Long, sBuffer_Size As Long, sBuffer As String
'
'    Select Case True
'        Case LenB(sURL) = 0&, LenB(sFileName) = 0&, Chunk < 2&:  Exit Function
'    End Select
'
'    hInternet = InternetOpenW(StrPtr(App.Title), INTERNET_OPEN_TYPE_DIRECT, 0&, 0&, 0&)
'    If hInternet Then
'        hURL = InternetOpenUrlW(hInternet, StrPtr(sURL), 0&, 0&, INTERNET_FLAG_DONT_CACHE Or INTERNET_FLAG_RELOAD, 0&)
'        If hURL Then
'            hFile = CreateFileW(StrPtr(sFileName), GENERIC_WRITE, FILE_SHARE_NONE, 0&, CREATE_ALWAYS) 'Overwrite existing
'            If hFile <> INVALID_HANDLE_VALUE Then
'                Do: SysReAllocStringLen VarPtr(sBuffer), , (sBuffer_Size + Chunk) * 0.5!
'                    sBuffer_Size = LenB(sBuffer):   sBuffer_Ptr = StrPtr(sBuffer)
'                    Do While InternetReadFile(hURL, sBuffer_Ptr, sBuffer_Size, nBytesRead)
'                        If nBytesRead Then
'                            bSuccess = (WriteFile(hFile, sBuffer_Ptr, nBytesRead, nBytesWritten) <> 0&) And (nBytesWritten = nBytesRead): Debug.Assert bSuccess
'                            If bSuccess Then DownloadURL2File = DownloadURL2File + nBytesWritten
'                            If GetQueueStatus(QS_ALLINPUT) And &HFFFF0000 Then DoEvents
'                        Else
'                            Exit Do
'                        End If
'                    Loop
'                Loop While Err.LastDllError = ERROR_INSUFFICIENT_BUFFER
'                hFile = CloseHandle(hFile):                               Debug.Assert hFile
'            End If
'            hURL = InternetCloseHandle(hURL):                             Debug.Assert hURL
'        End If
'        hInternet = InternetCloseHandle(hInternet):                       Debug.Assert hInternet
'    End If
'End Function
'
'
'Private Declare Function InternetCheckConnectionW Lib "wininet.dll" (Optional ByVal lpszUrl As Long, Optional ByVal dwFlags As Long, Optional ByVal dwReserved As Long) As Long
'
''Allows an application to check if a connection to the Internet can be established.
'Public Function IsInternetConnected(Optional ByRef sURL As String = "http://www.google.com/") As Boolean
'    Const FLAG_ICC_FORCE_CONNECTION = &H1&
'
'    IsInternetConnected = InternetCheckConnectionW(StrPtr(sURL), FLAG_ICC_FORCE_CONNECTION)
'End Function
