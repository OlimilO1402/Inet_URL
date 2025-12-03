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
Private Declare Function InternetOpenW Lib "wininet" (ByVal lpszAgent As longptr, ByVal dwAccessType As Long, ByVal lpszProxy As longptr, ByVal lpszProxyBypass As longptr, ByVal dwFlags As Long) As longptr 'HANDLE

'InternetOpenW:dwAccessType
Private Const INTERNET_OPEN_TYPE_PRECONFIG As Long = 0
Private Const INTERNET_OPEN_TYPE_DIRECT    As Long = 1
Private Const INTERNET_OPEN_TYPE_PROXY     As Long = 3
Private Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY  As Long = 4

'InternetOpenW:dwFlags
Private Const INTERNET_FLAG_NONE       As Long = &H0
Private Const INTERNET_FLAG_ASYNC      As Long = &H10000000
Private Const INTERNET_FLAG_FROM_CACHE As Long = &H1000000
Private Const INTERNET_FLAG_OFFLINE    As Long = INTERNET_FLAG_FROM_CACHE

'https://learn.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetopenurlw
Private Declare Function InternetOpenUrlW Lib "wininet" (ByVal hInternet As longptr, ByVal lpszUrl As longptr, ByVal lpszHeaders As longptr, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As longptr 'HANDLE

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
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As longptr, ByVal lpBuffer As longptr, ByVal dwNumberOfBytesToRead As Long, ByRef out_NumberOfBytesRead As Long) As Long 'BOOL

Private Type NTERNET_BUFFERSW
  dwStructSize As Long
  NextBUFFER   As longptr
  LPCWSTR                   lpcszHeader;
  DWORD                     dwHeadersLength;
  DWORD                     dwHeadersTotal;
  LPVOID                    lpvBuffer;
  DWORD                     dwBufferLength;
  DWORD                     dwBufferTotal;
  DWORD                     dwOffsetLow;
  DWORD                     dwOffsetHigh;
} INTERNET_BUFFERSW, *LPINTERNET_BUFFERSW;

'https://learn.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetreadfileexw
Private Declare Function InternetReadFileExW Lib "wininet" (ByRef hFile As longptr, ByRef lpBuffersOut As INTERNET_BUFFERSW, ByVal dwFlags As Long, ByVal dwContext As Long) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetclosehandle
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInternet As longptr) As Long 'BOOL

Private m_hInet  As longptr
Private m_hFile  As longptr
Private m_sURL   As String
Private m_BufSiz As Long

Private Sub Class_Initialize()
    Dim INET    As String:  INET = "INET"
    Dim PROXY   As String:  PROXY = vbNullString
    Dim PRXYBYP As String: PRXYBYP = vbNullString
    m_hInet = InternetOpenW(StrPtr(INET), INTERNET_OPEN_TYPE_DIRECT, StrPtr(PROXY), StrPtr(PRXYBYP), INTERNET_FLAG_NONE)
End Sub

Friend Sub New_(sURL As String, Optional bufferSize As Long = 16777216) '1048576) '32768) '2048)
    m_sURL = sURL: m_BufSiz = bufferSize
    Dim Headers As String: Headers = vbNullString
    m_hFile = InternetOpenUrlW(m_hInet, StrPtr(m_sURL), StrPtr(Headers), 0, INTERNET_FLAG_RELOAD, 0)
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
    Dim rv As Long: rv = InternetReadFile(m_hFile, StrPtr(Buffer_inout), len_in, pRead)
    Buffer_inout = StrConv(Buffer_inout, vbUnicode)
End Function

Private Sub Class_Terminate()
    If InternetCloseHandle(m_hFile) = 0 Then MsgBox "URL-File-handle could not be closed"
    If InternetCloseHandle(m_hInet) = 0 Then MsgBox "Internet-handle could not be closed"
End Sub
