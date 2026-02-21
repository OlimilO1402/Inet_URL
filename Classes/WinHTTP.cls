VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WinHTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum WinHttpRequestOption
    WinHttpRequestOption_UserAgentString = 0
    WinHttpRequestOption_URL = 1
    WinHttpRequestOption_URLCodePage = 2
    WinHttpRequestOption_EscapePercentInURL = 3
    WinHttpRequestOption_SslErrorIgnoreFlags = 4
    WinHttpRequestOption_SelectCertificate = 5
    WinHttpRequestOption_EnableRedirects = 6
    WinHttpRequestOption_UrlEscapeDisable = 7
    WinHttpRequestOption_UrlEscapeDisableQuery = 8
    WinHttpRequestOption_SecureProtocols = 9
    WinHttpRequestOption_EnableTracing = 10
    WinHttpRequestOption_RevertImpersonationOverSsl = 11
    WinHttpRequestOption_EnableHttpsToHttpRedirects = 12
    WinHttpRequestOption_EnablePassportAuthentication = 13
    WinHttpRequestOption_MaxAutomaticRedirects = 14
    WinHttpRequestOption_MaxResponseHeaderSize = 15
    WinHttpRequestOption_MaxResponseDrainSize = 16
    WinHttpRequestOption_EnableHttp1_1 = 17
    WinHttpRequestOption_EnableCertificateRevocationCheck = 18
    WinHttpRequestOption_RejectUserpwd = 19
End Enum

Public Enum WinHttpRequestAutoLogonPolicy
    AutoLogonPolicy_Always = 0
    AutoLogonPolicy_OnlyIfBypassProxy = 1
    AutoLogonPolicy_Never = 2
End Enum

Public Enum WinHttpRequestSslErrorFlags
    SslErrorFlag_UnknownCA = 256
    SslErrorFlag_CertWrongUsage = 512
    SslErrorFlag_CertCNInvalid = 4096
    SslErrorFlag_CertDateInvalid = 8192
    SslErrorFlag_Ignore_All = 13056
End Enum

Public Enum WinHttpRequestSecureProtocols
    SecureProtocol_SSL2 = 8
    SecureProtocol_SSL3 = 32
    SecureProtocol_TLS1 = 128
    SecureProtocol_TLS1_1 = 512
    SecureProtocol_TLS1_2 = 2048
    SecureProtocol_ALL = 168
End Enum

Public Event OnResponseStart(ByVal Status As Long, ByVal ContentType As String)
Public Event OnResponseDataAvailable(Data() As Byte)
Public Event OnResponseFinished()
Public Event OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
' default '


Public Sub SetProxy(ByVal ProxySetting As Long, Optional ByVal ProxyServer As Variant, Optional ByVal BypassList As Variant)
    ' Specify proxy configuration

End Sub

Public Sub SetCredentials(ByVal UserName As String, ByVal Password As String, ByVal Flags As Long)
    ' Specify authentication credentials

End Sub

Public Sub OOpen(ByVal Method As String, ByVal Url As String, Optional ByVal Async As Variant)
    ' Open HTTP connection
    
End Sub


Public Sub SetRequestHeader(ByVal Header As String, ByVal Value As String)
    ' Add HTTP request header

End Sub

Public Function GetResponseHeader(ByVal Header As String) As String
    ' Get HTTP response header
    
End Function

Public Function GetAllResponseHeaders() As String
    ' Get all HTTP response headers
    
End Function

Public Sub Send(Optional ByVal Body As Variant)
    ' Send HTTP request
    
End Sub

Public Property Get Status() As Long
    ' Get HTTP status code
    
End Property

Public Property Get StatusText() As String
    ' Get HTTP status text

End Property

Public Property Get ResponseText() As String
    ' Get response body as a string

End Property

Public Property Get ResponseBody() As Variant
    ' Get response body as a safearray of UI1

End Property


Public Property Get ResponseStream() As Variant
    ' Get response body as a stream

End Property

Public Property Get OOption(ByVal aOption As WinHttpRequestOption) As Variant
    
End Property
Public Property Let OOption(Value As Variant, ByVal aOption As WinHttpRequestOption)
    
End Property

Public Function WaitForResponse(Optional ByVal Timeout As Variant) As Boolean
    ' Wait for asynchronous send to complete, with optional timeout (in seconds)

End Function

Public Sub Abort()
    ' Abort an asynchronous operation in progress

End Sub

Public Sub SetTimeouts(ByVal ResolveTimeout As Long, ByVal ConnectTimeout As Long, ByVal SendTimeout As Long, ByVal ReceiveTimeout As Long)
    ' Specify timeout settings (in milliseconds)

End Sub

Public Sub SetClientCertificate(ByVal ClientCertificate As String)
    ' Specify a client certificate

End Sub

Public Sub SetAutoLogonPolicy(ByVal AutoLogonPolicy As WinHttpRequestAutoLogonPolicy)
    ' Specify if credentials should be sent automatically

End Sub

