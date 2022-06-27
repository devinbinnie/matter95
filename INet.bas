Attribute VB_Name = "INet"
Option Explicit

Public Const INTERNET_AUTODIAL_FORCE_ONLINE As Long = 1
Public Const INTERNET_OPEN_TYPE_PRECONFIG  As Long = 0
Public Const INTERNET_DEFAULT_HTTP_PORT    As Long = 80
Public Const INTERNET_SERVICE_HTTP         As Long = 3
Public Const INTERNET_FLAG_RELOAD          As Long = &H80000000
Public Const HTTP_ADDREQ_FLAG_REPLACE      As Long = &H80000000
Public Const HTTP_ADDREQ_FLAG_ADD          As Long = &H20000000

Public Const HTTP_QUERY_RAW_HEADERS         As Long = 21
Public Const HTTP_QUERY_RAW_HEADERS_CRLF    As Long = 22
Public Const HTTP_QUERY_SET_COOKIE  As Long = 43
Public Const HTTP_QUERY_STATUS_CODE As Long = 19

Public Const C_BUFFER_SIZE                  As Long = 1024
 
Public Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lModifiers As Long) As Long
Public Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As String, ByVal lOptionalLength As Long) As Long
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Long
Public Declare Function HttpEndRequest Lib "wininet.dll" Alias "HttpEndRequestA" (ByVal hRequest As Long, ByVal lpBuffersOut As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hRequest As Long, ByVal dwInfoLevel As Long, ByRef sBuffer As Any, ByRef lpdwBufferLength As Long, ByRef lpdwIndex As Long) As Long

Public Function GetRequest(sURL As String, sResult As String, Optional sToken As String) As Boolean
    Const STR_APP_NAME  As String = "Uploader"
    Dim hOpen           As Long
    Dim hConnection     As Long
    Dim hRequest        As Long
    Dim sHeader         As String
    Dim sBoundary       As String
    Dim sPostData       As String
    Dim sHttpServer     As String
    Dim lHttpPort       As Long
    Dim sUploadPage     As String
    Dim lResultBytes    As Long
    Dim bEnd            As Boolean
    Dim sResultBuffer   As String
 
    '--- parse url
    sHttpServer = sURL
    If InStr(sHttpServer, "://") > 0 Then
        sHttpServer = Mid$(sHttpServer, InStr(sHttpServer, "://") + 3)
    End If
    If InStr(sHttpServer, "/") > 0 Then
        sUploadPage = Mid$(sHttpServer, InStr(sHttpServer, "/"))
        sHttpServer = Left$(sHttpServer, InStr(sHttpServer, "/") - 1)
    End If
    If InStr(sHttpServer, ":") > 0 Then
        On Error Resume Next
        lHttpPort = CLng(Mid$(sHttpServer, InStr(sHttpServer, ":") + 1))
        On Error GoTo 0
        sHttpServer = Left$(sHttpServer, InStr(sHttpServer, ":") - 1)
    End If
    '--- prepare request
    If InternetAutodial(INTERNET_AUTODIAL_FORCE_ONLINE, 0) = 0 Then
        GoTo QH
    End If
    'MsgBox ("did we whatever this is")
    hOpen = InternetOpen(STR_APP_NAME, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If hOpen = 0 Then
        GoTo QH
    End If
    'MsgBox ("Did we internet")
    hConnection = InternetConnect(hOpen, sHttpServer, IIf(lHttpPort <> 0, lHttpPort, INTERNET_DEFAULT_HTTP_PORT), vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
    If hConnection = 0 Then
        GoTo QH
    End If
    'MsgBox ("Did we connect")
    hRequest = HttpOpenRequest(hConnection, "GET", sUploadPage, "HTTP/1.0", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
    If hRequest = 0 Then
        GoTo QH
    End If

    '--- prepare headers
    If Not sToken = "" Then
        Dim sTokenHeader As String
        sTokenHeader = "Authorization: Bearer " & sToken & vbCrLf
        'MsgBox (sTokenHeader)
        If HttpAddRequestHeaders(hRequest, sTokenHeader, Len(sTokenHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD) = 0 Then
            GoTo QH
        End If
    End If

    '--- send request
    If HttpSendRequest(hRequest, vbNullString, 0, vbNullString, 0) = 0 Then
        GoTo QH
    End If
    'MsgBox ("did we request")
    Do
        sResultBuffer = Space(C_BUFFER_SIZE)
        InternetReadFile hRequest, sResultBuffer, C_BUFFER_SIZE, lResultBytes
        'MsgBox (lResultBytes)
        If lResultBytes <= 0 Then
            bEnd = True
        Else
            sResult = sResult & sResultBuffer
        End If
    Loop Until bEnd
    'MsgBox ("did we read")
    '--- success
    GetRequest = True
    'MsgBox ("we did it")
    'MsgBox (sResult)
QH:
    'MsgBox (Err.LastDllError)
    If hRequest <> 0 Then
        Call InternetCloseHandle(hRequest)
    End If
    If hConnection <> 0 Then
        Call InternetCloseHandle(hConnection)
    End If
    If hOpen <> 0 Then
        Call InternetCloseHandle(hOpen)
    End If
End Function

Public Function PostRequest(sURL As String, sPostData As String, sResult As String, sHeaderResult As String, Optional sToken As String) As Boolean
    Const STR_APP_NAME  As String = "Uploader"
    Dim hOpen           As Long
    Dim hConnection     As Long
    Dim hRequest        As Long
    Dim sHeader         As String
    Dim sHeader2        As String
    Dim sBoundary       As String
    Dim nFile           As Integer
    Dim baData()        As Byte
    Dim sHttpServer     As String
    Dim lHttpPort       As Long
    Dim sUploadPage     As String
    Dim lResultBytes    As Long
    Dim bEnd            As Boolean
    Dim sResultBuffer   As String
    Dim sHeaderResultBuffer   As String * &H4D8
    Dim sHeaderResultBuffer2   As String * &H4D8
 
    '--- parse url
    sHttpServer = sURL
    If InStr(sHttpServer, "://") > 0 Then
        sHttpServer = Mid$(sHttpServer, InStr(sHttpServer, "://") + 3)
    End If
    If InStr(sHttpServer, "/") > 0 Then
        sUploadPage = Mid$(sHttpServer, InStr(sHttpServer, "/"))
        sHttpServer = Left$(sHttpServer, InStr(sHttpServer, "/") - 1)
    End If
    If InStr(sHttpServer, ":") > 0 Then
        On Error Resume Next
        lHttpPort = CLng(Mid$(sHttpServer, InStr(sHttpServer, ":") + 1))
        On Error GoTo 0
        sHttpServer = Left$(sHttpServer, InStr(sHttpServer, ":") - 1)
    End If
    
    '--- prepare request
    If InternetAutodial(INTERNET_AUTODIAL_FORCE_ONLINE, 0) = 0 Then
        GoTo QH
    End If
    'MsgBox ("did we whatever this is")
    
    
    hOpen = InternetOpen(STR_APP_NAME, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If hOpen = 0 Then
        GoTo QH
    End If
    'MsgBox ("Did we internet")
    
    
    hConnection = InternetConnect(hOpen, sHttpServer, IIf(lHttpPort <> 0, lHttpPort, INTERNET_DEFAULT_HTTP_PORT), vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
    If hConnection = 0 Then
        GoTo QH
    End If
    'MsgBox ("Did we connect")
    
    
    hRequest = HttpOpenRequest(hConnection, "POST", sUploadPage, "HTTP/1.1", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
    If hRequest = 0 Then
        GoTo QH
    End If
    'MsgBox ("Did we open")
    
    
    '--- prepare headers
    sHeader = "Content-Type: application/json;" & vbCrLf
    If HttpAddRequestHeaders(hRequest, sHeader, Len(sHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD) = 0 Then
        GoTo QH
    End If
    If Not sToken = "" Then
        Dim sTokenHeader As String
        sTokenHeader = "Authorization: Bearer " & sToken & vbCrLf
        'MsgBox (sTokenHeader)
        If HttpAddRequestHeaders(hRequest, sTokenHeader, Len(sTokenHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD) = 0 Then
            GoTo QH
        End If
    End If
    
    '--- send request
    If HttpSendRequest(hRequest, vbNullString, 0, sPostData, Len(sPostData)) = 0 Then
        GoTo QH
    End If
    'MsgBox ("did we request")
    
    '--- pull headers
    If HttpQueryInfo(hRequest, HTTP_QUERY_RAW_HEADERS_CRLF, ByVal sHeaderResultBuffer, Len(sHeaderResultBuffer), 0) = 0 Then
        GoTo QH
    End If
    sHeaderResult = sHeaderResult & sHeaderResultBuffer
    
    '--- get response
    Do
        sResultBuffer = Space(C_BUFFER_SIZE)
        InternetReadFile hRequest, sResultBuffer, C_BUFFER_SIZE, lResultBytes
        'MsgBox (lResultBytes)
        If lResultBytes <= 0 Then
            bEnd = True
        Else
            sResult = sResult & sResultBuffer
        End If
    Loop Until bEnd
    
    'MsgBox ("did we read")
    '--- success
    PostRequest = True
    'MsgBox ("we did it")
    'MsgBox (sResult)
QH:
    'MsgBox (Err.LastDllError)
    If hRequest <> 0 Then
        Call InternetCloseHandle(hRequest)
    End If
    If hConnection <> 0 Then
        Call InternetCloseHandle(hConnection)
    End If
    If hOpen <> 0 Then
        Call InternetCloseHandle(hOpen)
    End If
End Function


