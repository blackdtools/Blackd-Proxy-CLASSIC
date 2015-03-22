Attribute VB_Name = "modFTP"
Option Explicit

Public Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Declare Function GetTickCount& Lib "kernel32" ()
Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long

Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal pub_lngInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Sub InternetSetStatusCallback Lib "wininet.dll" (ByVal pub_lngInternetSession As Long, ByVal lpfnInternetCallback As Long)
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal pub_lngInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, sOptional As Any, ByVal lOptionalLength As Long) As Integer
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetWriteFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToWrite As Long, dwNumberOfBytesWritten As Long) As Integer
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Public Declare Function InternetQueryDataAvailable Lib "wininet.dll" (ByVal hInet As Long, dwAvail As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Public Declare Function InternetTimeToSystemTime Lib "wininet.dll" (ByVal lpszTime As String, ByRef pst As SYSTEMTIME, ByVal dwReserved As Long) As Long
         
Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Long
Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String, ByVal fdwAccess As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Long, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Long
Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Long
Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Long
Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Long
Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Long
                  
' Use registry access settings.
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1

' Number of the TCP/IP port on the server to connect to.
Public Const INTERNET_INVALID_PORT_NUMBER = 0
Public Const INTERNET_DEFAULT_FTP_PORT = 21
Public Const INTERNET_DEFAULT_GOPHER_PORT = 70
Public Const INTERNET_DEFAULT_HTTP_PORT = 80
Public Const INTERNET_DEFAULT_HTTPS_PORT = 443
Public Const INTERNET_DEFAULT_SOCKS_PORT = 1080

' Type of service to access.
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3

' Brings the data across the wire even if it locally cached.
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const ERROR_NO_MORE_FILES = 18

Public Const FTP_TRANSFER_TYPE_UNKNOWN As Long = &H0 '0x00000000
Public Const FTP_TRANSFER_TYPE_ASCII As Long = &H1 '0x00000001
Public Const FTP_TRANSFER_TYPE_BINARY  As Long = &H2 '0x00000002

' The possible values for the lInfoLevel parameter include:
Public Const HTTP_QUERY_CONTENT_TYPE = 1
Public Const HTTP_QUERY_CONTENT_LENGTH = 5
Public Const HTTP_QUERY_EXPIRES = 10
Public Const HTTP_QUERY_LAST_MODIFIED = 11
Public Const HTTP_QUERY_PRAGMA = 17
Public Const HTTP_QUERY_VERSION = 18
Public Const HTTP_QUERY_STATUS_CODE = 19
Public Const HTTP_QUERY_STATUS_TEXT = 20
Public Const HTTP_QUERY_RAW_HEADERS = 21
Public Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Public Const HTTP_QUERY_FORWARDED = 30
Public Const HTTP_QUERY_SERVER = 37
Public Const HTTP_QUERY_USER_AGENT = 39
Public Const HTTP_QUERY_SET_COOKIE = 43
Public Const HTTP_QUERY_REQUEST_METHOD = 45

' Add this flag to the about flags to get request header.
Public Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000

' flags for InternetOpenUrl
Public Const INTERNET_FLAG_RAW_DATA = &H40000000
Public Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000
Public Const INTERNET_FLAG_TRANSFER_ASCII = &H1&
Public Const INTERNET_FLAG_TRANSFER_BINARY = &H2&

' flags for InternetOpen
Public Const INTERNET_FLAG_ASYNC = &H10000000
Public Const INTERNET_FLAG_PASSIVE = &H8000000
Public Const INTERNET_FLAG_DONT_CACHE = &H4000000
Public Const INTERNET_FLAG_MAKE_PERSISTENT = &H2000000
Public Const INTERNET_FLAG_OFFLINE = &H1000000

Public Type INTERNET_ASYNC_RESULT
    dwResult As Long
    dwError As Long
End Type

Public Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000     ' don't write this item to the cache
Public Const INTERNET_STATUS_RESOLVING_NAME = 10
Public Const INTERNET_STATUS_NAME_RESOLVED = 11
Public Const INTERNET_STATUS_CONNECTING_TO_SERVER = 20
Public Const INTERNET_STATUS_CONNECTED_TO_SERVER = 21
Public Const INTERNET_STATUS_SENDING_REQUEST = 30
Public Const INTERNET_STATUS_REQUEST_SENT = 31
Public Const INTERNET_STATUS_RECEIVING_RESPONSE = 40
Public Const INTERNET_STATUS_RESPONSE_RECEIVED = 41
Public Const INTERNET_STATUS_CTL_RESPONSE_RECEIVED = 42
Public Const INTERNET_STATUS_PREFETCH = 43
Public Const INTERNET_STATUS_CLOSING_CONNECTION = 50
Public Const INTERNET_STATUS_CONNECTION_CLOSED = 51
Public Const INTERNET_STATUS_HANDLE_CREATED = 60
Public Const INTERNET_STATUS_HANDLE_CLOSING = 70
Public Const INTERNET_STATUS_REQUEST_COMPLETE = 100
Public Const INTERNET_STATUS_REDIRECT = 110
Public Const INTERNET_STATUS_STATE_CHANGE = 200
Public Const INTERNET_ERROR_BASE = 12000

Public Const ERROR_INTERNET_OUT_OF_HANDLES = 12001
Public Const ERROR_INTERNET_TIMEOUT = 12002
Public Const ERROR_INTERNET_EXTENDED_ERROR = 12003
Public Const ERROR_INTERNET_INTERNAL_ERROR = 12004
Public Const ERROR_INTERNET_INVALID_URL = 12005
Public Const ERROR_INTERNET_UNRECOGNIZED_SCHEME = 12006
Public Const ERROR_INTERNET_NAME_NOT_RESOLVED = 12007
Public Const ERROR_INTERNET_PROTOCOL_NOT_FOUND = 12008
Public Const ERROR_INTERNET_INVALID_OPTION = 12009
Public Const ERROR_INTERNET_BAD_OPTION_LENGTH = 12010
Public Const ERROR_INTERNET_OPTION_NOT_SETTABLE = 12011
Public Const ERROR_INTERNET_SHUTDOWN = 12012
Public Const ERROR_INTERNET_INCORRECT_USER_NAME = 12013
Public Const ERROR_INTERNET_INCORRECT_PASSWORD = 12014
Public Const ERROR_INTERNET_LOGIN_FAILURE = 12015
Public Const ERROR_INTERNET_INVALID_OPERATION = 12016
Public Const ERROR_INTERNET_OPERATION_CANCELLED = 12017
Public Const ERROR_INTERNET_INCORRECT_HANDLE_TYPE = 12018
Public Const ERROR_INTERNET_INCORRECT_HANDLE_STATE = 12019
Public Const ERROR_INTERNET_NOT_PROXY_REQUEST = 12020
Public Const ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND = 12021
Public Const ERROR_INTERNET_BAD_REGISTRY_PARAMETER = 12022
Public Const ERROR_INTERNET_NO_DIRECT_ACCESS = 12023
Public Const ERROR_INTERNET_NO_CONTEXT = 12024
Public Const ERROR_INTERNET_NO_CALLBACK = 12025
Public Const ERROR_INTERNET_REQUEST_PENDING = 12026
Public Const ERROR_INTERNET_INCORRECT_FORMAT = 12027
Public Const ERROR_INTERNET_ITEM_NOT_FOUND = 12028
Public Const ERROR_INTERNET_CANNOT_CONNECT = 12029
Public Const ERROR_INTERNET_CONNECTION_ABORTED = 12030
Public Const ERROR_INTERNET_CONNECTION_RESET = 12031
Public Const ERROR_INTERNET_FORCE_RETRY = 12032
Public Const ERROR_INTERNET_INVALID_PROXY_REQUEST = 12033
Public Const ERROR_INTERNET_NEED_UI = 12034

Public Const ERROR_INTERNET_HANDLE_EXISTS = 12036
Public Const ERROR_INTERNET_SEC_CERT_DATE_INVALID = 12037
Public Const ERROR_INTERNET_SEC_CERT_CN_INVALID = 12038
Public Const ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR = 12039
Public Const ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR = 12040
Public Const ERROR_INTERNET_MIXED_SECURITY = 12041
Public Const ERROR_INTERNET_CHG_POST_IS_NON_SECURE = 12042
Public Const ERROR_INTERNET_POST_IS_NON_SECURE = 12043
Public Const ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED = 12044
Public Const ERROR_INTERNET_INVALID_CA = 12045
Public Const ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP = 12046
Public Const ERROR_INTERNET_ASYNC_THREAD_FAILED = 12047
Public Const ERROR_INTERNET_REDIRECT_SCHEME_CHANGE = 12048

'//
'// FTP API errors
'//

Public Const ERROR_FTP_TRANSFER_IN_PROGRESS = 12110
Public Const ERROR_FTP_DROPPED = 12111

'//
'// gopher API errors
'//

Public Const ERROR_GOPHER_PROTOCOL_ERROR = 12130
Public Const ERROR_GOPHER_NOT_FILE = 12131
Public Const ERROR_GOPHER_DATA_ERROR = 12132
Public Const ERROR_GOPHER_END_OF_DATA = 12133
Public Const ERROR_GOPHER_INVALID_LOCATOR = 12134
Public Const ERROR_GOPHER_INCORRECT_LOCATOR_TYPE = 12135
Public Const ERROR_GOPHER_NOT_GOPHER_PLUS = 12136
Public Const ERROR_GOPHER_ATTRIBUTE_NOT_FOUND = 12137
Public Const ERROR_GOPHER_UNKNOWN_LOCATOR = 12138

'//
'// HTTP API errors
'//

Public Const ERROR_HTTP_HEADER_NOT_FOUND = 12150
Public Const ERROR_HTTP_DOWNLEVEL_SERVER = 12151
Public Const ERROR_HTTP_INVALID_SERVER_RESPONSE = 12152
Public Const ERROR_HTTP_INVALID_HEADER = 12153
Public Const ERROR_HTTP_INVALID_QUERY_REQUEST = 12154
Public Const ERROR_HTTP_HEADER_ALREADY_EXISTS = 12155
Public Const ERROR_HTTP_REDIRECT_FAILED = 12156
Public Const ERROR_HTTP_NOT_REDIRECTED = 12160               '// BUGBUG

Public Const ERROR_INTERNET_SECURITY_CHANNEL_ERROR = 12157   '// BUGBUG
Public Const ERROR_INTERNET_UNABLE_TO_CACHE_FILE = 12158    ' // BUGBUG
Public Const ERROR_INTERNET_TCPIP_NOT_INSTALLED = 12159      '// BUGBUG

Public Const INTERNET_ERROR_LAST = 12159

Public Const MAX_PATH = 260
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1
Public Const FILE_END = 2

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const OPEN_ALWAYS = 4

Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type
Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

Declare Function GetLastError Lib "kernel32" () As Long
Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpfilename As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpfilename As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FileTimeToSystemTime& Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME)
Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long
Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long


Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
                            (ByVal dwFlags As Long, lpSource As Any, _
                            ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
                            ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Public Const LANG_USER_DEFAULT = &H400&

Const SWP_DRAWFRAME = &H20
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Public Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const RAS_MAXENTRYNAME As Integer = 256
Public Const RAS_MAXDEVICETYPE As Integer = 16
Public Const RAS_MAXDEVICENAME As Integer = 128
Public Const RAS_RASCONNSIZE As Integer = 412
 Type RASCONN
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
    szDeviceType(RAS_MAXDEVICETYPE) As Byte
    szDeviceName(RAS_MAXDEVICENAME) As Byte
End Type

Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (udtRasConn As Any, lpcb As Long, lpcConnections As Long) As Long
Declare Function RasHangUp Lib "rasapi32.dll" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long

Const sReadBuffer = 1024
Dim lSesion As Long
Dim lServer As Long
Global bSemBajada As Boolean
Public finbajada As Long

Public Function filtroBD(strRuta As String) As String
    Dim lngPos As Long
    Dim strRes As String
    lngPos = InStr(1, strRuta, "/")
    If lngPos = 0 Then
      strRes = strRuta
    Else
      strRes = Right$(strRuta, Len(strRuta) - lngPos)
    End If
    filtroBD = strRes
End Function
Public Function conectar(sName As String, sServer As String, sUser As String, sPass As String, Optional sDir As String = "") As Boolean
    On Error GoTo goterr
    Dim udtWFD As WIN32_FIND_DATA
    lSesion = InternetOpen(sName, INTERNET_OPEN_TYPE_DIRECT, "", "", INTERNET_FLAG_NO_CACHE_WRITE)
    If lSesion <> 0 Then
        lServer = InternetConnect(lSesion, sServer, 21, sUser, sPass, INTERNET_SERVICE_FTP, INTERNET_FLAG_EXISTING_CONNECT, &H0)
        If lServer <> 0 Then
            If sDir <> "" Then
                FtpSetCurrentDirectory lSesion, "/" & sDir & "/*.*"
            End If
            conectar = True
            Exit Function
        End If
    End If
    frmMainDownload.AddLog "Can't connect to FTP. (Can't open session)"
    conectar = False
    Exit Function
goterr:
    frmMainDownload.AddLog "Can't connect to FTP. Error number " & CStr(Err.Number) & " ;  Error description: " & Err.Description
    conectar = False
End Function

Public Function desconectar() As Boolean
    On Error GoTo goterr
    InternetCloseHandle lServer
    InternetCloseHandle lSesion
    lServer = 0
    lSesion = 0
    desconectar = True
    Exit Function
goterr:
    desconectar = False
End Function

Public Function estaConectado() As Boolean

    estaConectado = (lSesion <> 0 And lServer <> 0)

End Function

Public Function bajar(ByVal sDir As String, ByVal sPat As String, oFrm As Form) As String
    On Error GoTo huboerror
    Dim strFile As String
    Dim hFile As Long
    Dim udtWFD As WIN32_FIND_DATA
    Dim bBajar As Boolean
    Dim bAlguno As Boolean
    Dim sPat2 As String
    bAlguno = False

    'Búsqueda de ficheros
    sDir = Replace(sDir, "\", "/")
    sPat = Replace(sPat, "\", "/")
    sDir = Replace(sDir, " ", "?")
    sPat2 = Replace(sPat, " ", "?")
    hFile = FtpFindFirstFile(lServer, sDir & sPat2, udtWFD, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE, 0&)
    If hFile Then

        DoEvents
        DoEvents
        strFile = Left(udtWFD.cFileName, InStr(1, udtWFD.cFileName, Chr(0)) - 1)
            If Len(strFile) > 0 Then
                If udtWFD.dwFileAttributes And vbDirectory Then
                    'Directorio
                Else
                    bAlguno = True
                    'Fichero
                    bBajar = Download(sDir, sPat, udtWFD.nFileSizeLow, oFrm)
                    'If bBajar Then
                        'Debug.Print ""
                    'End If
                End If
            End If
        If Not bAlguno Then
            'txtInfo = txtInfo & "- No existen datos para recibir." & vbCrLf
        End If
    Else
        If finbajada <> 2 Then
            finbajada = 3
        End If
        bajar = "Sorry, unable to download ( error #" & Err.LastDllError & " )"
        Exit Function
    End If
    bajar = ""
    Exit Function
huboerror:
    bajar = "Error " & CStr(Err.Number) & " : " & Err.Description
End Function

Public Function Download(sDir As String, sFile As String, lTam As Long, oFrm As Form) As Boolean
    Dim sBuffer As String
    Dim FileData As String
    Dim Ret As Long
    Dim hOrig As Long
    Dim hDest As Integer
    Dim sOrig As String
    Dim sDest As String
    Dim nLong As Long
    Dim strRealDest As String
    strRealDest = App.Path
    If Right$(strRealDest, 1) = "\" Or Right$(strRealDest, 1) = "/" Then
        strRealDest = Left$(strRealDest, Len(strRealDest) - 1)
    End If
    sOrig = sDir & sFile
    sDest = strRealDest & "\" & sFile
    sDest = Replace(sDest, "/", "\")
    sBuffer = Space(sReadBuffer)
    FileData = ""
    hOrig = FtpOpenFile(lServer, sOrig, GENERIC_READ, FTP_TRANSFER_TYPE_BINARY, 0)
    If hOrig = 0 Then
        Download = False
        If finbajada <> 2 Then
            finbajada = 3
        End If
    Else
        hDest = FreeFile
        Open sDest For Binary As #hDest
        Do
            DoEvents
            DoEvents
            InternetReadFile hOrig, sBuffer, sReadBuffer, Ret
            If Ret <> sReadBuffer Then
                sBuffer = Left$(sBuffer, Ret)
            End If
            nLong = nLong + Len(sBuffer)
            'FileData = FileData + sBuffer
            Put #hDest, , sBuffer 'FileData
            oFrm.pintar nLong / lTam * 100
        Loop Until Ret <> sReadBuffer Or (finbajada > 0)
        Close #hDest
        InternetCloseHandle hOrig
        DoEvents
        'borrar sFile
        If nLong = lTam Then
            If finbajada <> 2 Then
            finbajada = 1
            End If
        End If
        Download = True
    End If
End Function



