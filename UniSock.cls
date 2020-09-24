VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UniSock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************************
'* UniSock - Winsock replacement
'* -----------------------------
'* By Vesa Piittinen aka Merri, http://vesa.piittinen.name/ <vesa@piittinen.name>
'*
'* LICENSE
'* -------
'* http://creativecommons.org/licenses/by-sa/1.0/fi/deed.en
'*
'* Terms: 1) If you make your own version, share using this same license.
'*        2) When used in a program, mention my name in the program's credits.
'*        3) Free for commercial and non-commercial usage.
'*        4) Use at your own risk. No support guaranteed.
'*
'* REQUIREMENTS
'* ------------
'* Huh what? Just this one class module. No extra files required.
'*
'* VERSION HISTORY
'* ---------------
'* Version 1.2 (2008-08-12)
'* - Vista connection error resolved!
'*
'* Version 1.1 (2008-08-07)
'* - CreateWindowExW fixed: had forgotten to call DefWindowProc, no more subclassing!
'* - Fixed crash bug when using an invalid RequestID value for Accept.
'* - No more blank error messages, all errors should be covered.
'* - Unicode support: you may now send and receive Unicode! You can get into this mode by using Mode = [Socket Text Mode]
'*   - event: TextArrival
'*   - property: Mode (default: [Socket Binary Mode])
'*   - method: SendText (default: UTF-8)
'*   How it works? Well, first of all: you do not need to care about line changes. SendText adds line change to each line.
'*   TextArrival also processes data one line at a time. Unicode detection is done by looking for line changes.
'*   For example, [CR][00][LF][00] guarantees the data is UTF-16. [CR][LF] guarantees data is either UTF-8 or ANSI.
'*   If data is UTF-16 or UTF-8 you get it in "ready to use" format. ANSI however is provided as raw data (use StrConv).
'*
'* Version 1.0 (2008-08-06)
'* - All the Winsock control features implemented.
'* - Changes & extra features:
'*   - Autocomplete: datatypes have changed in some events and methods to be more precise.
'*   - GetData and SendData support a wide range of array datatypes. (Including string arrays!)
'*   - LineChange: line change used when sending string arrays.
'*   - RequestHost, RequestIP and RequestPort must be used instead of Remote ones when using Accept event.
'*   - Some methods return True on success.
'*
'* Important: UDP SUPPORT IS UNTESTED. Bug reports with sample projects welcome!
'*
'* CREDITS
'* -------
'* Paul Caton and LaVolpe for their work on SelfSub, SelfHook and SelfCallback
'*************************************************************************************************
Option Explicit

Public Event Closing()
Public Event Connect()
Public Event ConnectionRequest(ByVal RequestID As Long)
Public Event DataArrival(ByVal BytesTotal As Long)
Public Event Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Public Event SendComplete()
Public Event SendProgress(ByVal BytesSent As Long, ByVal BytesRemaining As Long)
Public Event StatusChange(ByVal Status As StateConstants)
Public Event TextArrival(Text As String, ByVal LineChange As UniSockLineChange, ByVal ANSI As Boolean)

' for compatibility...
Public Enum ErrorConstants
    sckAddressInUse = 10048
    sckAddressNotAvailable = 10049
    sckAlreadyComplete = 10037
    sckAlreadyConnected = 10056
    sckBadState = 40006
    sckConnectAborted = 10053
    sckConnectionRefused = 10061
    sckConnectionReset = 10054
    sckGetNotSupported = 394
    sckHostNotFound = 11001
    sckHostNotFoundTryAgain = 11002
    sckInProgress = 10036
    sckInvalidArg = 40014
    sckInvalidArgument = 10014
    sckInvalidOp = 40020
    sckInvalidPropertyValue = 380
    sckMsgTooBig = 10040
    sckNetReset = 10052
    sckNetworkSubsystemFailed = 10050
    sckNetworkUnreachable = 10051
    sckNoBufferSpace = 10055
    sckNoData = 11004
    sckNonRecoverableError = 11003
    sckNotConnected = 10057
    sckNotInitialized = 10093
    sckNotSocket = 10038
    sckOpCanceled = 10004
    sckOutOfMemory = 7
    sckOutOfRange = 40021
    sckPortNotSupported = 10043
    sckSetNotSupported = 383
    sckSocketShutdown = 10058
    sckSuccess = 40017
    sckTimedout = 10060
    sckUnsupported = 40018
    sckWouldBlock = 10035
    sckWrongProtocol = 40026
End Enum

' for compatibility...
Public Enum ProtocolConstants
    sckTCPProtocol
    sckUDPProtocol
End Enum

' for compatibility...
Public Enum StateConstants
    sckClosed
    sckOpen
    sckListening
    sckConnectionPending
    sckResolvingHost
    sckHostResolved
    sckConnecting
    sckConnected
    sckClosing
    sckError
End Enum

Public Enum UniSockLineChange
    [Socket Null]
    [Socket CRLF]
    [Socket LF]
    [Socket CR]
End Enum

Public Enum UniSockMode
    [Socket Binary Mode]
    [Socket Text Mode]
End Enum

Public Enum UniSockTextFormat
    [Socket UTF-8]
    [Socket ANSI]
    [Socket UTF-16]
End Enum

Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Type SOCKADDR_IN
    Family As Integer
    Port As Integer
    Addr As Long
    Zero(7) As Byte
End Type

Private Type WNDCLASS
    Style As Long
    lpfnWndProc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As Long
    lpszClassName As Long
End Type

Private Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription(256) As Byte
    szSystemStatus(128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FindWindowW Lib "user32" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function FormatMessageW Lib "kernel32" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As Long, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function PostMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RegisterClassW Lib "user32" (Class As WNDCLASS) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function UnregisterClassW Lib "user32" (ByVal lpClassName As Long, ByVal hInstance As Long) As Long

Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (Destination As Any, Value As Any)
Private Declare Sub PutMem8 Lib "msvbvm60" (Destination As Any, Value As Any)
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SysAllocStringANSI Lib "oleaut32" Alias "SysAllocStringByteLen" (ByVal OleStr As Long, ByVal bLen As Long) As String
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal OleStr As Long, ByVal bLen As Long) As Long
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal OleStr As Long, ByVal bLen As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long

Private Declare Function SocketAccept Lib "ws2_32" Alias "accept" (ByVal S As Long, ByRef Addr As SOCKADDR_IN, ByRef AddrLen As Long) As Long
Private Declare Function SocketBind Lib "ws2_32" Alias "bind" (ByVal S As Long, ByRef Name As SOCKADDR_IN, ByRef NameLen As Long) As Long
Private Declare Function SocketClose Lib "ws2_32" Alias "closesocket" (ByVal S As Long) As Long
Private Declare Function SocketConnect Lib "ws2_32" Alias "connect" (ByVal S As Long, ByRef Name As SOCKADDR_IN, ByVal NameLen As Long) As Long
Private Declare Function SocketGetHostName Lib "ws2_32" Alias "gethostname" (ByVal Hostname As String, ByVal NameLen As Long) As Long
Private Declare Function SocketGetHostByName Lib "ws2_32" Alias "gethostbyname" (ByVal Hostname As String) As Long
Private Declare Function SocketGetIP Lib "ws2_32" Alias "inet_addr" (ByVal CP As String) As Long
Private Declare Function SocketGetName Lib "ws2_32" Alias "getsockname" (ByVal S As Long, ByRef Name As SOCKADDR_IN, ByRef NameLen As Long) As Long
Private Declare Function SocketGetNameInfo Lib "ws2_32" Alias "getnameinfo" (ByRef SA As SOCKADDR_IN, ByVal SALen As Long, ByVal Host As String, ByVal HostLen As Long, ByVal Serv As String, ByVal ServLen As Long, ByVal Flags As Long) As Long
Private Declare Function SocketGetOptions Lib "ws2_32" Alias "getsockopt" (ByVal S As Long, ByVal Level As Long, ByVal OptName As Long, OptVal As Any, OptLen As Long) As Long
Private Declare Function SocketGetPeerName Lib "ws2_32" Alias "getpeername" (ByVal S As Long, ByRef Name As SOCKADDR_IN, ByRef NameLen As Long) As Long
Private Declare Function SockethTons Lib "ws2_32" Alias "htons" (ByVal hostshort As Integer) As Integer
Private Declare Function SocketInet_nToa Lib "ws2_32" Alias "inet_ntoa" (ByVal in_addr As Long) As Long
Private Declare Function SocketInit Lib "ws2_32" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Private Declare Function SocketIOCtl Lib "ws2_32" Alias "ioctlsocket" (ByVal S As Long, ByVal Cmd As Long, ByRef ArgP As Long) As Long
Private Declare Function SocketListen Lib "ws2_32" Alias "listen" (ByVal S As Long, ByVal BackLog As Long) As Long
Private Declare Function SocketnTons Lib "ws2_32" Alias "htons" (ByVal netshort As Integer) As Integer
Private Declare Function SocketReceive Lib "ws2_32" Alias "recv" (ByVal S As Long, Buf As Any, ByVal BufLen As Long, ByVal Flags As Long) As Long
Private Declare Function SocketReceiveFrom Lib "ws2_32" Alias "recv" (ByVal S As Long, Buf As Any, ByVal BufLen As Long, ByVal Flags As Long, ByRef FromAddr As SOCKADDR_IN, ByVal FromLen As Long) As Long
Private Declare Function SocketSend Lib "ws2_32" Alias "send" (ByVal S As Long, Buf As Any, ByVal BufLen As Long, ByVal Flags As Long) As Long
Private Declare Function SocketSendTo Lib "ws2_32" Alias "sendto" (ByVal S As Long, Buf As Any, ByVal BufLen As Long, ByVal Flags As Long, ByRef ToAddr As SOCKADDR_IN, ByVal ToLen As Long) As Long
Private Declare Function SocketSetOptions Lib "ws2_32" Alias "setsockopt" (ByVal S As Long, ByVal Level As Long, ByVal OptName As Long, OptVal As Any, ByVal OptLen As Long) As Long

Private Declare Function WSAAsyncGetHostByName Lib "ws2_32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal strHostName As String, Buf As Any, ByVal BufLen As Long) As Long
Private Declare Function WSAAsyncSelect Lib "ws2_32" (ByVal S As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function WSACancelAsyncRequest Lib "ws2_32" (ByVal hAsyncTaskHandle As Long) As Long
Private Declare Function WSACleanup Lib "ws2_32" () As Long
Private Declare Function WSAGetLastError Lib "ws2_32" () As Long
Private Declare Function WSAStartup Lib "ws2_32" (ByVal wVersionRequired As Long, lpWSADATA As WSAData) As Long

Private Const AF_INET = &H2&
Private Const FD_READ = &H1&
Private Const FD_WRITE = &H2&
Private Const FD_ACCEPT = &H8&
Private Const FD_CONNECT = &H10&
Private Const FD_CLOSE = &H20&
Private Const FIONREAD = &H4004667F
Private Const INADDR_NONE = -1&
Private Const INVALID_SOCKET = -1&
Private Const IPPROTO_TCP = 6&
Private Const IPPROTO_UDP = 17&
Private Const MSG_PEEK = &H2&
Private Const SO_BROADCAST = &H20&
Private Const SO_MAX_MSG_SIZE = &H2003&
Private Const SO_RCVBUF = &H1002&
Private Const SO_SNDBUF = &H1001&
Private Const SOCK_DGRAM = 2&
Private Const SOCK_STREAM = 1&
Private Const SOCKET_ERR = -1&
Private Const SOL_SOCKET = &HFFFF&
Private Const SOMAXCONN = 5&

Private Const WM_USER = &H400
Private Const RESOLVE_MESSAGE = WM_USER + &H400
Private Const SOCKET_MESSAGE = WM_USER + &H401

Private m_AsyncHost As Long
Private m_Buffer As String
Private m_BufferSize As Long
Private m_Class As WNDCLASS
Private m_ClassName As String
Private m_HostentHandle As Long
Private m_HostentPointer As Long
Private m_hWnd As Long
Private m_Initialized As Boolean
Private m_LineChange As String
Private m_LocalIP As String
Private m_LocalPort As Long
Private m_LocalPortBind As Long
Private m_Mode As UniSockMode
Private m_Peek As Boolean
Private m_Protocol As ProtocolConstants
Private m_ReceiveSize As Long
Private m_RegisterClass As Long
Private m_RemoteHost As String
Private m_RemoteIP As Long
Private m_RemotePort As Long
Private m_Request As Collection
Private m_RequestBuffer As String
Private m_RequestID As Long
Private m_RequestSA As SOCKADDR_IN
Private m_SendBuffer As String
Private m_SendBufferSize As Long
Private m_SendSize As Long
Private m_Socket As Long
Private m_State As StateConstants
Private m_WSAData As WSAData

Private CRLF As String
Private CR As String
Private LF As String
Private NULLCHAR As String

' the following are the requirements for SelfCallback
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Byte)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Integer)
Private Declare Sub RtlMachineCodeCopy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Long, Source As MachineCode, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

' self-documentation: less comments and constants
Private Type MachineCode        ' 37 * 4 = 148 bytes
    ' thunk
    OwnerPtr As Long            ' 0
    CallbackAddress As Long     ' 1
    API_EbMode As Long          ' 2
    API_IsBadCodePtr As Long    ' 3
    API_KillTimer As Long       ' 4
    ' code
    MC1(5 To 5) As Long         ' 5
    AllocatedDataPtr As Long    ' 6
    MC2(7 To 17) As Long        ' 7 - 17
    ParamCount As Long          ' 18
    MC3(19 To 35) As Long       ' 19 - 35
    Ordinal As Long             ' 36
End Type

Public Function Accept(ByVal RequestID As Long) As Boolean
    Dim clsSocket As UniSock, clsTemp As UniSock, lngBuffer As Long, udtSockAddr As SOCKADDR_IN
    ' validate
    If (m_State = sckClosed) And (RequestID <> 0) Then
        ' ensure this RequestID is from an existing UniSock
        If FindWindowW(0, StrPtr(CStr(RequestID))) Then
            ' make sure socket is closed
            CloseSocket
            ' make a reference to the listening object
            PutMem4 clsTemp, ByVal RequestID
            Set clsSocket = clsTemp
            PutMem4 clsTemp, ByVal 0&
            ' now copy the data, get real RequestID etc.
            clsSocket.Friend_Request m_hWnd, m_RequestID, m_RequestSA, m_Buffer
            ' we do not need it for more than that
            Set clsSocket = Nothing
            ' correct some settings
            m_Socket = m_RequestID
            m_BufferSize = LenB(m_Buffer)
            m_Protocol = sckTCPProtocol
            If SocketGetOptions(m_RequestID, SOL_SOCKET, SO_RCVBUF, lngBuffer, 4) <> SOCKET_ERR Then m_ReceiveSize = lngBuffer
            If SocketGetOptions(m_RequestID, SOL_SOCKET, SO_SNDBUF, lngBuffer, 4) <> SOCKET_ERR Then m_SendSize = lngBuffer
            ' register messages
            If WSAAsyncSelect(m_RequestID, m_hWnd, SOCKET_MESSAGE, FD_READ Or FD_WRITE Or FD_ACCEPT Or FD_CONNECT Or FD_CLOSE) = SOCKET_ERR Then
                Private_SetState sckError, "Accept", WSAGetLastError
            End If
            ' local information
            If SocketGetName(m_RequestID, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR Then
                m_LocalIP = CStr(udtSockAddr.Addr And &HFF&) & "." & _
                    CStr((udtSockAddr.Addr And &HFF00&) \ &H100&) & "." & _
                    CStr((udtSockAddr.Addr And &HFF0000) \ &H10000) & "." & _
                    CStr(((udtSockAddr.Addr And &HFF000000) \ &H1000000) And &HFF&)
                m_LocalPortBind = SocketnTons(udtSockAddr.Port) And &HFFFF&
            Else
                Private_SetState sckError, "Accept", WSAGetLastError
            End If
            ' remote information
            m_RemoteHost = RequestHost
            m_RemoteIP = m_RequestSA.Addr
            m_RemotePort = SocketnTons(m_RequestSA.Port) And &HFFFF&
            ' done!
            m_RequestID = 0
            m_RequestSA = udtSockAddr
            ' status changes
            Private_SetState sckConnected
            RaiseEvent Connect
            ' last but not least, data has arrived
            If m_BufferSize Then Private_Message FD_READ
            ' success!
            Accept = True
        Else
            Err.Raise sckInvalidArg, "UniSock.Accept", Private_FormatWSAError(sckInvalidArg)
        End If
    Else
        Err.Raise sckInvalidOp, "UniSock.Accept", Private_FormatWSAError(sckInvalidOp)
    End If
End Function

Public Function Bind(Optional ByVal LocalPort As Long, Optional LocalIP As String) As Boolean
    Dim lngBuffer As Long
    If m_Initialized Then
        ' create socket if we do not have one
        If m_Socket = INVALID_SOCKET Then
            ' create socket for that protocol
            If m_Protocol = sckTCPProtocol Then
                m_Socket = SocketInit(AF_INET, SOCK_STREAM, IPPROTO_TCP)
                ' if we got a socket...
                If m_Socket <> INVALID_SOCKET Then
                    ' get the buffer sizes
                    If SocketGetOptions(m_Socket, SOL_SOCKET, SO_RCVBUF, lngBuffer, 4) <> SOCKET_ERR Then m_ReceiveSize = lngBuffer
                    If SocketGetOptions(m_Socket, SOL_SOCKET, SO_SNDBUF, lngBuffer, 4) <> SOCKET_ERR Then m_SendSize = lngBuffer
                    ' register messages
                    If WSAAsyncSelect(m_Socket, m_hWnd, SOCKET_MESSAGE, FD_READ Or FD_WRITE Or FD_ACCEPT Or FD_CONNECT Or FD_CLOSE) = SOCKET_ERR Then
                        Private_SetState sckError, "Bind", WSAGetLastError
                    End If
                End If
            Else
                m_Socket = SocketInit(AF_INET, SOCK_DGRAM, IPPROTO_UDP)
                ' if we get a socket...
                If m_Socket <> INVALID_SOCKET Then
                    lngBuffer = 1
                    SocketSetOptions m_Socket, SOL_SOCKET, SO_BROADCAST, lngBuffer, 4
                    ' get the buffer sizes
                    If SocketGetOptions(m_Socket, SOL_SOCKET, SO_MAX_MSG_SIZE, lngBuffer, 4) <> SOCKET_ERR Then
                        m_ReceiveSize = lngBuffer
                        m_SendSize = lngBuffer
                    End If
                    ' register messages
                    If WSAAsyncSelect(m_Socket, m_hWnd, SOCKET_MESSAGE, FD_READ Or FD_WRITE Or FD_ACCEPT Or FD_CONNECT Or FD_CLOSE) = SOCKET_ERR Then
                        Private_SetState sckError, "Bind", WSAGetLastError
                    End If
                End If
            End If
        End If
        ' start binding
        If m_Socket <> INVALID_SOCKET Then Bind = Private_Bind(LocalPort, LocalIP)
    End If
End Function

Public Function BytesReceived() As Long
    If m_Protocol = sckTCPProtocol Then
        BytesReceived = m_BufferSize
    ElseIf m_Socket Then
        SocketIOCtl m_Socket, FIONREAD, BytesReceived
    End If
End Function

Public Function CloseSocket() As Boolean
    Dim udtSockAddr As SOCKADDR_IN
    If m_State <> sckClosed Then
        Private_SetState sckClosing
        ' clear asynchronous process if active plus clean up the related memory
        If m_AsyncHost Then WSACancelAsyncRequest m_AsyncHost: m_AsyncHost = 0
        If m_HostentHandle Then GlobalFree m_HostentHandle: m_HostentHandle = 0: m_HostentPointer = 0
        ' now close the socket
        If m_Socket <> INVALID_SOCKET Then
            CloseSocket = (SocketClose(m_Socket) <> SOCKET_ERR)
            If Not CloseSocket Then
                ' something failed while closing the socket
                Private_SetState sckError, "CloseSocket", WSAGetLastError
            End If
        End If
        ' clean up
        m_Buffer = vbNullString
        m_BufferSize = 0
        m_LocalIP = vbNullString
        m_LocalPort = 0
        m_LocalPortBind = 0
        Set m_Request = New Collection
        m_RequestBuffer = vbNullString
        m_RequestID = 0
        m_RequestSA = udtSockAddr
        m_SendBuffer = vbNullString
        m_SendBufferSize = 0
        m_Socket = INVALID_SOCKET
        ' closed!
        Private_SetState sckClosed
    End If
End Function

Public Function Connect(Optional RemoteHost As String, Optional RemotePort As Long) As Boolean
    Dim lngBuffer As Long, lngIP As Long
    If m_Initialized Then
        ' remember given parameters
        If LenB(RemoteHost) Then m_RemoteHost = RemoteHost
        If RemotePort > 0 And RemotePort < 65536 Then m_RemotePort = RemotePort
        ' must have remote host
        If LenB(m_RemoteHost) Then
            ' create socket if we do not have one
            If m_Socket = INVALID_SOCKET Then
                ' create socket for that protocol
                If m_Protocol = sckTCPProtocol Then
                    m_Socket = SocketInit(AF_INET, SOCK_STREAM, IPPROTO_TCP)
                    ' if we got a socket...
                    If m_Socket <> INVALID_SOCKET Then
                        ' get the buffer sizes
                        If SocketGetOptions(m_Socket, SOL_SOCKET, SO_RCVBUF, lngBuffer, 4) <> SOCKET_ERR Then m_ReceiveSize = lngBuffer
                        If SocketGetOptions(m_Socket, SOL_SOCKET, SO_SNDBUF, lngBuffer, 4) <> SOCKET_ERR Then m_SendSize = lngBuffer
                        ' register messages
                        If WSAAsyncSelect(m_Socket, m_hWnd, SOCKET_MESSAGE, FD_READ Or FD_WRITE Or FD_ACCEPT Or FD_CONNECT Or FD_CLOSE) = SOCKET_ERR Then
                            Private_SetState sckError, "Connect", WSAGetLastError
                        End If
                    End If
                Else
                    m_Socket = SocketInit(AF_INET, SOCK_DGRAM, IPPROTO_UDP)
                    ' if we get a socket...
                    If m_Socket <> INVALID_SOCKET Then
                        lngBuffer = 1
                        SocketSetOptions m_Socket, SOL_SOCKET, SO_BROADCAST, lngBuffer, 4
                        ' get the buffer sizes
                        If SocketGetOptions(m_Socket, SOL_SOCKET, SO_MAX_MSG_SIZE, lngBuffer, 4) <> SOCKET_ERR Then
                            m_ReceiveSize = lngBuffer
                            m_SendSize = lngBuffer
                        End If
                        ' register messages
                        If WSAAsyncSelect(m_Socket, m_hWnd, SOCKET_MESSAGE, FD_READ Or FD_WRITE Or FD_ACCEPT Or FD_CONNECT Or FD_CLOSE) = SOCKET_ERR Then
                            Private_SetState sckError, "Connect", WSAGetLastError
                        End If
                    End If
                End If
            End If
            ' see if we not have a socket
            If m_Socket <> INVALID_SOCKET Then
                If m_Protocol = sckTCPProtocol Then
                    ' see if it is a string IP
                    lngIP = SocketGetIP(m_RemoteHost)
                    ' failed?
                    If lngIP = INADDR_NONE Then
                        ' we have to resolve the remote host...
                        If m_AsyncHost Then
                            ' cancel previous host query
                            WSACancelAsyncRequest m_AsyncHost
                            ' however, we keep using the previous memory block for HOSTENT
                        Else
                            ' get a new memory block
                            m_HostentHandle = GlobalAlloc(0, 1024)
                        End If
                        If m_HostentHandle Then
                            ' try to get the pointer...
                            m_HostentPointer = GlobalLock(m_HostentHandle)
                            If m_HostentPointer Then
                                GlobalUnlock m_HostentHandle
                                ' success, now we can resolve the host asynchronously
                                m_AsyncHost = WSAAsyncGetHostByName(m_hWnd, RESOLVE_MESSAGE, m_RemoteHost, ByVal m_HostentPointer, 1024)
                                ' success if we got that
                                Connect = m_AsyncHost <> 0
                                ' see which status we raise
                                If Connect Then
                                    Private_SetState sckResolvingHost
                                Else
                                    Private_SetState sckError
                                End If
                            Else
                                ' failed, free it up
                                GlobalFree m_HostentHandle
                                m_HostentHandle = 0
                                Private_SetState sckError, "Connect", GetLastError, "Failed to get pointer to allocated memory."
                            End If
                        Else
                            ' we failed to get memory for the asynchronous process
                            Private_SetState sckError, "Connect", GetLastError, "Could not allocate memory for asynchronous process."
                        End If
                    Else
                        ' we got a valid IP, try to connect!
                        Connect = ConnectIP(lngIP)
                    End If
                Else
                    ' UDP is a connectionless protocol
                    Connect = Private_Bind
                End If
            Else
                ' could not create socket
                Private_SetState sckError, "Connect", WSAGetLastError
            End If
        Else
            ' cannot start connection thanks to invalid settings
            Private_SetState sckError, "Connect"
        End If
    Else
        ' we do not even have all the requirements!
        Private_SetState sckError, "Connect"
    End If
End Function

Private Function ConnectIP(ByVal RemoteIP As Long) As Boolean
    Dim lngError As Long, udtSockAddr As SOCKADDR_IN
    If m_Initialized Then
        ' validate port
        If m_Socket <> INVALID_SOCKET And m_RemotePort > 0 And m_RemotePort < 65536 Then
            Private_SetState sckConnecting
            ' remember IP
            m_RemoteIP = RemoteIP
            With udtSockAddr
                .Addr = RemoteIP
                .Family = AF_INET
                .Port = SockethTons(m_RemotePort And &HFFFF)
            End With
            ' connect!
            ConnectIP = (SocketConnect(m_Socket, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
            If Not ConnectIP Then
                lngError = WSAGetLastError
                ' ignore "resource not available"
                If lngError <> sckWouldBlock Then
                    ' then imitate Winsock control
                    If lngError = sckAddressNotAvailable Then
                        Err.Raise lngError, "UniSock.ConnectIP", Private_FormatWSAError(lngError)
                    Else
                        Private_SetState sckError, "ConnectIP", lngError
                    End If
                Else
                    ' and thus we succeed
                    ConnectIP = True
                End If
            End If
        Else
            Private_SetState sckError, "ConnectIP"
        End If
    Else
        ' we do not even have all the requirements!
        Private_SetState sckError, "ConnectIP"
    End If
End Function

Friend Sub Friend_Request(ByVal hWnd As Long, RequestID As Long, SA As SOCKADDR_IN, Buffer As String)
    Dim udtSockAddr As SOCKADDR_IN
    If m_RequestID Then
        ' remember hWnd
        On Error Resume Next: m_Request.Add hWnd, CStr(m_RequestID): On Error GoTo 0
        ' copy there
        RequestID = m_RequestID
        SA = m_RequestSA
        Buffer = m_RequestBuffer
        ' clear here
        m_RequestID = 0
        m_RequestSA = udtSockAddr
        m_RequestBuffer = vbNullString
    End If
End Sub

Public Function GetData(Data, Optional VarType As VbVarType = vbByte Or vbArray, Optional ByVal MaxLen As Long) As Boolean
    Dim blnArray() As Boolean, bytArray() As Byte, curArray() As Currency, dtmArray() As Date, dblArray() As Double, intArray() As Integer, lngArray() As Long, sngArray() As Single
    Dim blnData As Boolean, bytData As Byte, curData As Currency, dtmData As Date, dblData As Double, intData As Integer, lngData As Long, sngData As Single, strData As String
    Dim lngFlags As Long, lngLen As Long, lngPtr As Long, udtSockAddr As SOCKADDR_IN
    ' which datatype to use?
    If VBA.VarType(Data) <> vbEmpty Then VarType = VBA.VarType(Data)
    ' which protocol we are using?
    If (m_Protocol = sckTCPProtocol) And (m_State = sckConnected) And (m_Mode = [Socket Binary Mode]) Then
        ' how much data to get?
        If MaxLen < 1 Then
            MaxLen = m_BufferSize
        ElseIf MaxLen > m_BufferSize Then
            MaxLen = m_BufferSize
        End If
        ' see if we get data at all
        If MaxLen Then
            ' now decide what to do
            If VarType And vbArray Then
                Select Case (VarType And Not vbArray)
                    Case vbBoolean
                        ReDim blnArray((MaxLen - 1) \ 2)
                        RtlMoveMemory blnArray(0), ByVal StrPtr(m_Buffer), MaxLen
                        Data = blnArray
                        GetData = True
                    Case vbByte
                        bytArray = LeftB$(m_Buffer, MaxLen)
                        Data = bytArray
                        GetData = True
                    Case vbCurrency
                        ReDim curArray((MaxLen - 1) \ 8)
                        RtlMoveMemory curArray(0), ByVal StrPtr(m_Buffer), MaxLen
                        Data = curArray
                        GetData = True
                    Case vbDate
                        ReDim dtmArray((MaxLen - 1) \ 8)
                        RtlMoveMemory dtmArray(0), ByVal StrPtr(m_Buffer), MaxLen
                        Data = dtmArray
                        GetData = True
                    Case vbDouble
                        ReDim dblArray((MaxLen - 1) \ 8)
                        RtlMoveMemory dblArray(0), ByVal StrPtr(m_Buffer), MaxLen
                        Data = dblArray
                        GetData = True
                    Case vbInteger
                        ReDim intArray((MaxLen - 1) \ 2)
                        RtlMoveMemory intArray(0), ByVal StrPtr(m_Buffer), MaxLen
                        Data = intArray
                        GetData = True
                    Case vbLong
                        ReDim lngArray((MaxLen - 1) \ 4)
                        RtlMoveMemory lngArray(0), ByVal StrPtr(m_Buffer), MaxLen
                        Data = lngArray
                        GetData = True
                    Case vbSingle
                        ReDim sngArray((MaxLen - 1) \ 4)
                        RtlMoveMemory sngArray(0), ByVal StrPtr(m_Buffer), MaxLen
                        Data = sngArray
                        GetData = True
                    Case vbString
                        strData = StrConv(LeftB$(m_Buffer, MaxLen), vbUnicode)
                        ' we make a string array simply by looking for line changes
                        If InStr(strData, vbNewLine) Then
                            Data = Split(strData, vbNewLine)
                        ElseIf InStr(strData, vbLf) Then
                            Data = Split(strData, vbLf)
                        ElseIf InStr(strData, vbCr) Then
                            Data = Split(strData, vbCr)
                        Else
                            Data = Split(strData, vbNullChar)
                        End If
                        GetData = True
                    Case Else
                        Err.Raise sckUnsupported, "UniSock.GetData", Private_FormatWSAError(sckUnsupported)
                End Select
            Else
                Select Case VarType
                    Case vbBoolean
                        If MaxLen > 2 Then MaxLen = 2
                        RtlMoveMemory blnData, ByVal StrPtr(m_Buffer), MaxLen
                        GetData = True
                    Case vbByte
                        MaxLen = 1
                        RtlMoveMemory bytData, ByVal StrPtr(m_Buffer), MaxLen
                        GetData = True
                    Case vbCurrency
                        If MaxLen > 8 Then MaxLen = 8
                        RtlMoveMemory curData, ByVal StrPtr(m_Buffer), MaxLen
                        GetData = True
                    Case vbDate
                        If MaxLen > 8 Then MaxLen = 8
                        RtlMoveMemory dtmData, ByVal StrPtr(m_Buffer), MaxLen
                        GetData = True
                    Case vbDouble
                        If MaxLen > 8 Then MaxLen = 8
                        RtlMoveMemory dblData, ByVal StrPtr(m_Buffer), MaxLen
                        GetData = True
                    Case vbInteger
                        If MaxLen > 2 Then MaxLen = 2
                        RtlMoveMemory intData, ByVal StrPtr(m_Buffer), MaxLen
                        GetData = True
                    Case vbLong
                        If MaxLen > 4 Then MaxLen = 4
                        RtlMoveMemory lngData, ByVal StrPtr(m_Buffer), MaxLen
                        GetData = True
                    Case vbSingle
                        If MaxLen > 4 Then MaxLen = 4
                        RtlMoveMemory sngData, ByVal StrPtr(m_Buffer), MaxLen
                        GetData = True
                    Case vbString
                        Data = StrConv(LeftB$(m_Buffer, MaxLen), vbUnicode)
                        GetData = True
                    Case Else
                        Err.Raise sckUnsupported, "UniSock.GetData", Private_FormatWSAError(sckUnsupported)
                End Select
            End If
            ' if we got the data and if not in peek mode...
            If GetData And Not m_Peek Then
                ' remove the data we got from the buffer
                m_Buffer = MidB$(m_Buffer, MaxLen + 1)
                m_BufferSize = m_BufferSize - MaxLen
            End If
        End If
    ElseIf (m_Protocol = sckUDPProtocol) And (m_State = sckOpen) Then
        ' peek mode?
        If m_Peek Then lngFlags = MSG_PEEK
        ' how much data to get?
        SocketIOCtl m_Socket, FIONREAD, lngLen
        If MaxLen < 1 Then
            MaxLen = lngLen
        ElseIf MaxLen > lngLen Then
            MaxLen = lngLen
        End If
        ' see if we get data at all
        If MaxLen Then
            ' now decide what to do
            If VarType And vbArray Then
                Select Case (VarType And Not vbArray)
                    Case vbBoolean
                        ReDim blnArray((MaxLen - 1) \ 2)
                        GetData = (SocketReceiveFrom(m_Socket, blnArray(0), MaxLen, lngFlags, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
                        If GetData Then
                            Data = blnArray
                        Else
                            Private_SetState sckError, "GetData", WSAGetLastError
                        End If
                    Case vbByte
                        ReDim bytArray(MaxLen - 1)
                        GetData = (SocketReceiveFrom(m_Socket, bytArray(0), MaxLen, lngFlags, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
                        If GetData Then
                            Data = bytArray
                        Else
                            Private_SetState sckError, "GetData", WSAGetLastError
                        End If
                    Case vbCurrency
                        ReDim curArray((MaxLen - 1) \ 8)
                        GetData = (SocketReceiveFrom(m_Socket, curArray(0), MaxLen, lngFlags, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
                        If GetData Then
                            Data = curArray
                        Else
                            Private_SetState sckError, "GetData", WSAGetLastError
                        End If
                    Case vbDate
                        ReDim dtmArray((MaxLen - 1) \ 8)
                        GetData = (SocketReceiveFrom(m_Socket, dtmArray(0), MaxLen, lngFlags, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
                        If GetData Then
                            Data = dtmArray
                        Else
                            Private_SetState sckError, "GetData", WSAGetLastError
                        End If
                    Case vbDouble
                        ReDim dblArray((MaxLen - 1) \ 8)
                        GetData = (SocketReceiveFrom(m_Socket, dblArray(0), MaxLen, lngFlags, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
                        If GetData Then
                            Data = dblArray
                        Else
                            Private_SetState sckError, "GetData", WSAGetLastError
                        End If
                    Case vbInteger
                        ReDim intArray((MaxLen - 1) \ 2)
                        GetData = (SocketReceiveFrom(m_Socket, intArray(0), MaxLen, lngFlags, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
                        If GetData Then
                            Data = intArray
                        Else
                            Private_SetState sckError, "GetData", WSAGetLastError
                        End If
                    Case vbLong
                        ReDim lngArray((MaxLen - 1) \ 4)
                        GetData = (SocketReceiveFrom(m_Socket, lngArray(0), MaxLen, lngFlags, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
                        If GetData Then
                            Data = lngArray
                        Else
                            Private_SetState sckError, "GetData", WSAGetLastError
                        End If
                    Case vbSingle
                        ReDim sngArray((MaxLen - 1) \ 4)
                        GetData = (SocketReceiveFrom(m_Socket, sngArray(0), MaxLen, lngFlags, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
                        If GetData Then
                            Data = sngArray
                        Else
                            Private_SetState sckError, "GetData", WSAGetLastError
                        End If
                    Case vbString
                        lngPtr = SysAllocStringByteLen(0, MaxLen)
                        PutMem4 ByVal VarPtr(strData), ByVal lngPtr
                        GetData = (SocketReceiveFrom(m_Socket, ByVal lngPtr, MaxLen, lngFlags, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
                        If GetData Then
                            ' we make a string array simply by looking for line changes
                            If InStrB(strData, CRLF) Then
                                Data = Split(StrConv(strData, vbUnicode), vbNewLine)
                            ElseIf InStrB(strData, LF) Then
                                Data = Split(StrConv(strData, vbUnicode), vbLf)
                            ElseIf InStrB(strData, CR) Then
                                Data = Split(StrConv(strData, vbUnicode), vbCr)
                            Else
                                Data = Split(StrConv(strData, vbUnicode), vbNullChar)
                            End If
                        Else
                            Private_SetState sckError, "GetData", WSAGetLastError
                        End If
                    Case Else
                        Err.Raise sckUnsupported, "UniSock.GetData", Private_FormatWSAError(sckUnsupported)
                End Select
            Else
                Select Case VarType
                    Case vbBoolean
                        If MaxLen > 2 Then MaxLen = 2
                        GetData = (SocketReceiveFrom(m_Socket, blnData, MaxLen, lngFlags, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
                        If GetData Then
                            Data = blnData
                        Else
                            Private_SetState sckError, "GetData", WSAGetLastError
                        End If
                    Case vbByte
                        MaxLen = 1
                        GetData = (SocketReceiveFrom(m_Socket, bytData, MaxLen, lngFlags, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
                        If GetData Then
                            Data = bytData
                        Else
                            Private_SetState sckError, "GetData", WSAGetLastError
                        End If
                    Case vbCurrency
                        If MaxLen > 8 Then MaxLen = 8
                        GetData = (SocketReceiveFrom(m_Socket, curData, MaxLen, lngFlags, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
                        If GetData Then
                            Data = curData
                        Else
                            Private_SetState sckError, "GetData", WSAGetLastError
                        End If
                    Case vbDate
                        If MaxLen > 8 Then MaxLen = 8
                        GetData = (SocketReceiveFrom(m_Socket, dtmData, MaxLen, lngFlags, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
                        If GetData Then
                            Data = dtmData
                        Else
                            Private_SetState sckError, "GetData", WSAGetLastError
                        End If
                    Case vbDouble
                        If MaxLen > 8 Then MaxLen = 8
                        GetData = (SocketReceiveFrom(m_Socket, dblData, MaxLen, lngFlags, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
                        If GetData Then
                            Data = dblData
                        Else
                            Private_SetState sckError, "GetData", WSAGetLastError
                        End If
                    Case vbInteger
                        If MaxLen > 2 Then MaxLen = 2
                        GetData = (SocketReceiveFrom(m_Socket, intData, MaxLen, lngFlags, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
                        If GetData Then
                            Data = intData
                        Else
                            Private_SetState sckError, "GetData", WSAGetLastError
                        End If
                    Case vbLong
                        If MaxLen > 4 Then MaxLen = 4
                        GetData = (SocketReceiveFrom(m_Socket, lngData, MaxLen, lngFlags, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
                        If GetData Then
                            Data = lngData
                        Else
                            Private_SetState sckError, "GetData", WSAGetLastError
                        End If
                    Case vbSingle
                        If MaxLen > 4 Then MaxLen = 4
                        GetData = (SocketReceiveFrom(m_Socket, sngData, MaxLen, lngFlags, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
                        If GetData Then
                            Data = sngData
                        Else
                            Private_SetState sckError, "GetData", WSAGetLastError
                        End If
                    Case vbString
                        lngPtr = SysAllocStringByteLen(0, MaxLen)
                        PutMem4 ByVal VarPtr(strData), ByVal lngPtr
                        GetData = (SocketReceiveFrom(m_Socket, ByVal lngPtr, MaxLen, lngFlags, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
                        If GetData Then
                            Data = StrConv(strData, vbUnicode)
                        Else
                            Private_SetState sckError, "GetData", WSAGetLastError
                        End If
                    Case Else
                        Err.Raise sckUnsupported, "UniSock.GetData", Private_FormatWSAError(sckUnsupported)
                End Select
            End If
        End If
    Else
        Err.Raise sckBadState, "UniSock.GetData", Private_FormatWSAError(sckBadState)
    End If
End Function

Public Property Get LineChange() As UniSockLineChange
    If LenB(m_LineChange) = 4 Then
        LineChange = [Socket CRLF]
    ElseIf m_LineChange = vbLf Then
        LineChange = [Socket LF]
    ElseIf m_LineChange = vbCr Then
        LineChange = [Socket CR]
    Else
        LineChange = [Socket Null]
    End If
End Property

Public Property Let LineChange(ByVal NewValue As UniSockLineChange)
    Select Case NewValue
        Case [Socket CRLF]
            m_LineChange = vbNewLine
        Case [Socket LF]
            m_LineChange = vbLf
        Case [Socket CR]
            m_LineChange = vbCr
        Case [Socket Null]
            m_LineChange = vbNullChar
        Case Else
            Err.Raise sckInvalidArg, "UniSock.LineChange", Private_FormatWSAError(sckInvalidArg)
    End Select
End Property

Public Function Listen() As Boolean
    Dim lngError As Long
    If (m_State = sckClosed) Or (m_State = sckOpen) Then
        ' first we must bind
        If Bind Then
            ' try to start listening
            Listen = (SocketListen(m_Socket, SOMAXCONN) <> SOCKET_ERR)
            If Not Listen Then
                ' failed
                lngError = WSAGetLastError
                Err.Raise lngError, "UniSock.Listen", Private_FormatWSAError(lngError)
            End If
        End If
    Else
        Err.Raise sckInvalidOp, "UniSock.Listen", Private_FormatWSAError(sckInvalidOp)
    End If
End Function

Public Function LocalHostname() As String
    LocalHostname = Private_LocalHostName
End Function

Public Function LocalIP() As String
    Dim lngLocalIP As Long
    If (m_State = sckOpen) Or (m_State = sckListening) Then
        LocalIP = m_LocalIP
    Else
        lngLocalIP = Private_LocalIP
        LocalIP = CStr(lngLocalIP And &HFF&) & "." & _
            CStr((lngLocalIP And &HFF00&) \ &H100&) & "." & _
            CStr((lngLocalIP And &HFF0000) \ &H10000) & "." & _
            CStr(((lngLocalIP And &HFF000000) \ &H1000000) And &HFF&)
    End If
End Function

Public Property Get LocalPort() As Long
    If m_LocalPortBind = 0 Then
        LocalPort = m_LocalPort
    Else
        LocalPort = m_LocalPortBind
    End If
End Property

Public Property Let LocalPort(ByVal NewValue As Long)
    If m_State = sckClosed Then
        If NewValue >= 0 And NewValue < 65536 Then
            m_LocalPort = NewValue
        Else
            Err.Raise sckInvalidArg, "UniSock.LocalPort", Private_FormatWSAError(sckInvalidArg)
        End If
    Else
        Err.Raise sckInvalidOp, "UniSock.LocalPort", Private_FormatWSAError(sckInvalidOp)
    End If
End Property

Public Property Get Mode() As UniSockMode
    Mode = m_Mode
End Property

Public Property Let Mode(ByVal NewValue As UniSockMode)
    m_Mode = NewValue
End Property

Public Function PeekData(Data, Optional VarType As VbVarType, Optional ByVal MaxLen As Long) As Boolean
    m_Peek = True
    PeekData = GetData(Data, VarType, MaxLen)
    m_Peek = False
End Function

Private Function Private_Bind(Optional ByVal LocalPort As Long, Optional LocalIP As String) As Boolean
    Dim lngLocalPort As Long, lngLocalIP As String, strIP As String, udtSockAddr As SOCKADDR_IN
    ' first of all check if we have bind already
    Private_Bind = (m_State = sckOpen)
    ' and if we have not...
    If (Not Private_Bind) And (m_Socket <> INVALID_SOCKET) Then
        ' determine local port to use
        If LocalPort > 0 And LocalPort < 65536 Then
            lngLocalPort = LocalPort
        Else
            lngLocalPort = m_LocalPort
        End If
        ' determine local ip
        lngLocalIP = Private_LocalIP(LocalIP)
        ' make a string too
        strIP = CStr(lngLocalIP And &HFF&) & "." & _
            CStr((lngLocalIP And &HFF00&) \ &H100&) & "." & _
            CStr((lngLocalIP And &HFF0000) \ &H10000) & "." & _
            CStr(((lngLocalIP And &HFF000000) \ &H1000000) And &HFF&)
        ' fill the structure for bind
        With udtSockAddr
            .Addr = lngLocalIP
            .Family = AF_INET
            .Port = SocketnTons(lngLocalPort And &HFFFF)
        End With
        ' bind it
        Private_Bind = (SocketBind(m_Socket, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR)
        If Private_Bind Then
            ' remember IP
            m_LocalIP = strIP
            ' see if we had valid port
            If lngLocalPort Then
                m_LocalPort = lngLocalPort
                Private_SetState sckOpen
            Else
                ' get the local port
                lngLocalPort = Private_LocalPort
                If lngLocalPort <> SOCKET_ERR Then
                    m_LocalPortBind = lngLocalPort
                Else
                    Private_SetState sckError, "Private_Bind", WSAGetLastError
                End If
            End If
        Else
            Private_SetState sckError, "Private_Bind", WSAGetLastError
        End If
    End If
End Function

Private Function Private_FormatError(ByVal Number As Long) As String
    Dim lngResult As Long, lngStrPtr As Long, strError As String
    If Number Then
        ' initialize the buffer
        lngStrPtr = SysAllocStringLen(0, 65535)
        PutMem4 ByVal VarPtr(strError), ByVal lngStrPtr
        ' get the error message
        lngResult = FormatMessageW(&H1000&, ByVal 0&, Number, 0, lngStrPtr, 65536, ByVal 0&)
        ' if we got a string, place it in the string
        If lngResult Then PutMem4 ByVal VarPtr(Private_FormatError), ByVal SysAllocStringLen(lngStrPtr, lngResult)
    End If
End Function

Private Function Private_FormatWSAError(ByVal Number As Long) As String
    Select Case Number
        ' WSA_INVALID_HANDLE
        Case 6
            Private_FormatWSAError = "Specified event object handle is invalid."
        ' WSA_NOT_ENOUGH_MEMORY
        Case 8
            Private_FormatWSAError = "Insufficient memory available."
        ' WSA_INVALID_PARAMETER
        Case 87
            Private_FormatWSAError = "One or more parameters are invalid."
        ' WSA_OPERATION_ABORTED
        Case 995
            Private_FormatWSAError = "Overlapped operation aborted."
        ' WSA_IO_INCOMPLETE
        Case 996
            Private_FormatWSAError = "Overlapped I/O event object not in signaled state."
        ' WSA_IO_PENDING
        Case 997
            Private_FormatWSAError = "Overlapped operations will complete later."
        ' WSAEINTR
        Case 10004
            Private_FormatWSAError = "Interrupted function call."
        ' WSAEBADF
        Case 10009
            Private_FormatWSAError = "File handle is not valid."
        ' WSAEACCES
        Case 10013
            Private_FormatWSAError = "Permission denied."
        ' WSAEFAULT
        Case 10014
            Private_FormatWSAError = "Bad address."
        ' WSAEINVAL
        Case 10022
            Private_FormatWSAError = "Invalid argument."
        ' WSAEMFILE
        Case 10024
            Private_FormatWSAError = "Too many open files."
        ' WSAEWOULDBLOCK
        Case 10035
            Private_FormatWSAError = "Resource temporarily unavailable."
        ' WSAEINPROGRESS
        Case 10036
            Private_FormatWSAError = "Operation now in progress."
        ' WSAEALREADY
        Case 10037
            Private_FormatWSAError = "Operation already in progress."
        ' WSAENOTSOCK
        Case 10038
            Private_FormatWSAError = "Socket operation on nonsocket."
        ' WSAEDESTADDRREQ
        Case 10039
            Private_FormatWSAError = "Destination address required."
        ' WSAEMSGSIZE
        Case 10040
            Private_FormatWSAError = "Message too long."
        ' WSAEPROTOTYPE
        Case 10041
            Private_FormatWSAError = "Protocol wrong type for socket."
        ' WSAENOPROTOOPT
        Case 10042
            Private_FormatWSAError = "Bad protocol option."
        ' WSAEPROTONOSUPPORT
        Case 10043
            Private_FormatWSAError = "Protocol not supported."
        ' WSAESOCKTNOSUPPORT
        Case 10044
            Private_FormatWSAError = "Socket type not supported."
        ' WSAEOPNOTSUPP
        Case 10045
            Private_FormatWSAError = "Operation not supported."
        ' WSAEPFNOSUPPORT
        Case 10046
            Private_FormatWSAError = "Protocol family not supported."
        ' WSAEAFNOSUPPORT
        Case 10047
            Private_FormatWSAError = "Address family not supported by protocol family."
        ' WSAEADDRINUSE
        Case 10048
            Private_FormatWSAError = "Address already in use."
        ' WSAEADDRNOTAVAIL
        Case 10049
            Private_FormatWSAError = "Cannot assign requested address."
        ' WSAENETDOWN
        Case 10050
            Private_FormatWSAError = "Network is down."
        ' WSAENETUNREACH
        Case 10051
            Private_FormatWSAError = "Network is unreachable."
        ' WSAENETRESET
        Case 10052
            Private_FormatWSAError = "Network dropped connection on reset."
        ' WSAECONNABORTED
        Case 10053
            Private_FormatWSAError = "Software caused connection abort."
        ' WSAECONNRESET
        Case 10054
            Private_FormatWSAError = "Connection reset by peer."
        ' WSAENOBUFS
        Case 10055
            Private_FormatWSAError = "No buffer space available."
        ' WSAEISCONN
        Case 10056
            Private_FormatWSAError = "Socket is already connected."
        ' WSAENOTCONN
        Case 10057
            Private_FormatWSAError = "Socket is not connected."
        ' WSAESHUTDOWN
        Case 10058
            Private_FormatWSAError = "Cannot send after socket shutdown."
        ' WSAETOOMANYREFS
        Case 10059
            Private_FormatWSAError = "Too many references."
        ' WSAETIMEDOUT
        Case 10060
            Private_FormatWSAError = "Connection timed out."
        ' WSAECONNREFUSED
        Case 10061
            Private_FormatWSAError = "Connection refused."
        ' WSAELOOP
        Case 10062
            Private_FormatWSAError = "Cannot translate name."
        ' WSAENAMETOOLONG
        Case 10063
            Private_FormatWSAError = "Name too long."
        ' WSAEHOSTDOWN
        Case 10064
            Private_FormatWSAError = "Host is down."
        ' WSAEHOSTUNREACH
        Case 10065
            Private_FormatWSAError = "No route to host."
        ' WSAENOTEMPTY
        Case 10066
            Private_FormatWSAError = "Directory not empty."
        ' WSAEPROCLIM
        Case 10067
            Private_FormatWSAError = "Too many processes."
        ' WSAEUSERS
        Case 10068
            Private_FormatWSAError = "User quota exceeded."
        ' WSAEDQUOT
        Case 10069
            Private_FormatWSAError = "Disk quota exceeded."
        ' WSAESTALE
        Case 10070
            Private_FormatWSAError = "Stale file handle reference."
        ' WSAEREMOTE
        Case 10071
            Private_FormatWSAError = "Item is remote."
        ' WSASYSNOTREADY
        Case 10091
            Private_FormatWSAError = "Network subsystem is unavailable."
        ' WSAVERNOTSUPPORTED
        Case 10092
            Private_FormatWSAError = "Winsock.dll version out of range."
        ' WSANOTINITIALISED
        Case 10093
            Private_FormatWSAError = "Successful WSAStartup not yet performed."
        ' WSAEDISCON
        Case 10101
            Private_FormatWSAError = "Graceful shutdown in progress."
        ' WSAENOMORE
        Case 10102
            Private_FormatWSAError = "No more results."
        ' WSAECANCELLED
        Case 10103
            Private_FormatWSAError = "Call has been canceled."
        ' WSAEINVALIDPROCTABLE
        Case 10104
            Private_FormatWSAError = "Procedure call table is invalid."
        ' WSAEINVALIDPROVIDER
        Case 10105
            Private_FormatWSAError = "Service provider is invalid."
        ' WSAEPROVIDERFAILEDINIT
        Case 10106
            Private_FormatWSAError = "Service provider failed to initialize."
        ' WSASYSCALLFAILURE
        Case 10107
            Private_FormatWSAError = "System call failure."
        ' WSASERVICE_NOT_FOUND
        Case 10108
            Private_FormatWSAError = "Service not found."
        ' WSATYPE_NOT_FOUND
        Case 10109
            Private_FormatWSAError = "Class type not found."
        ' WSA_E_NO_MORE
        Case 10110
            Private_FormatWSAError = "No more results."
        ' WSA_E_CANCELLED
        Case 10111
            Private_FormatWSAError = "Call was canceled."
        ' WSAEREFUSED
        Case 10112
            Private_FormatWSAError = "Database query was refused."
        ' WSAHOST_NOT_FOUND
        Case 11001
            Private_FormatWSAError = "Host not found."
        ' WSATRY_AGAIN
        Case 11002
            Private_FormatWSAError = "Nonauthoritative host not found."
        ' WSANO_RECOVERY
        Case 11003
            Private_FormatWSAError = "This is a nonrecoverable error."
        ' WSANO_DATA
        Case 11004
            Private_FormatWSAError = "Valid name, no data record of requested type."
        ' WSA_QOS_RECEIVERS
        Case 11005
            Private_FormatWSAError = "QOS receivers."
        ' WSA_QOS_SENDERS
        Case 11006
            Private_FormatWSAError = "QOS senders."
        ' WSA_QOS_NO_SENDERS
        Case 11007
            Private_FormatWSAError = "No QOS senders."
        ' WSA_QOS_NO_RECEIVERS
        Case 11008
            Private_FormatWSAError = "QOS no receivers."
        ' WSA_QOS_REQUEST_CONFIRMED
        Case 11009
            Private_FormatWSAError = "QOS request confirmed."
        ' WSA_QOS_ADMISSION_FAILURE
        Case 11010
            Private_FormatWSAError = "QOS admission error."
        ' WSA_QOS_POLICY_FAILURE
        Case 11011
            Private_FormatWSAError = "QOS policy failure."
        ' WSA_QOS_BAD_STYLE
        Case 11012
            Private_FormatWSAError = "QOS bad style."
        ' WSA_QOS_BAD_OBJECT
        Case 11013
            Private_FormatWSAError = "QOS bad object."
        ' WSA_QOS_TRAFFIC_CTRL_ERROR
        Case 11014
            Private_FormatWSAError = "QOS traffic control error."
        ' WSA_QOS_GENERIC_ERROR
        Case 11015
            Private_FormatWSAError = "QOS generic error."
        ' WSA_QOS_ESERVICETYPE
        Case 11016
            Private_FormatWSAError = "QOS service type error."
        ' WSA_QOS_EFLOWSPEC
        Case 11017
            Private_FormatWSAError = "QOS flowspec error."
        ' WSA_QOS_EPROVSPECBUF
        Case 11018
            Private_FormatWSAError = "Invalid QOS provider buffer."
        ' WSA_QOS_EFILTERSTYLE
        Case 11019
            Private_FormatWSAError = "Invalid QOS filter style."
        ' WSA_QOS_EFILTERTYPE
        Case 11020
            Private_FormatWSAError = "Invalid QOS filter type."
        ' WSA_QOS_EFILTERCOUNT
        Case 11021
            Private_FormatWSAError = "Incorrect QOS filter count."
        ' WSA_QOS_EOBJLENGTH
        Case 11022
            Private_FormatWSAError = "Invalid QOS object length."
        ' WSA_QOS_EFLOWCOUNT
        Case 11023
            Private_FormatWSAError = "Incorrect QOS flow count."
        ' WSA_QOS_EUNKOWNPSOBJ
        Case 11024
            Private_FormatWSAError = "Unrecognized QOS object."
        ' WSA_QOS_EPOLICYOBJ
        Case 11025
            Private_FormatWSAError = "Invalid QOS policy object."
        ' WSA_QOS_EFLOWDESC
        Case 11026
            Private_FormatWSAError = "Invalid QOS flow descriptor."
        ' WSA_QOS_EPSFLOWSPEC
        Case 11027
            Private_FormatWSAError = "Invalid QOS provider-specific flowspec."
        ' WSA_QOS_EPSFILTERSPEC
        Case 11028
            Private_FormatWSAError = "Invalid QOS provider-specific filterspec."
        ' WSA_QOS_ESDMODEOBJ
        Case 11029
            Private_FormatWSAError = "Invalid QOS shape discard mode object."
        ' WSA_QOS_ESHAPERATEOBJ
        Case 11030
            Private_FormatWSAError = "Invalid QOS shaping rate object."
        ' WSA_QOS_RESERVED_PETYPE
        Case 11031
            Private_FormatWSAError = "Reserved policy QOS element type."
        Case sckBadState
            Private_FormatWSAError = "Wrong protocol or connection state for the requested transaction or request."
        Case sckGetNotSupported
            Private_FormatWSAError = "Property is write-only."
        Case sckInvalidArg
            Private_FormatWSAError = "The argument passed to a function was not in the correct format or in the specified range."
        Case sckInvalidOp
            Private_FormatWSAError = "Invalid operation at current state."
        Case sckInvalidPropertyValue
            Private_FormatWSAError = "Invalid property value."
        Case sckOutOfMemory
            Private_FormatWSAError = "Out of memory."
        Case sckOutOfRange
            Private_FormatWSAError = "Argument is out of range."
        Case sckSetNotSupported
            Private_FormatWSAError = "Property is read-only."
        Case sckSuccess
            Private_FormatWSAError = "Successful."
        Case sckUnsupported
            Private_FormatWSAError = "Unsupported variant types."
        Case sckWrongProtocol
            Private_FormatWSAError = "Wrong protocol for the requested transaction or request."
    End Select
End Function

Private Function Private_IsUTF8(ByRef Text As String, Optional ByVal lngReadSize As Long = 2048) As Boolean
    Dim bytArray() As Byte, lngArraySize As Long, lngBytes As Long, lngPos As Long, lngUTF8 As Long
    lngArraySize = LenB(Text)
    If lngArraySize Then
        If lngReadSize > 0 Then
            bytArray = Text
            If lngReadSize > lngArraySize Then lngReadSize = lngArraySize
            Do While lngPos < lngReadSize
                If bytArray(lngPos) < &H80 Then
                    lngPos = lngPos + 1
                ElseIf bytArray(lngPos) < &HC0 Then
                    Exit Function
                ElseIf (bytArray(lngPos) >= &HC0) And (bytArray(lngPos) <= &HFD) Then
                    If (bytArray(lngPos) And &HFC) = &HFC Then
                        lngBytes = 5
                    ElseIf (bytArray(lngPos) And &HF8) = &HF8 Then
                        lngBytes = 4
                    ElseIf (bytArray(lngPos) And &HF0) = &HF0 Then
                        lngBytes = 3
                    ElseIf (bytArray(lngPos) And &HE0) = &HE0 Then
                        lngBytes = 2
                    ElseIf (bytArray(lngPos) And &HC0) = &HC0 Then
                        lngBytes = 1
                    End If
                    For lngPos = (lngPos + 1) To (lngPos + lngBytes)
                        If Not ((bytArray(lngPos) >= &H80) And (bytArray(lngPos) <= &HBF)) Then
                            Exit Function
                        ElseIf lngPos >= lngReadSize Then
                            Exit Do
                        End If
                    Next lngPos
                    'lngUTF8 = lngUTF8 + 1
                Else
                    lngPos = lngPos + 1
                End If
            Loop
            Private_IsUTF8 = True 'lngUTF8 > 0
        End If
    End If
End Function

Private Function Private_LocalHostName() As String
    Dim lngPos As Long, strBuffer As String * 256
    ' attempt to get the local hostname
    If SocketGetHostName(strBuffer, 256) <> SOCKET_ERR Then
        ' find the null character
        lngPos = InStr(strBuffer, vbNullChar)
        If lngPos Then
            ' get the hostname from the buffer
            Private_LocalHostName = Left$(strBuffer, lngPos - 1)
        Else
            ' return entire buffer
            Private_LocalHostName = strBuffer
        End If
    Else
        Private_SetState sckError, "Private_LocalHostName", WSAGetLastError
    End If
End Function

Private Function Private_LocalIP(Optional Hostname As String) As Long
    Dim lngHostentPtr As Long, lngIPPtr As Long, udtHostent As HOSTENT
    ' see if we got a hostname as input...
    If LenB(Hostname) = 0 Then
        Hostname = Private_LocalHostName
    Else
        ' see if it is actually an IP and not a hostname
        Private_LocalIP = SocketGetIP(Hostname)
        ' if it was an IP then we are done!
        If Private_LocalIP <> INADDR_NONE Then Exit Function
    End If
    ' attempt to get a pointer to hostent structure
    lngHostentPtr = SocketGetHostByName(Hostname)
    If lngHostentPtr Then
        ' get the hostent structure
        RtlMoveMemory udtHostent, ByVal lngHostentPtr, Len(udtHostent)
        ' get pointer to IP
        GetMem4 udtHostent.hAddrList, lngIPPtr
        ' copy IP numbers to byte array from that pointer
        GetMem4 lngIPPtr, Private_LocalIP
    Else
        Private_SetState sckError, "Private_LocalIP", WSAGetLastError
    End If
End Function

Private Function Private_LocalPort() As Long
    Dim udtSockAddr As SOCKADDR_IN
    ' attempt to get socket information
    If SocketGetName(m_Socket, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR Then
        ' return the socket
        Private_LocalPort = (SocketnTons(udtSockAddr.Port) And &HFFFF&)
    Else
        Private_LocalPort = SOCKET_ERR
    End If
End Function

Private Sub Private_Message(ByVal Msg As Long)
    Dim lngBuffer As Long, lngBytesReceived As Long, lngOldBuffer As Long, lngPos As Long, lngResult As Long, udtSockAddr As SOCKADDR_IN
    Dim lngIPPtr As Long, lngRemoteIP As Long, udtHostent As HOSTENT
    Dim strArray() As String, strBuffer As String, strLine As String
    Select Case Msg
        Case FD_ACCEPT
            ' get the new socket
            m_RequestID = SocketAccept(m_Socket, m_RequestSA, Len(m_RequestSA))
            If m_RequestID <> INVALID_SOCKET Then
                ' now let the user choose whether to accept or deny
                RaiseEvent ConnectionRequest(ObjPtr(Me))
            Else
                Private_SetState sckError, "Private_Message", WSAGetLastError
            End If
            ' clean up
            m_RequestSA = udtSockAddr
            m_RequestBuffer = vbNullString
            ' if we still have request ID the connection was not accepted
            If m_RequestID Then
                ' must close
                If SocketClose(m_RequestID) = SOCKET_ERR Then
                    Private_SetState sckError, "Private_Message", WSAGetLastError
                End If
                m_RequestID = 0
            End If
        Case FD_CLOSE
            RaiseEvent Closing
            ' if user did not close the socket we do
            CloseSocket
        Case FD_CONNECT
            If SocketGetPeerName(m_Socket, udtSockAddr, Len(udtSockAddr)) <> SOCKET_ERR Then
                ' get remote port and remote ip
                m_RemotePort = SocketnTons(udtSockAddr.Port) And &HFFFF&
                m_RemoteIP = udtSockAddr.Addr
                ' we are connected!
                Private_SetState sckConnected
                RaiseEvent Connect
            Else
                Private_SetState sckError, "Private_Message", WSAGetLastError
            End If
        Case FD_READ
            ' data has arrived
            If m_Protocol = sckTCPProtocol Then
                ' TCP: get the data to our own buffer
                If (LenB(m_Buffer) - m_BufferSize) < m_ReceiveSize Then
                    ' ensure we always have a buffer that is big enough for any incoming data
                    lngOldBuffer = StrPtr(m_Buffer)
                    lngBuffer = SysAllocStringByteLen(lngOldBuffer, LenB(m_Buffer) + m_ReceiveSize * 2)
                    If lngOldBuffer Then m_Buffer = vbNullString
                    PutMem4 ByVal VarPtr(m_Buffer), ByVal lngBuffer
                Else
                    lngBuffer = StrPtr(m_Buffer)
                End If
                ' make sure we have a buffer...
                If lngBuffer Then
                    ' read to buffer
                    lngBytesReceived = SocketReceive(m_Socket, ByVal lngBuffer + m_BufferSize, m_ReceiveSize, 0)
                    If lngBytesReceived <> SOCKET_ERR Then
                        ' buffer size has increased
                        m_BufferSize = m_BufferSize + lngBytesReceived
                        If m_Mode = [Socket Binary Mode] Then
                            RaiseEvent DataArrival(lngBytesReceived)
                        Else
                            ' fix buffer size
                            'strBuffer = LeftB$(m_Buffer, m_BufferSize)
                            lngBuffer = SysAllocStringByteLen(lngBuffer, m_BufferSize)
                            PutMem4 ByVal VarPtr(strBuffer), ByVal lngBuffer
                            ' then see which line change to split with, if there is any line change at all
                            If InStr(strBuffer, vbNewLine) Then
                                ' UTF-16
                                strArray = Split(strBuffer, vbNewLine)
                                For lngPos = 0 To UBound(strArray) - 1
                                    RaiseEvent TextArrival(strArray(lngPos), [Socket CRLF], False)
                                Next lngPos
                                If LenB(strArray(lngPos)) = 0 Then
                                    m_Buffer = vbNullString
                                    m_BufferSize = 0
                                Else
                                    m_Buffer = strArray(lngPos)
                                    m_BufferSize = LenB(m_Buffer)
                                End If
                            ElseIf InStr(strBuffer, vbLf) Then
                                ' UTF-16
                                strArray = Split(strBuffer, vbLf)
                                For lngPos = 0 To UBound(strArray) - 1
                                    RaiseEvent TextArrival(strArray(lngPos), [Socket LF], False)
                                Next lngPos
                                If LenB(strArray(lngPos)) = 0 Then
                                    m_Buffer = vbNullString
                                    m_BufferSize = 0
                                Else
                                    m_Buffer = strArray(lngPos)
                                    m_BufferSize = LenB(m_Buffer)
                                End If
                            ElseIf InStr(strBuffer, vbCr) Then
                                ' UTF-16
                                strArray = Split(strBuffer, vbNewLine)
                                For lngPos = 0 To UBound(strArray) - 1
                                    RaiseEvent TextArrival(strArray(lngPos), [Socket CR], False)
                                Next lngPos
                                If LenB(strArray(lngPos)) = 0 Then
                                    m_Buffer = vbNullString
                                    m_BufferSize = 0
                                Else
                                    m_Buffer = strArray(lngPos)
                                    m_BufferSize = LenB(m_Buffer)
                                End If
                            ElseIf InStrB(strBuffer, CRLF) Then
                                ' UTF-8 or ANSI
                                lngOldBuffer = 1
                                lngPos = InStrB(strBuffer, CRLF)
                                Do While lngPos
                                    strLine = MidB$(strBuffer, lngOldBuffer, lngPos - lngOldBuffer)
                                    If Private_IsUTF8(strLine) Then
                                        Private_UTF8toUTF16 strLine, strLine
                                        RaiseEvent TextArrival(strLine, [Socket CRLF], False)
                                    Else
                                        RaiseEvent TextArrival(strLine, [Socket CRLF], True)
                                    End If
                                    lngOldBuffer = lngPos + 2
                                    lngPos = InStrB(lngOldBuffer, strBuffer, CRLF)
                                Loop
                                If lngOldBuffer = m_BufferSize Then
                                    m_Buffer = vbNullString
                                    m_BufferSize = 0
                                Else
                                    m_Buffer = MidB$(strBuffer, lngOldBuffer)
                                    m_BufferSize = LenB(m_Buffer)
                                End If
                            ElseIf InStrB(strBuffer, LF) Then
                                ' UTF-8 or ANSI
                                lngOldBuffer = 1
                                lngPos = InStrB(strBuffer, LF)
                                Do While lngPos
                                    strLine = MidB$(strBuffer, lngOldBuffer, lngPos - lngOldBuffer)
                                    If Private_IsUTF8(strLine) Then
                                        Private_UTF8toUTF16 strLine, strLine
                                        RaiseEvent TextArrival(strLine, [Socket LF], False)
                                    Else
                                        RaiseEvent TextArrival(strLine, [Socket LF], True)
                                    End If
                                    lngOldBuffer = lngPos + 1
                                    lngPos = InStrB(lngOldBuffer, strBuffer, LF)
                                Loop
                                If lngOldBuffer = m_BufferSize Then
                                    m_Buffer = vbNullString
                                    m_BufferSize = 0
                                Else
                                    m_Buffer = MidB$(strBuffer, lngOldBuffer)
                                    m_BufferSize = LenB(m_Buffer)
                                End If
                            ElseIf InStrB(strBuffer, CR) Then
                                ' UTF-8 or ANSI
                                lngOldBuffer = 1
                                lngPos = InStrB(strBuffer, CR)
                                Do While lngPos
                                    strLine = MidB$(strBuffer, lngOldBuffer, lngPos - lngOldBuffer)
                                    If Private_IsUTF8(strLine) Then
                                        Private_UTF8toUTF16 strLine, strLine
                                        RaiseEvent TextArrival(strLine, [Socket CR], False)
                                    Else
                                        RaiseEvent TextArrival(strLine, [Socket CR], True)
                                    End If
                                    lngOldBuffer = lngPos + 1
                                    lngPos = InStrB(lngOldBuffer, strBuffer, CR)
                                Loop
                                If lngOldBuffer = m_BufferSize Then
                                    m_Buffer = vbNullString
                                    m_BufferSize = 0
                                Else
                                    m_Buffer = MidB$(strBuffer, lngOldBuffer)
                                    m_BufferSize = LenB(m_Buffer)
                                End If
                            End If
                        End If
                    Else
                        Private_SetState sckError, "Private_Message", WSAGetLastError
                    End If
                Else
                    Private_SetState sckError, "Private_Message"
                End If
            Else
                ' UDP: just tell how many bytes has arrived, user does what s/he wants
                SocketIOCtl m_Socket, FIONREAD, lngBytesReceived
                If lngBytesReceived Then RaiseEvent DataArrival(lngBytesReceived)
                ' we do not store the data further (this clears the buffer)
                SocketReceive m_Socket, lngBuffer, 1, 0
            End If
        Case FD_WRITE
            ' buffer is empty: ready to send more data
            If m_Protocol = sckTCPProtocol Then
                lngBuffer = StrPtr(m_SendBuffer)
                If (lngBuffer <> 0) And (m_SendBufferSize > 0) Then
                    Do While (m_SendBufferSize - m_SendSize) > lngPos
                        lngResult = SocketSend(m_Socket, ByVal lngBuffer + lngPos, m_SendSize, 0)
                        If lngResult = SOCKET_ERR Then
                            lngResult = WSAGetLastError
                            ' WSAEWOULDBLOCK
                            If lngResult = 10035 Then
                                ' decrease buffer
                                m_SendBuffer = MidB$(m_SendBuffer, lngPos + 1)
                                m_SendBufferSize = m_SendBufferSize - lngPos
                                ' report
                                RaiseEvent SendProgress(lngPos, m_SendBufferSize)
                                lngPos = &H7FFFFFFF
                            Else
                                Private_SetState sckError, "Private_Message", lngResult
                            End If
                            Exit Do
                        End If
                        lngPos = lngPos + m_SendSize
                    Loop
                    ' see if still bytes to send
                    If (m_SendBufferSize - lngPos) > 0 Then
                        ' send the remaining data
                        lngResult = SocketSend(m_Socket, ByVal lngBuffer + lngPos, m_SendBufferSize - lngPos, 0)
                        If lngResult = SOCKET_ERR Then
                            lngResult = WSAGetLastError
                            ' WSAEWOULDBLOCK
                            If lngResult = 10035 Then
                                ' decrease buffer
                                m_SendBuffer = MidB$(m_SendBuffer, lngPos + 1)
                                m_SendBufferSize = m_SendBufferSize - lngPos
                                ' report
                                RaiseEvent SendProgress(lngPos, m_BufferSize)
                            Else
                                Private_SetState sckError, "Private_Message", lngResult
                            End If
                        Else
                            ' we have sent everything we got
                            RaiseEvent SendProgress(m_SendBufferSize, 0)
                            m_SendBuffer = vbNullString
                            m_SendBufferSize = 0
                            RaiseEvent SendComplete
                        End If
                    End If
                End If
            Else
                lngBuffer = StrPtr(m_SendBuffer)
                If (lngBuffer <> 0) And (m_SendBufferSize > 0) Then
                    ' UDP: we get IP on each send
                    lngRemoteIP = SocketGetIP(m_RemoteHost)
                    If lngRemoteIP = INADDR_NONE Then
                        ' resolve IP of hostname
                        lngResult = SocketGetHostByName(m_RemoteHost)
                        If lngResult Then
                            ' get the resolved IP
                            RtlMoveMemory udtHostent, ByVal m_HostentPointer, Len(udtHostent)
                            GetMem4 udtHostent.hAddrList, lngIPPtr
                            GetMem4 lngIPPtr, lngRemoteIP
                        Else
                            Private_SetState sckError, "Private_Message", WSAGetLastError
                        End If
                    End If
                    ' now that we have IP we can do something with it
                    With udtSockAddr
                        .Addr = lngRemoteIP
                        .Family = AF_INET
                        .Port = SockethTons(m_RemotePort And &HFFFF)
                    End With
                    ' now we can send it all
                    lngResult = SocketSendTo(m_Socket, ByVal StrPtr(m_SendBuffer), m_SendBufferSize, 0, udtSockAddr, Len(udtSockAddr))
                    If lngResult <> SOCKET_ERR Then
                        ' we simply tell the buffer size is 0
                        m_SendBufferSize = 0
                    Else
                        Private_SetState sckError, "Private_Message", WSAGetLastError
                    End If
                End If
            End If
    End Select
End Sub

Private Sub Private_SetState(ByVal NewStatus As StateConstants, Optional ErrorSource As String, Optional ByVal ErrorNumber As Long, Optional ErrorDescription As String)
    Dim blnCancelDisplay As Boolean
    If NewStatus <> sckError Then m_State = NewStatus
    RaiseEvent StatusChange(NewStatus)
    ' see if we raise the error event
    If ErrorNumber Then
        ' get description if we have not got one
        If LenB(ErrorDescription) = 0 Then ErrorDescription = Private_FormatWSAError(ErrorNumber)
        ' set error source
        If LenB(ErrorSource) = 0 Then ErrorSource = "UniSock" Else ErrorSource = "UniSock." & ErrorSource
        ' raise the event: the developer may choose to show a message box
        blnCancelDisplay = True
        RaiseEvent Error(ErrorNumber And &HFFFF, ErrorDescription, 0, ErrorSource, vbNullString, 0, blnCancelDisplay)
        If Not blnCancelDisplay Then MsgBox ErrorDescription, vbExclamation, ErrorSource
    End If
End Sub

Private Function Private_Start() As Long
    Dim bytValue As Byte, bytSignature As Byte, lngA As Long
    Dim lngAddress As Long, lngMethod As Long, lngObject As Long
    Dim lngCallback As Long, lngMCmem As Long, udtMC As MachineCode

    ' index of ordinal (always 1 in this class, the last procedure)
    Const Ordinal As Long = 1
    ' number of parameters (for TimerProc this is always 4, thus this is a constant)
    Const ParamCount As Long = 4

    ' make sure we do this just once
    If m_Class.lpfnWndProc = 0 Then
        ' get object pointer
        lngObject = ObjPtr(Me)
        ' get VTable address
        GetMem4 lngObject, lngAddress
        ' Class method (see SelfSub code for non-Class values)
        lngAddress = lngAddress + &H1C&
        ' get method pointer
        GetMem4 lngAddress, lngMethod
        ' get method signature byte: &H33 = pseudo-code, &HE9 = native code
        GetMem1 lngMethod, bytSignature
        ' next VTable address
        lngAddress = lngAddress + 4&
        ' try a "reasonable" amount of VTable entries
        For lngA = 511 To 1 Step -1
            ' get method pointer
            GetMem4 lngAddress, lngMethod
            ' see if we are out of VTable (I use "Then Else" because True conditions are faster)
            If IsBadCodePtr(lngMethod) = 0& Then Else Exit For
            ' get method signature byte
            GetMem1 lngMethod, bytValue
            ' if it is invalid we are out of VTable
            If bytValue = bytSignature Then Else Exit For
            ' try next one
            lngAddress = lngAddress + 4&
        Next lngA
        ' if lngA = 0 we looped through the entire loop; if that did not happen, we get the pointer
        If lngA Then GetMem4 lngAddress - (Ordinal * 4&), lngCallback
        ' verify we got the TimerProc callback address of ordinal 1
        If lngCallback Then
            ' allocate executable memory
            lngMCmem = VirtualAlloc(0, LenB(udtMC), &H1000&, &H40&) 'Length, MEM_COMMIT, PAGE_RWX
            ' verify we got it
            If lngMCmem Then
                With udtMC
                    ' thunk
                    .OwnerPtr = lngObject
                    .CallbackAddress = lngCallback
                    If App.LogMode = 0 Then
                        ' for IDE safety, store the EbMode function address in the thunk data
                        .API_EbMode = GetProcAddress(GetModuleHandleA("vba6"), "EbMode")
                    End If
                    .API_IsBadCodePtr = GetProcAddress(GetModuleHandleA("kernel32"), "IsBadCodePtr")
                    '.API_KillTimer = GetProcAddress(GetModuleHandleA("user32"), "KillTimer")
                    ' actual machine code
                    .MC1(5&) = &HBB60E089:    .MC2(7&) = &H73FFC589
                    .MC2(8&) = &HC53FF04:     .MC2(9&) = &H59E80A74
                    .MC2(10) = &HE9000000:    .MC2(11) = &H30&
                    .MC2(12) = &H87B81:       .MC2(13) = &H75000000
                    .MC2(14) = &H9090902B:    .MC2(15) = &H42DE889
                    .MC2(16) = &H50000000:    .MC2(17) = &HB9909090
                    .MC3(19) = &H90900AE3:    .MC3(20) = &H8D74FF
                    .MC3(21) = &H9090FAE2:    .MC3(22) = &H53FF33FF
                    .MC3(23) = &H90909004:    .MC3(24) = &H2BADC261
                    .MC3(25) = &H3D0853FF:    .MC3(26) = &H1&
                    .MC3(27) = &H23DCE74:     .MC3(28) = &H74000000
                    .MC3(29) = &HAE807:       .MC3(30) = &H90900000
                    .MC3(31) = &H4589C031:    .MC3(32) = &H90DDEBFC
                    .MC3(33) = &HFF0C75FF:    .MC3(34) = &H53FF0475
                    .MC3(35) = &HC310&
                    ' settings within the code
                    .AllocatedDataPtr = lngMCmem
                    .Ordinal = Ordinal
                    .ParamCount = ParamCount
                    PutMem2 VarPtr(.MC3(24)) + 2&, CInt(ParamCount * 4&)
                End With
                ' copy thunk code to executable memory
                RtlMachineCodeCopy ByVal lngMCmem, udtMC, LenB(udtMC)
                ' remember the procedure address (add thunk offset)
                Private_Start = lngMCmem + &H14&
            End If
        End If
    End If
End Function

Private Sub Private_Stop()
    ' only do this if we still have the procedure
    If m_Class.lpfnWndProc Then
        ' free the procedure callback
        VirtualFree m_Class.lpfnWndProc, 0&, &H8000& 'MEM_RELEASE
        ' reset procedure pointer to prevent this getting ran twice
        m_Class.lpfnWndProc = 0
    End If
End Sub

Private Sub Private_UTF16toUTF8(ByRef Text As String, ByRef TextOut As String)
    Dim lngBufferLen As Long, lngBufferPtr As Long, lngTextLen As Long, strBuffer As String
    ' first check text length
    lngTextLen = Len(Text)
    If lngTextLen Then
        ' calculate maximum output size
        lngBufferLen = lngTextLen * 4
        ' create the buffer string to hold output data
        lngBufferPtr = SysAllocStringByteLen(0, lngBufferLen)
        PutMem4 ByVal VarPtr(strBuffer), ByVal lngBufferPtr
        ' convert to UTF-8
        lngBufferLen = WideCharToMultiByte(65001, 0&, StrPtr(Text), lngTextLen, lngBufferPtr, lngBufferLen, 0&, ByVal 0&)
        ' output as is best fit
        If lngBufferLen > 0 Then
            TextOut = LeftB$(strBuffer, lngBufferLen)
        Else
            TextOut = vbNullString
        End If
        ' clean buffer
        strBuffer = vbNullString
    Else
        TextOut = vbNullString
    End If
End Sub

Private Sub Private_UTF8toUTF16(ByRef Text As String, ByRef TextOut As String)
    Dim lngBufferLen As Long, lngBufferPtr As Long, lngTextLen As Long, strBuffer As String
    ' first check text length
    lngTextLen = LenB(Text)
    If lngTextLen Then
        ' calculate maximum output size
        lngBufferLen = lngTextLen * 2
        ' create the buffer string to hold output data
        lngBufferPtr = SysAllocStringByteLen(0, lngBufferLen)
        PutMem4 ByVal VarPtr(strBuffer), ByVal lngBufferPtr
        ' convert from UTF-8
        lngBufferLen = MultiByteToWideChar(65001, 0&, StrPtr(Text), lngTextLen, lngBufferPtr, lngBufferLen)
        ' output as is best fit
        If lngBufferLen > 0 Then
            TextOut = Left$(strBuffer, lngBufferLen)
        Else
            TextOut = vbNullString
        End If
        ' clean buffer
        strBuffer = vbNullString
    Else
        TextOut = vbNullString
    End If
End Sub

Public Property Get Protocol() As ProtocolConstants
    Protocol = m_Protocol
End Property

Public Property Let Protocol(ByVal NewValue As ProtocolConstants)
    If m_State = sckClosed Then
        m_Protocol = NewValue
    Else
        Err.Raise sckInvalidOp, "UniSock.Protocol", "Invalid operation at current state."
    End If
End Property

Public Property Get RemoteHost() As String
    RemoteHost = m_RemoteHost
End Property

Public Property Let RemoteHost(ByRef NewValue As String)
    If Not (m_Protocol = sckTCPProtocol And m_State <> sckClosed) Then
        m_RemoteHost = NewValue
    Else
        Err.Raise sckInvalidOp, "UniSock.RemotePort", Private_FormatWSAError(sckInvalidOp)
    End If
End Property

Public Function RemoteHostIP() As String
    ' do some math to build IP string
    RemoteHostIP = CStr(m_RemoteIP And &HFF&) & "." & _
        CStr((m_RemoteIP And &HFF00&) \ &H100&) & "." & _
        CStr((m_RemoteIP And &HFF0000) \ &H10000) & "." & _
        CStr(((m_RemoteIP And &HFF000000) \ &H1000000) And &HFF&)
End Function

Public Function RemoteIP() As Long
    RemoteIP = m_RemoteIP
End Function

Public Property Get RemotePort() As Long
    RemotePort = m_RemotePort
End Property

Public Property Let RemotePort(ByVal NewValue As Long)
    If Not (m_Protocol = sckTCPProtocol And m_State <> sckClosed) Then
        If NewValue >= 0 And NewValue < 65536 Then
            m_RemotePort = NewValue
        Else
            Err.Raise sckInvalidArg, "UniSock.RemotePort", Private_FormatWSAError(sckInvalidArg)
        End If
    Else
        Err.Raise sckInvalidOp, "UniSock.RemotePort", Private_FormatWSAError(sckInvalidOp)
    End If
End Property

Public Function RequestHost() As String
    Dim lngPos As Long, strHost As String * 256
    If m_RequestID Then
        If SocketGetNameInfo(m_RequestSA, Len(m_RequestSA), strHost, 256, vbNullString, 0, 0) <> SOCKET_ERR Then
            lngPos = InStr(strHost, vbNullChar)
            If lngPos Then
                RequestHost = Left$(strHost, lngPos - 1)
            Else
                RequestHost = strHost
            End If
        Else
            Private_SetState sckError, "RequestHost", WSAGetLastError
        End If
    End If
End Function

Public Function RequestIP() As String
    If m_RequestID Then
        RequestIP = CStr(m_RequestSA.Addr And &HFF&) & "." & _
            CStr((m_RequestSA.Addr And &HFF00&) \ &H100&) & "." & _
            CStr((m_RequestSA.Addr And &HFF0000) \ &H10000) & "." & _
            CStr(((m_RequestSA.Addr And &HFF000000) \ &H1000000) And &HFF&)
    End If
End Function

Public Function RequestPort() As Long
    If m_RequestID Then
        RequestPort = SocketnTons(m_RequestSA.Port) And &HFFFF&
    End If
End Function

Public Function SendData(Data) As Boolean
    Dim enmType As VbVarType, lngLen As Long, lngPtr As Long, strData As String
    ' validate connection state
    If m_Protocol = sckTCPProtocol Then
        If m_State <> sckConnected Then
            Err.Raise sckBadState, "UniSock.SendData", Private_FormatWSAError(sckBadState)
            Exit Function
        End If
    Else
        If m_State <> sckOpen Then
            If Not Connect Then Exit Function
        End If
    End If
    ' now determinte the type of the data
    enmType = VarType(Data)
    If enmType And vbArray Then
        Select Case (enmType And Not vbArray)
            Case vbByte
                strData = CStr(Data)
            Case vbCurrency, vbDate, vbDouble
                lngLen = (UBound(Data) - LBound(Data) + 1) * 8&
                lngPtr = SysAllocStringByteLen(0, lngLen)
                PutMem4 ByVal VarPtr(strData), ByVal lngPtr
                RtlMoveMemory ByVal lngPtr, Data(LBound(Data)), lngLen
            Case vbBoolean, vbInteger
                lngLen = (UBound(Data) - LBound(Data) + 1) * 2&
                lngPtr = SysAllocStringByteLen(0, lngLen)
                PutMem4 ByVal VarPtr(strData), ByVal lngPtr
                RtlMoveMemory ByVal lngPtr, Data(LBound(Data)), lngLen
            Case vbLong, vbSingle
                lngLen = (UBound(Data) - LBound(Data) + 1) * 4&
                lngPtr = SysAllocStringByteLen(0, lngLen)
                PutMem4 ByVal VarPtr(strData), ByVal lngPtr
                RtlMoveMemory ByVal lngPtr, Data(LBound(Data)), lngLen
            Case vbString
                strData = StrConv(Join(Data, m_LineChange), vbFromUnicode)
            Case Else
                Err.Raise sckUnsupported, "SendData", Private_FormatWSAError(sckUnsupported)
        End Select
    Else
        Select Case enmType
            Case vbBoolean
                If Data Then strData = ChrW$(&HFFFF&) Else strData = vbNullChar
            Case vbByte
                strData = ChrB$(Data)
            Case vbCurrency
                strData = "    "
                PutMem8 ByVal StrPtr(strData), ByVal CCur(Data)
            Case vbDate
                strData = "    "
                PutMem8 ByVal StrPtr(strData), ByVal CDate(Data)
            Case vbDouble
                strData = "    "
                PutMem8 ByVal StrPtr(strData), ByVal CDbl(Data)
            Case vbInteger
                strData = vbNullChar
                PutMem2 ByVal StrPtr(strData), ByVal CInt(Data)
            Case vbLong
                strData = "  "
                PutMem4 ByVal StrPtr(strData), ByVal CLng(Data)
            Case vbSingle
                strData = "  "
                PutMem4 ByVal StrPtr(strData), ByVal CSng(Data)
            Case vbString
                strData = StrConv(Data, vbFromUnicode)
            Case Else
                Err.Raise sckUnsupported, "SendData", Private_FormatWSAError(sckUnsupported)
        End Select
    End If
    ' see if we have anything to send
    lngLen = LenB(strData)
    If lngLen Then
        ' see if we already have unsent data
        If m_SendBufferSize > 0 Then
            ' see if we only have to add into the buffer
            If lngLen > (LenB(m_SendBuffer) - m_SendBufferSize) Then
                ' we have to increase the buffer
                lngPtr = SysAllocStringByteLen(StrPtr(m_SendBuffer), (LenB(m_SendBuffer) + lngLen) * 2)
                m_SendBuffer = vbNullString
                PutMem4 ByVal VarPtr(m_SendBuffer), ByVal lngPtr
            End If
            ' add it in
            MidB$(m_SendBuffer, m_SendBufferSize + 1, lngLen) = strData
            m_SendBufferSize = m_SendBufferSize + lngLen
        Else
            ' simply set the output buffer size to the size of the data
            m_SendBuffer = strData
            m_SendBufferSize = lngLen
        End If
        ' send buffer
        Private_Message FD_WRITE
        ' success!
        SendData = True
    End If
End Function

' note: line changes are added automatically!
Public Function SendText(ByVal Text As String, Optional ByVal TextFormat As UniSockTextFormat = [Socket UTF-8]) As Boolean
    Dim lngLen As Long, lngPtr As Long
    ' see what we do...
    If TextFormat = [Socket UTF-8] Then
        ' simple enough, convert BSTR to UTF-8
        Private_UTF16toUTF8 Text & m_LineChange, Text
    ElseIf TextFormat = [Socket ANSI] Then
        Text = Text & StrConv(m_LineChange, vbFromUnicode)
    Else
        Text = Text & m_LineChange
    End If
    ' see if we have anything to send
    lngLen = LenB(Text)
    If lngLen Then
        ' see if we already have unsent data
        If m_SendBufferSize > 0 Then
            ' see if we only have to add into the buffer
            If lngLen > (LenB(m_SendBuffer) - m_SendBufferSize) Then
                ' we have to increase the buffer
                lngPtr = SysAllocStringByteLen(StrPtr(m_SendBuffer), (LenB(m_SendBuffer) + lngLen) * 2)
                m_SendBuffer = vbNullString
                PutMem4 ByVal VarPtr(m_SendBuffer), ByVal lngPtr
            End If
            ' add it in
            MidB$(m_SendBuffer, m_SendBufferSize + 1, lngLen) = Text
            m_SendBufferSize = m_SendBufferSize + lngLen
        Else
            ' simply set the output buffer size to the size of the data
            m_SendBuffer = Text
            m_SendBufferSize = lngLen
        End If
        ' send buffer
        Private_Message FD_WRITE
        ' success!
        SendText = True
    End If
End Function

Public Function SocketHandle() As Long
    SocketHandle = m_Socket
End Function

Public Function State() As StateConstants
    State = m_State
End Function

Private Sub Class_Initialize()
    ' byte mode characters
    CRLF = ChrW$(&HA0D)
    CR = ChrB$(&HD)
    LF = ChrB$(&HA)
    NULLCHAR = ChrB$(&H0)
    ' first of all we must create a window that receives all the messages
    m_ClassName = "UniSock" & RightB$("0000000" & Hex$(ObjPtr(Me)), 16)
    With m_Class
        .lpszClassName = StrPtr(m_ClassName)
        .hInstance = App.hInstance
        .lpfnWndProc = Private_Start
    End With
    ' create a custom class for this purpose
    m_RegisterClass = RegisterClassW(m_Class)
    If m_RegisterClass Then
        ' now create the actual window
        m_hWnd = CreateWindowExW(0&, StrPtr(m_ClassName), StrPtr(CStr(ObjPtr(Me))), 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)
        ' ensure it was created
        If m_hWnd Then
            ' now attempt initializing Windows Sockets
            m_Initialized = (WSAStartup(&H202, m_WSAData) = 0&)
        End If
    End If
    ' initial socket is always invalid
    m_Socket = INVALID_SOCKET
    ' use HTTP port as default
    m_RemotePort = 80
    ' default line change is CRLF
    m_LineChange = vbNewLine
    ' request hWnd tracking
    Set m_Request = New Collection
End Sub

Private Sub Class_Terminate()
    ' some serious cleanup business
    CloseSocket
    If m_Initialized Then WSACleanup
    If m_hWnd Then DestroyWindow m_hWnd
    If m_RegisterClass Then UnregisterClassW m_RegisterClass, App.hInstance
    Private_Stop
End Sub

' ordinal #1, last procedure!
Private Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lngIPPtr As Long, lngRemoteIP As Long, udtHostent As HOSTENT
    Dim lngBytes As Long, lngError As Long, lngPtr As Long, strBuffer As String
    If uMsg = RESOLVE_MESSAGE Then
        ' get the possible error code
        lngError = ((lParam And &HFFFF0000) \ &H10000) And &HFFFF&
        ' we have resolved the host?
        If m_AsyncHost Then m_AsyncHost = 0
        ' see if there were any errors
        If lngError = 0 Then
            ' host is resolved!
            Private_SetState sckHostResolved
            ' get the resolved IP
            RtlMoveMemory udtHostent, ByVal m_HostentPointer, Len(udtHostent)
            GetMem4 udtHostent.hAddrList, lngIPPtr
            GetMem4 lngIPPtr, lngRemoteIP
            ' free memory
            If m_HostentHandle Then
                GlobalFree m_HostentHandle
                m_HostentHandle = 0
                m_HostentPointer = 0
            End If
            ' we are ready to connect!
            ConnectIP lngRemoteIP
        Else
            ' free memory
            If m_HostentHandle Then
                GlobalFree m_HostentHandle
                m_HostentHandle = 0
                m_HostentPointer = 0
            End If
            ' darn
            Private_SetState sckError, "Connect", lngError
        End If
    Else
        If uMsg = SOCKET_MESSAGE Then
            ' get the possible error code
            lngError = ((lParam And &HFFFF0000) \ &H10000) And &HFFFF&
            ' now see which socket we are dealing with...
            If wParam = m_Socket Then
                ' own message queue
                If lngError = 0 Then
                    Private_Message lParam And &HFFFF&
                Else
                    Private_SetState sckError, "WndProc", lngError
                End If
            ElseIf wParam = m_RequestID Then
                ' local request processing
                If lngError = 0 Then
                    Select Case lParam And &HFFFF&
                        Case FD_READ
                            lngPtr = SysAllocStringByteLen(0, m_ReceiveSize)
                            PutMem4 ByVal VarPtr(strBuffer), ByVal lngPtr
                            lngBytes = SocketReceive(wParam, lngPtr, m_ReceiveSize, 0)
                            If lngBytes <> SOCKET_ERR Then
                                m_RequestBuffer = m_RequestBuffer & LeftB$(strBuffer, lngBytes)
                            Else
                                Private_SetState sckClosed, "WndProc", WSAGetLastError
                            End If
                    End Select
                Else
                    Private_SetState sckError, "WndProc", lngError
                End If
            ElseIf m_Socket <> INVALID_SOCKET Then
                ' redirect to new location
                PostMessageW m_Request(CStr(wParam)), uMsg, wParam, lParam
            End If
        Else
            ' we must do this so that CreateWindowExW works
            WndProc = DefWindowProcW(hWnd, uMsg, wParam, lParam)
        End If
    End If
End Function
