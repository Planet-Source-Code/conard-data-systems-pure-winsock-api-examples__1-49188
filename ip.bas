Attribute VB_Name = "ip"
Public Const MIN_SOCKETS_REQD As Long = 1
Public Const WS_VERSION_REQD As Long = &H101
Public Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Public Const SOCKET_ERROR As Long = -1
Public Const WSADESCRIPTION_LEN = 257
Public Const WSASYS_STATUS_LEN = 129
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type
Type WSADataInfo
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSADESCRIPTION_LEN
    szSystemStatus As String * WSASYS_STATUS_LEN
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As String
End Type
Public Type Hostent
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long
End Type
Declare Function WSAStartupInfo Lib "WSOCK32" Alias "WSAStartup" (ByVal wVersionRequested As Integer, lpWSAdata As WSADataInfo) As Long
Declare Function WSACleanup Lib "WSOCK32" () As Long
Declare Function WSAGetLastError Lib "WSOCK32" () As Long
Declare Function WSAStartup Lib "WSOCK32" (ByVal wVersionRequired As Long, lpWSAdata As WSAdata) As Long
Declare Function gethostname Lib "WSOCK32" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Declare Function GetHostByName Lib "WSOCK32" Alias "gethostbyname" (ByVal szHost As String) As Long
Declare Sub CopyMemoryIP Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Public Function GetIPAddress() As String
    Dim sHostName As String * 256
    Dim lpHost As Long
    Dim HOST As Hostent
    Dim dwIPAddr As Long
    Dim tmpIPAddr() As Byte
    Dim I As Integer
    Dim sIPAddr As String
    If Not SocketsInitialize() Then
        GetIPAddress = ""
        Exit Function
    End If
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPAddress = ""
        MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & " has occurred. Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    sHostName = Trim$(sHostName)
    lpHost = GetHostByName(sHostName)
    If lpHost = 0 Then
        GetIPAddress = ""
        MsgBox "Windows Sockets are not responding. " & "Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    CopyMemoryIP HOST, lpHost, Len(HOST)
    CopyMemoryIP dwIPAddr, HOST.hAddrList, 4
    ReDim tmpIPAddr(1 To HOST.hLen)
    CopyMemoryIP tmpIPAddr(1), dwIPAddr, HOST.hLen
    For I = 1 To HOST.hLen
        sIPAddr = sIPAddr & tmpIPAddr(I) & "."
    Next
    GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    SocketsCleanup
End Function

Public Function GetIPHostName() As String
    Dim sHostName As String * 256
    If Not SocketsInitialize() Then
        GetIPHostName = ""
        Exit Function
    End If
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPHostName = ""
        MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & " has occurred. Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup
End Function

Public Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function

Public Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function

Public Sub SocketsCleanup()
    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
End Sub

Public Function SocketsInitialize() As Boolean
    Dim WSAD As WSAdata
    Dim sLoByte As String
    Dim sHiByte As String
    If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
        MsgBox "The 32-bit Windows Socket is not responding."
        SocketsInitialize = False
        Exit Function
    End If
    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & CStr(MIN_SOCKETS_REQD) & " supported sockets."
        SocketsInitialize = False
        Exit Function
    End If
    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
        sHiByte = CStr(HiByte(WSAD.wVersion))
        sLoByte = CStr(LoByte(WSAD.wVersion))
        MsgBox "Sockets version " & sLoByte & "." & sHiByte & " is not supported by 32-bit Windows Sockets."
        SocketsInitialize = False
        Exit Function
    End If
    'must be OK, so lets do it
    SocketsInitialize = True
End Function

