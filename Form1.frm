VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Pure Winsock API fun"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "TCP table"
      Height          =   3255
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   6375
      Begin MSComctlLib.ListView ListView1 
         Height          =   2415
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4260
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Get list"
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   2760
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Other stuff"
      Height          =   2775
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton Command3 
         Caption         =   "Ping a server"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Get your IP"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Get website source"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox Text2 
         Height          =   1935
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   720
         Width           =   3975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "GO"
         Height          =   255
         Left            =   3480
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Text            =   "www.microsoft.com"
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Website:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Want more API? Email me at rconard@mikroware.com!"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   6480
      Width           =   4215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------
'BEGIN: Stuff for PING
'----------------------------------
Const SOCKET_ERROR = 0
Private Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
Private Type Hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Private Type IP_OPTION_INFORMATION
    TTL As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Long
    OptionsData As String * 128
End Type
Private Type IP_ECHO_REPLY
    Address(0 To 3) As Byte
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    data As Long
    Options As IP_OPTION_INFORMATION
End Type
Private Declare Function GetHostByName Lib "wsock32.dll" Alias "gethostbyname" (ByVal HostName As String) As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired&, lpWSAdata As WSAdata) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean
Private Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean
'----------------------------------
'END: Stuff for PING
'----------------------------------

'----------------------------------
'BEGIN: Stuff for TCP Table
'----------------------------------
Private Type MIB_TCPROW
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
End Type

Private Const ERROR_SUCCESS            As Long = 0
Private Const MIB_TCP_STATE_CLOSED     As Long = 1
Private Const MIB_TCP_STATE_LISTEN     As Long = 2
Private Const MIB_TCP_STATE_SYN_SENT   As Long = 3
Private Const MIB_TCP_STATE_SYN_RCVD   As Long = 4
Private Const MIB_TCP_STATE_ESTAB      As Long = 5
Private Const MIB_TCP_STATE_FIN_WAIT1  As Long = 6
Private Const MIB_TCP_STATE_FIN_WAIT2  As Long = 7
Private Const MIB_TCP_STATE_CLOSE_WAIT As Long = 8
Private Const MIB_TCP_STATE_CLOSING    As Long = 9
Private Const MIB_TCP_STATE_LAST_ACK   As Long = 10
Private Const MIB_TCP_STATE_TIME_WAIT  As Long = 11
Private Const MIB_TCP_STATE_DELETE_TCB As Long = 12

Private Declare Function GetTcpTable Lib "iphlpapi.dll" (ByRef pTcpTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Private Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long
Private Declare Function inet_ntoa Lib "wsock32.dll" (ByVal addr As Long) As Long
Private Declare Function ntohs Lib "wsock32.dll" (ByVal addr As Long) As Long
   
Public Function GetInetAddrStr(Address As Long) As String
    GetInetAddrStr = GetString(inet_ntoa(Address))
End Function

Public Function GetString(ByVal lpszA As Long) As String
    GetString = String$(lstrlenA(ByVal lpszA), 0)
    Call lstrcpyA(ByVal GetString, ByVal lpszA)
End Function
'----------------------------------
'END: Stuff for TCP Table
'----------------------------------

Private Sub Command1_Click()
    'Code for the website source function
    lSocket = ConnectSock(Text1.Text, 80, 0, Me.hwnd, False)
End Sub

Private Sub Command2_Click()
    'Code for the IP get function
    MsgBox "IP-address: " + GetIPAddress
End Sub

Private Sub Command3_Click()
    'Code for the ping
    Dim hFile As Long, lpWSAdata As WSAdata
    Dim hHostent As Hostent, AddrList As Long
    Dim Address As Long, rIP As String
    Dim OptInfo As IP_OPTION_INFORMATION
    Dim EchoReply As IP_ECHO_REPLY
    Call WSAStartup(&H101, lpWSAdata)
    Dim Server As String
    Server = InputBox("What server?", "Sever", "www.microsoft.com")
    
    If GetHostByName(Server + String(64 - Len(Server), 0)) <> SOCKET_ERROR Then
        CopyMemory hHostent.h_name, ByVal GetHostByName(Server + String(64 - Len(Server), 0)), Len(hHostent)
        CopyMemory AddrList, ByVal hHostent.h_addr_list, 4
        CopyMemory Address, ByVal AddrList, 4
    End If
    hFile = IcmpCreateFile()
    If hFile = 0 Then
        MsgBox "Unable to Create File Handle"
        Exit Sub
    End If
    OptInfo.TTL = 255
    If IcmpSendEcho(hFile, Address, String(32, "A"), 32, OptInfo, EchoReply, Len(EchoReply) + 8, 2000) Then
        rIP = CStr(EchoReply.Address(0)) + "." + CStr(EchoReply.Address(1)) + "." + CStr(EchoReply.Address(2)) + "." + CStr(EchoReply.Address(3))
    Else
        MsgBox "Timeout"
    End If
    If EchoReply.Status = 0 Then
        MsgBox "Reply from " + Server + " (" + rIP + ") recieved after " + Trim$(CStr(EchoReply.RoundTripTime)) + "ms"
    Else
        MsgBox "Failure ..."
    End If
    Call IcmpCloseHandle(hFile)
    Call WSACleanup
End Sub

Private Sub Command4_Click()
'More stuff for TCP table
Dim TcpRow As MIB_TCPROW
Dim buff() As Byte
Dim lngRequired As Long
Dim lngStrucSize As Long
Dim lngRows As Long
Dim lngCnt As Long
Dim strTmp As String
Dim lstLine As ListItem

Call GetTcpTable(ByVal 0&, lngRequired, 1)

If lngRequired > 0 Then
    ReDim buff(0 To lngRequired - 1) As Byte
    If GetTcpTable(buff(0), lngRequired, 1) = ERROR_SUCCESS Then
        lngStrucSize = LenB(TcpRow)
        'first 4 bytes indicate the number of entries
        CopyMemory lngRows, buff(0), 4
        
        For lngCnt = 1 To lngRows
            'moves past the four bytes obtained above
            'to get data and cast into a TcpRow stucture
            CopyMemory TcpRow, buff(4 + (lngCnt - 1) * lngStrucSize), lngStrucSize
            'sends results to the listview
            
            With TcpRow
                Set lstLine = ListView1.ListItems.Add(, , GetInetAddrStr(.dwLocalAddr))
                lstLine.SubItems(1) = ntohs(.dwLocalPort)
                lstLine.SubItems(2) = GetInetAddrStr(.dwRemoteAddr)
                lstLine.SubItems(3) = ntohs(.dwRemotePort)
                lstLine.SubItems(4) = (.dwState)
                Select Case .dwState
                    Case MIB_TCP_STATE_CLOSED:       strTmp = "closed"
                    Case MIB_TCP_STATE_LISTEN:       strTmp = "listening"
                    Case MIB_TCP_STATE_SYN_SENT:     strTmp = "sent"
                    Case MIB_TCP_STATE_SYN_RCVD:     strTmp = "received"
                    Case MIB_TCP_STATE_ESTAB:        strTmp = "established"
                    Case MIB_TCP_STATE_FIN_WAIT1:    strTmp = "fin wait 1"
                    Case MIB_TCP_STATE_FIN_WAIT2:    strTmp = "fin wait 1"
                    Case MIB_TCP_STATE_CLOSE_WAIT:   strTmp = "close wait"
                    Case MIB_TCP_STATE_CLOSING:      strTmp = "closing"
                    Case MIB_TCP_STATE_LAST_ACK:     strTmp = "last ack"
                    Case MIB_TCP_STATE_TIME_WAIT:    strTmp = "time wait"
                    Case MIB_TCP_STATE_DELETE_TCB:   strTmp = "TCB deleted"
                End Select
                lstLine.SubItems(4) = lstLine.SubItems(4) & "( " & strTmp & " )"
                strTmp = ""
            End With
        
        Next
    
    End If
End If
End Sub

Private Sub Form_Load()
    'Get source stuff
    Dim sSave As String
    Me.AutoRedraw = True
    Set Obj = Me.Text2
    HookForm Me
    StartWinsock sSave
    
    'TCP table stuff
    With ListView1
    .View = lvwReport
    .ColumnHeaders.Add , , "Local IP Address"
    .ColumnHeaders.Add , , "Local Port"
    .ColumnHeaders.Add , , "Remote IP Address"
    .ColumnHeaders.Add , , "Remote Port"
    .ColumnHeaders.Add , , "Status "
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'More stuff for get source
    closesocket lSocket
    EndWinsock
    UnHookForm Me
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'More stuff for TCP table
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.SortOrder = Abs(Not ListView1.SortOrder = 1)
    ListView1.Sorted = True
End Sub
