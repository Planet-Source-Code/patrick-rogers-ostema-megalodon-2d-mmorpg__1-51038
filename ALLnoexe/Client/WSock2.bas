Attribute VB_Name = "WSock2"
Option Explicit
Private Const FD_READ = &H1
Private Const FD_WRITE = &H2
Private Const FD_OOB = &H4
Private Const FD_ACCEPT = &H8
Private Const FD_CONNECT = &H10
Private Const FD_CLOSE = &H20
Private Const SOCKET_ERROR = -1
Private Const INADDR_NONE = &HFFFF
Private mSocket As Long
Public Const WSABASEERR = 10000
Private Declare Sub MemCopy Lib "Kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (xDest As Any, xSource As Any, ByVal nbytes As Long)
Private Const WSADESCRIPTION_LEN = 257
Private Const WSASYS_STATUS_LEN = 129
Private Type WSAData
    wVersion       As Integer
    wHighVersion   As Integer
    szDescription  As String * WSADESCRIPTION_LEN
    szSystemStatus As String * WSASYS_STATUS_LEN
    iMaxSockets    As Integer
    iMaxUdpDg      As Integer
    lpVendorInfo   As Long
End Type

Private Type HOSTENT
    hName     As Long
    hAliases  As Long
    hAddrType As Integer
    hLength   As Integer
    hAddrList As Long
End Type
Private Type sockaddr_in
    sin_family       As Integer
    sin_port         As Integer
    sin_addr         As Long
    sin_zero(1 To 8) As Byte
End Type
Private Declare Function WSAStartup _
    Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Function gethostbyaddr _
    Lib "ws2_32.dll" (addr As Long, ByVal addr_len As Long, _
                      ByVal addr_type As Long) As Long
Private Declare Function gethostbyname _
    Lib "ws2_32.dll" (ByVal host_name As String) As Long
Private Declare Function gethostname _
    Lib "ws2_32.dll" (ByVal host_name As String, _
                      ByVal namelen As Long) As Long
Private Declare Function getservbyname _
    Lib "ws2_32.dll" (ByVal serv_name As String, _
                      ByVal proto As String) As Long
Private Declare Function getprotobynumber _
    Lib "ws2_32.dll" (ByVal proto As Long) As Long
Private Declare Function getprotobyname _
    Lib "ws2_32.dll" (ByVal proto_name As String) As Long
Private Declare Function getservbyport _
    Lib "ws2_32.dll" (ByVal Port As Integer, ByVal proto As Long) As Long
Private Declare Function inet_addr _
    Lib "ws2_32.dll" (ByVal cp As String) As Long
Private Declare Function inet_ntoa _
    Lib "ws2_32.dll" (ByVal inn As Long) As Long
Private Declare Function htons _
    Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Private Declare Function htonl _
    Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Private Declare Function ntohl _
    Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Private Declare Function ntohs _
    Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer
Private Declare Function Socket _
    Lib "ws2_32.dll" Alias "socket" (ByVal af As _
    Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Private Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Private Declare Function Connect _
Lib "ws2_32.dll" Alias "connect" (ByVal s As _
                  Long, ByRef Name As _
                  sockaddr_in, ByVal namelen As Long) As Long
Private Declare Function bind _
Lib "ws2_32.dll" (ByVal s As Long, _
                  ByRef Name As sockaddr_in, _
                  ByRef namelen As Long) As Long
Private Declare Function recv _
Lib "ws2_32.dll" (ByVal s As Long, _
                  ByRef buf As Any, _
                  ByVal buflen As Long, _
                  ByVal flags As Long) As Long
Private Declare Function Send _
Lib "ws2_32.dll" Alias "send" (ByVal s As _
                  Long, ByRef buf As _
                  Any, ByVal buflen As _
                  Long, ByVal flags As Long) As Long
Private Declare Sub RtlMoveMemory _
    Lib "Kernel32" (hpvDest As Any, _
                    ByVal hpvSource As Long, _
                    ByVal cbCopy As Long)
Private Declare Function lstrcpy _
    Lib "Kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, _
                                     ByVal lpString2 As Long) As Long
Private Declare Function lstrlen _
    Lib "Kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Enum AddressFamily
    AF_UNSPEC = 0      '/* unspecified */
    AF_UNIX = 1        '/* local to host (pipes, portals) */
    AF_INET = 2        '/* internetwork: UDP, TCP, etc. */
End Enum
Private Enum SocketType
    SOCK_STREAM = 1    ' /* stream socket */
    SOCK_DGRAM = 2     ' /* datagram socket */
    SOCK_RAW = 3       ' /* raw-protocol interface */
    SOCK_RDM = 4       ' /* reliably-delivered message */
    SOCK_SEQPACKET = 5 ' /* sequenced packet stream */
End Enum
Private Const hostent_size = 16
Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767

Private Function UnsignedToLong(Value As Double) As Long
    '
    'The function takes a Double containing a value in the 
    'range of an unsigned Long and returns a Long that you 
    'can pass to an API that requires an unsigned Long
    '
    If Value < 0 Or Value >= OFFSET_4 Then Error 6 ' Overflow
    '
    If Value <= MAXINT_4 Then
        UnsignedToLong = Value
    Else
        UnsignedToLong = Value - OFFSET_4
    End If
    '
End Function

Private Function LongToUnsigned(Value As Long) As Double
    '
    'The function takes an unsigned Long from an API and 
    'converts it to a Double for display or arithmetic purposes
    '
    If Value < 0 Then
        LongToUnsigned = Value + OFFSET_4
    Else
        LongToUnsigned = Value
    End If
    '
End Function

Private Function UnsignedToInteger(Value As Long) As Integer
    '
    'The function takes a Long containing a value in the range 
    'of an unsigned Integer and returns an Integer that you 
    'can pass to an API that requires an unsigned Integer
    '
    If Value < 0 Or Value >= OFFSET_2 Then Error 6 ' Overflow
    '
    If Value <= MAXINT_2 Then
        UnsignedToInteger = Value
    Else
        UnsignedToInteger = Value - OFFSET_2
    End If
    '
End Function

Private Function IntegerToUnsigned(Value As Integer) As Long
    '
    'The function takes an unsigned Integer from and API and 
    'converts it to a Long for display or arithmetic purposes
    '
    If Value < 0 Then
        IntegerToUnsigned = Value + OFFSET_2
    Else
        IntegerToUnsigned = Value
    End If
    '
End Function

Private Function StringFromPointer(ByVal lPointer As Long) As String
    '
    Dim strTemp As String
    Dim lRetVal As Long
    '
    'prepare the strTemp buffer
    strTemp = String$(lstrlen(ByVal lPointer), 0)
    '
    'copy the string into the strTemp buffer
    lRetVal = lstrcpy(ByVal strTemp, ByVal lPointer)
    '
    'return a string
    If lRetVal Then StringFromPointer = strTemp
    '
End Function
Public Sub Read()
    Dim BytesReceived As Long
    Dim buf As String
    Dim buflen As Long
    Dim rc As String
    Dim i As Byte
    Dim TempS As String
    'Allocate a buffer
    buflen = 255
    buf = String$(buflen, Chr$(0))
    
    'Continue reading the data until the buffer is empty
    Do
        
        BytesReceived = recv(mSocket, ByVal buf, buflen, 0)
        
        'Add to the buffer
        If BytesReceived > 0 Then
            rc = rc & Left$(buf, BytesReceived)
        Else
            Exit Do
        End If
     
    Loop
    If rc = "" Then Exit Sub
    'modNPC.shat = rc
    If Left$(rc, 1) = "$" Then
        If InStr(1, rc, "#") > 0 Then
            buf = Right$(rc, Len(rc) - InStr(1, rc, "#"))
            rc = Left$(rc, InStr(1, rc, "#") - 1)
            modNPC.UpdateInfo buf
        End If
ChopS:
        TempS = Left$(rc, InStr(1, rc, Chr(0)) - 1)
        If InStr(1, rc, Chr(0)) < Len(rc) Then
            rc = TempS & ";" & Right$(rc, Len(rc) - InStr(1, rc, Chr(0)) - 1)
        Else
            rc = Right$(TempS, (Len(TempS) - 1))
        End If
        If InStr(1, rc, Chr(0)) > 0 Then GoTo ChopS
        DXEngine.MoveChat rc
    ElseIf Left$(rc, 1) = "#" Then
        rc = Left$(rc, Len(rc) - 1)
        rc = Right$(rc, Len(rc) - 1)
        modNPC.UpdateInfo rc
    Else
        modNPC.shat = "probw" & rc
    End If
End Sub

Public Function SendToServer(Data As String) As Boolean
    Dim buf As String
    Dim BytesSent As Long
    Data = modNPC.MyMapIndex & "," & modNPC.MyIndex & "," & Data
    'Send the data
    BytesSent = Send(mSocket, ByVal Data, Len(Data) + 1, 0)
    'Return the result
    SendToServer = (BytesSent > 0)
End Function

Public Function Connecter(Host As String, Optional Service As String, Optional Port As Long) As Boolean
    Dim s As Long
    Dim SelectOps As Long
    Dim rc As Long
    Dim sck As sockaddr_in
    
    'Create the socket
    mSocket = Socket(AF_INET, SOCK_STREAM, 0)
    If mSocket <> SOCKET_ERROR Then
        'Populate the socket
        With sck
            .sin_family = AF_INET
            .sin_addr = Resolve(Host)
            If Service <> "" Then
                .sin_port = getservbyname(ByVal Service, ByVal "TCP")
            Else
                .sin_port = htons(Port)
            End If
        End With
        
        'Attempt to connect
        rc = Connect(mSocket, sck, Len(sck))
        
        If rc <> SOCKET_ERROR Then
            
                'Go into asynchronous mode and trap
                'the  Read and Close events
                rc = WSAAsyncSelect(mSocket, _
                        FrmMega.WSTrigger.hWnd, _
                        ByVal &H101, _
                        ByVal FD_READ Or FD_CLOSE)
                If rc <> SOCKET_ERROR Then
                    'Fire the Connect event
                    
                    
                    Connecter = True
                Else
                    Connecter = False
                End If
            
                'Fire the Connect event
             
            
                Connecter = True
            
        Else
            Connecter = False
        End If
    Else
        'The connection was unsuccessful
        Connecter = False
    End If
End Function
Private Function Resolve(Host As String) As Long
    Dim phe As Long
    Dim heDestHost As HOSTENT
    Dim addrList As Long
    Dim rc As Long
    
    'Attempt to resolve by IP first
    rc = inet_addr(ByVal Host)

    'If we couldn't resolve by IP,
    'then try by host name
    If rc = SOCKET_ERROR Then
        phe = gethostbyname(ByVal Host)
        If phe <> 0 Then
            CopyMemory heDestHost, ByVal phe, hostent_size
            CopyMemory addrList, ByVal heDestHost.hAddrList, 4
            CopyMemory rc, ByVal addrList, heDestHost.hLength
        Else
            rc = INADDR_NONE
        End If
    End If
    
    'Return the address
    Resolve = rc
End Function
Public Function FireWS() As Boolean
Dim udtWinsockData As WSAData
If WSAStartup(&H202, udtWinsockData) = 0 Then FireWS = True
End Function
Public Sub KillWS()
Call WSACleanup
End Sub



