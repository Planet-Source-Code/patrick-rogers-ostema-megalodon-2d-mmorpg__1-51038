Attribute VB_Name = "WSock2"
Option Explicit
Private Type Clients
    Socket As Long
    IP As String
    Name As String
    Active As Boolean
    Index As Integer
End Type
Private Type MapzS
    Sockers() As Clients
    MaxCon As Integer
End Type
Public MapSock() As MapzS
Private Const FD_READ = &H1
Private Const FD_WRITE = &H2
Private Const FD_OOB = &H4
Private Const FD_ACCEPT = &H8
Private Const FD_CONNECT = &H10
Private Const FD_CLOSE = &H20
Public Const SOCKET_ERROR = -1
Private mSocket As Long
Public Const WSABASEERR = 10000
Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
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
Private Declare Function Listen _
Lib "ws2_32.dll" Alias "listen" (ByVal s As _
                  Long, ByVal backlog As Long) As Long
Private Declare Function Accept _
Lib "ws2_32.dll" Alias "accept" (ByVal s As _
                  Long, ByRef addr As _
                  sockaddr_in, ByRef addrlen As Long) As Long
Private Declare Sub RtlMoveMemory _
    Lib "kernel32" (hpvDest As Any, _
                    ByVal hpvSource As Long, _
                    ByVal cbCopy As Long)
Private Declare Function lstrcpy _
    Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, _
                                     ByVal lpString2 As Long) As Long
Private Declare Function lstrlen _
    Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
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
Public Function LocalHostName() As String
    Dim buf As String
    Dim rc As Long
    
    'Allocate a buffer
    buf = Space$(255)
    
    'Call the API
    rc = gethostname(buf, Len(buf))
    rc = InStr(buf, vbNullChar)
    
    'Return the host name
    If rc > 0 Then
        LocalHostName = Left$(buf, rc - 1)
    Else
        LocalHostName = ""
    End If
End Function

Public Function ServListen(Port As Long) As Boolean
    Dim sck As sockaddr_in
    Dim rc As Long
    Dim i As Integer
    ReDim MapSock(0 To 0)
    MapSock(0).MaxCon = -1
    
    'Get a new socket
    mSocket = Socket(AF_INET, SOCK_STREAM, 0)
    
    If mSocket > 0 Then
        'Prepare the socket
        With sck
            .sin_family = AF_INET
            .sin_port = htons(Port)
            .sin_addr = 0
        End With
        
        'Bind it to the adapter
        rc = bind(mSocket, sck, Len(sck))
        
        If rc = 0 Then
            'Start listening
            rc = Listen(mSocket, 5)
            If rc <> SOCKET_ERROR Then
                rc = WSAAsyncSelect(mSocket, _
                        MegaServer.WSTrigger.hWnd, _
                        ByVal &H100, _
                        FD_ACCEPT)
                ServListen = True
            
            Else
                'Could not listen
                ServListen = False
            End If
        Else
            'Failed to bind
            ServListen = False
        End If
    Else
        'Failed to create socket
        ServListen = False
    End If
End Function
Public Function Disconn(ByRef ZMap As Integer, Index As Integer) As Boolean
closesocket MapSock(ZMap).Sockers(Index).Socket
MapSock(ZMap).Sockers(Index).Active = False
MapSock(ZMap).Sockers(Index).IP = ""
MapSock(ZMap).Sockers(Index).Socket = 0
MapSock(ZMap).Sockers(Index).Name = ""
End Function

Public Function ServAccept() As Boolean
    Dim sck As sockaddr_in
    Dim rc As Long
    Dim i As Integer
    Dim b As Integer
    Dim intX As Integer
    Dim intY As Integer
    Dim TempAddy As String
    'Return the port that the call was accepted on
    rc = Accept(mSocket, sck, Len(sck))

    If rc <> SOCKET_ERROR Then
    If MapSock(0).MaxCon > -1 Then
    For i = 0 To MapSock(0).MaxCon
        If MapSock(0).Sockers(i).Active = False Then
            MapSock(0).Sockers(i).Socket = rc
            MapSock(0).Sockers(i).Active = True
            GoTo Skipp:
        End If
    Next
    End If
    MapSock(0).MaxCon = MapSock(0).MaxCon + 1
    ReDim Preserve MapSock(0).Sockers(0 To MapSock(0).MaxCon)
    MapSock(0).Sockers(MapSock(0).MaxCon).Socket = rc
    MapSock(0).Sockers(MapSock(0).MaxCon).Active = True
    'MegaServer.GameChat.Text = MegaServer.GameChat.Text + vbCrLf + "created new index"
    i = MapSock(0).MaxCon
    ReDim Preserve DataManage.ALLMyNPC(0).MyNPC(0 To i)
Skipp:
     'MegaServer.GameChat.Text = MegaServer.GameChat.Text + vbCrLf + Str$(i)
    DataManage.ALLMyNPC(0).MyNPC(i).X = 24
    DataManage.ALLMyNPC(0).MyNPC(i).Y = 25
    For intX = 0 To 5
        For intY = 0 To 5
            If (ALLMyNPC(0).MyNPC(i).X + intX) < 50 And (ALLMyNPC(0).MyNPC(i).Y + intY) < 50 Then
                If SMapArray(ALLMyNPC(0).MyNPC(i).X + intX, ALLMyNPC(0).MyNPC(i).Y + intY).NPC(0).Index = -1 Then
                    ALLMyNPC(0).MyNPC(i).X = ALLMyNPC(0).MyNPC(i).X + intX
                    ALLMyNPC(0).MyNPC(i).Y = ALLMyNPC(0).MyNPC(i).Y + intY
                    GoTo Skipp2
                End If
            End If
            If (ALLMyNPC(0).MyNPC(i).X - intX) > -1 And (ALLMyNPC(0).MyNPC(i).Y - intY) > -1 Then
                If SMapArray(ALLMyNPC(0).MyNPC(i).X - intX, ALLMyNPC(0).MyNPC(i).Y - intY).NPC(0).Index = -1 Then
                    ALLMyNPC(0).MyNPC(i).X = ALLMyNPC(0).MyNPC(i).X - intX
                    ALLMyNPC(0).MyNPC(i).Y = ALLMyNPC(0).MyNPC(i).Y - intY
                    GoTo Skipp2
                End If
            End If
        Next
    Next
Skipp2:
     For b = 0 To 24
        DataManage.ALLMyNPC(0).MyNPC(i).BPack(b).Index = -1
        DataManage.ALLMyNPC(0).MyNPC(i).BPack(b).Amount = 0
        DataManage.ALLMyNPC(0).MyNPC(i).BPack(b).Equipped = False
     Next
        'Go into asynchronous receive mode
       
            rc = WSAAsyncSelect(MapSock(0).Sockers(i).Socket, _
                    MegaServer.WSTrigger.hWnd, _
                     ByVal &H101, _
                    ByVal FD_READ Or FD_CLOSE)
    TempAddy = GetAscIP(sck.sin_addr)
    MapSock(0).Sockers(i).IP = TempAddy
    MegaServer.GameChat.Text = MegaServer.GameChat.Text + vbCrLf + "User Connected on IP: " & TempAddy
    
    End If
    'Return the result
    ServAccept = (rc <> SOCKET_ERROR)
End Function
Public Function Read(Socket As Long) As String
    Dim BytesReceived As Long
    Dim buf As String
    Dim buflen As Long
    Dim rc As String
    Dim i As Byte
    'Allocate a buffer
    buflen = 255
    buf = String$(buflen, Chr$(0))
    
    'Continue reading the data until the buffer is empty
    Do
        
        BytesReceived = recv(Socket, ByVal buf, buflen, 0)
        
        'Add to the buffer
        If BytesReceived > 0 Then
            rc = rc & Left$(buf, BytesReceived)
        Else
            Exit Do
        End If
     
    Loop
    
    'Return the buffer
    Read = rc
End Function

Public Function SendChatToClients(ZMap As Integer, Data As String) As Boolean
    Dim buf As String
    Dim BytesSent As Long
    Dim i As Byte
    For i = 0 To MapSock(ZMap).MaxCon
    'Send the data
        If MapSock(ZMap).Sockers(i).Active Then
            BytesSent = Send(MapSock(ZMap).Sockers(i).Socket, ByVal Data, Len(Data) + 1, 0)
        End If
    Next
    'Return the result
    SendChatToClients = (BytesSent > 0)
End Function
Public Sub Get411()
    Dim i As Integer
    Dim lngPtrToHOSTENT As Long
    Dim udtHostent      As HOSTENT
    Dim lngPtrToIP      As Long
    Dim arrIpAddress()  As Byte
    Dim strIpAddress    As String
    Dim strHostName As String * 256
    Dim lngRetVal As Long
    MegaServer.NFOList.Clear
    MegaServer.NFOList.AddItem "Server Name: " & (LocalHostName)
    lngPtrToHOSTENT = gethostbyname(Trim$((Left$(strHostName, InStr(1, strHostName, Chr(0)) - 1))))
If lngPtrToHOSTENT = 0 Then
        MsgBox "Bad monkies"
    Else
        RtlMoveMemory udtHostent, lngPtrToHOSTENT, LenB(udtHostent)
        RtlMoveMemory lngPtrToIP, udtHostent.hAddrList, 4
        Do Until lngPtrToIP = 0
            ReDim arrIpAddress(1 To udtHostent.hLength)
            RtlMoveMemory arrIpAddress(1), lngPtrToIP, udtHostent.hLength
            For i = 1 To udtHostent.hLength
                strIpAddress = strIpAddress & arrIpAddress(i) & "."
            Next
            strIpAddress = Left$(strIpAddress, Len(strIpAddress) - 1)
            MegaServer.NFOList.AddItem "IP: " & strIpAddress
            strIpAddress = ""
            udtHostent.hAddrList = udtHostent.hAddrList + LenB(udtHostent.hAddrList)
            RtlMoveMemory lngPtrToIP, udtHostent.hAddrList, 4
         Loop
End If
End Sub
Private Function GetAscIP(ByVal inn As Long) As String
    Dim nStr As Long
    Dim lpStr As Long
    Dim retString As String
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr Then
        nStr = lstrlen(lpStr)
        If nStr > 32 Then nStr = 32
        MemCopy ByVal retString, ByVal lpStr, nStr
        retString = Left$(retString, nStr)
        GetAscIP = retString
    Else
        GetAscIP = "255.255.255.255"
    End If
End Function
Public Function GetIndexOnDisc(ByRef ZMap, Socket As Long) As Integer
Dim j As Integer
For ZMap = 0 To UBound(SMapArray(0, 0).TileProp, 1)
    For j = 0 To MapSock(ZMap).MaxCon
        'MegaServer.GameChat.Text = MegaServer.GameChat.Text + vbCrLf + Str$(ZMap) + ";" + Str$(j)
        If MapSock(ZMap).Sockers(j).Socket = Socket And MapSock(ZMap).Sockers(j).Active = True Then
            GetIndexOnDisc = j
            ZMap = ZMap
            'MegaServer.GameChat.Text = MegaServer.GameChat.Text + vbCrLf + "found on map: " + Str$(ZMap) + " Index:" + Str$(j)
            Exit Function
        End If
    Next
Next
ZMap = 0
GetIndexOnDisc = 0
End Function
Public Function CheckZIndex(ZMap As Integer, ZIndex As Integer, Socket As Long) As Boolean
If ZMap = -1 Then ZMap = ZMap + 1
If ZIndex <= UBound(MapSock(ZMap).Sockers, 1) Then
    If MapSock(ZMap).Sockers(ZIndex).Socket = Socket And MapSock(ZMap).Sockers(ZIndex).Active Then CheckZIndex = True
End If
End Function
Public Function FireWS() As Boolean
Dim udtWinsockData As WSAData
If WSAStartup(&H202, udtWinsockData) = 0 Then FireWS = True
End Function
Public Sub KillWS()
Call WSACleanup
End Sub
Public Sub DiscClient(ByRef ZMap As Integer, Index As Integer, Socket As Long)
Dim i As Integer
    If ZMap > DataManage.MapTotal Then Exit Sub
    MegaServer.GameChat.Text = MegaServer.GameChat.Text + vbCrLf + (ALLMyNPC(ZMap).MyNPC(Index).Namer & " on " & MapSock(ZMap).Sockers(Index).IP & " left at " & ALLMyNPC(ZMap).MyNPC(Index).X & ", " & ALLMyNPC(ZMap).MyNPC(Index).Y) & "on map" & ZMap
    SMapArray(ALLMyNPC(ZMap).MyNPC(Index).X, ALLMyNPC(ZMap).MyNPC(Index).Y).NPC(ZMap).Index = -1
    ALLMyNPC(ZMap).MyNPC(Index).Body = 6
    ALLMyNPC(ZMap).MyNPC(Index).Active = False
    Disconn ZMap, Index
    SendDataToClients ZMap, Index, "#N(" & Index + DataManage.ALLServNPC(ZMap).NPCTotal & ")K"
    MegaServer.GameChat.Text = MegaServer.GameChat.Text + " with health:" + Str(ALLMyNPC(ZMap).MyNPC(Index).Attribs.HP)
End Sub
Public Function SendDataToClients(ZMap As Integer, SkipIndex As Integer, Data As String) As Boolean
    Dim buf As String
    Dim BytesSent As Long
    Dim i As Byte
    For i = 0 To MapSock(ZMap).MaxCon
    
    'Send the data
        If MapSock(ZMap).Sockers(i).Active Then
            If i <> SkipIndex Then
                'MegaServer.GameChat.Text = MegaServer.GameChat.Text + vbCrLf + "out:" & Data
                BytesSent = Send(MapSock(ZMap).Sockers(i).Socket, ByVal Data, Len(Data) + 1, 0)
            End If
        End If
    Next
    'Return the result
    SendDataToClients = (BytesSent > 0)
End Function
Public Function SendDataToAClient(ZMap As Integer, ClientIndex As Integer, Data As String) As Boolean
    Dim buf As String
    Dim BytesSent As Long
    'Send the data
                'MegaServer.GameChat.Text = MegaServer.GameChat.Text + vbCrLf + "out:" & Data
                BytesSent = Send(MapSock(ZMap).Sockers(ClientIndex).Socket, ByVal Data, Len(Data) + 1, 0)
    'Return the result
    SendDataToAClient = (BytesSent > 0)
End Function
