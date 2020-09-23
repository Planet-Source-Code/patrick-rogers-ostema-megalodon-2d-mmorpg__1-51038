Attribute VB_Name = "MegaInput"
Option Explicit
Private Type MyKeys
    Pressed As Boolean
    String As String
    Last As Boolean
End Type
Private oDX As New DirectX7
Private oDI As DirectInput
Private oDIDEV As DirectInputDevice
Private oDIState As DIKEYBOARDSTATE
Private ChatString As String
Private aKeys(211) As MyKeys
Public Function GetKeyState(lKey As Long) As Boolean
    GetKeyState = aKeys(lKey).Pressed
End Function
Public Function GetKeyLast(lKey As Long) As Boolean
    GetKeyLast = aKeys(lKey).Last
End Function
Public Sub SetKeyLast(lKey As Long, val As Boolean)
    aKeys(lKey).Last = val
End Sub
Public Function Init() As Boolean
On Error GoTo ErrorHandler
    Set oDI = oDX.DirectInputCreate()
    Set oDIDEV = oDI.CreateDevice("GUID_SysKeyboard")
    oDIDEV.SetCommonDataFormat DIFORMAT_KEYBOARD
    oDIDEV.SetCooperativeLevel FrmMega.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    oDIDEV.Acquire
    LoadKeyStrings
    Init = True
    Exit Function
ErrorHandler:
    Init = False
End Function
Public Sub CheckAllKeys()
    Dim i As Integer
    oDIDEV.GetDeviceStateKeyboard oDIState
    For i = 1 To 211
        If oDIState.Key(i) <> 0 Then
            aKeys(i).Pressed = True

            If DXEngine.Chatting Then
                If aKeys(i).Last = False Then
                aKeys(i).Last = True
                    If aKeys(28).Pressed Then
                    If ChatString <> "" Then
                        If WSock2.SendToServer("$" & ChatString) Then
                            ChatString = ""
                            DXEngine.Chatting = False
                            Exit Sub
                        End If
                    End If
                    End If
                    If aKeys(42).Pressed Or aKeys(54).Pressed Then
                        ChatString = ChatString & UCase(aKeys(i).String)
                    Else
                        ChatString = ChatString & aKeys(i).String
                    End If
                    If aKeys(14).Pressed Then
                        ChatString = Left$(ChatString, Len(ChatString) - 10)
                    End If
                    
                End If
            
            End If
            
        Else
            aKeys(i).Last = False
            aKeys(i).Pressed = False
        End If
    Next
End Sub
Public Sub Done()
    oDIDEV.Unacquire
End Sub
Public Sub ClearChat()
ChatString = ""
End Sub
Private Sub LoadKeyStrings()
    aKeys(1).String = "ESCAPE"
    aKeys(2).String = "!"
    aKeys(3).String = "@"
    aKeys(4).String = "#"
    aKeys(5).String = "$"
    aKeys(6).String = "%"
    aKeys(7).String = "^"
    aKeys(8).String = "&"
    aKeys(9).String = "*"
    aKeys(10).String = "("
    aKeys(11).String = ")"
    aKeys(12).String = "-"
    aKeys(13).String = "="
    aKeys(14).String = "BackSpace"
    aKeys(15).String = "     "
    aKeys(16).String = "q"
    aKeys(17).String = "w"
    aKeys(18).String = "e"
    aKeys(19).String = "r"
    aKeys(20).String = "t"
    aKeys(21).String = "y"
    aKeys(22).String = "u"
    aKeys(23).String = "i"
    aKeys(24).String = "o"
    aKeys(25).String = "p"
    aKeys(26).String = "["
    aKeys(27).String = "]"
    aKeys(28).String = "Enter"
    aKeys(29).String = "LCONTROL"
    aKeys(30).String = "a"
    aKeys(31).String = "s"
    aKeys(32).String = "d"
    aKeys(33).String = "f"
    aKeys(34).String = "g"
    aKeys(35).String = "h"
    aKeys(36).String = "j"
    aKeys(37).String = "k"
    aKeys(38).String = "l"
    aKeys(39).String = ";"
    aKeys(40).String = "'"
    aKeys(41).String = "`"
    aKeys(42).String = ""
    aKeys(43).String = "\"
    aKeys(44).String = "z"
    aKeys(45).String = "x"
    aKeys(46).String = "c"
    aKeys(47).String = "v"
    aKeys(48).String = "b"
    aKeys(49).String = "n"
    aKeys(50).String = "m"
    aKeys(51).String = ","
    aKeys(52).String = "."
    aKeys(53).String = "?"
    aKeys(54).String = ""
    aKeys(55).String = "*"
    aKeys(56).String = "Left ALT"
    aKeys(57).String = " "
    aKeys(58).String = "CAPS LOCK"
    aKeys(59).String = "F1"
    aKeys(60).String = "F2"
    aKeys(61).String = "F3"
    aKeys(62).String = "F4"
    aKeys(63).String = "F5"
    aKeys(64).String = "F6"
    aKeys(65).String = "F7"
    aKeys(66).String = "F8"
    aKeys(67).String = "F9"
    aKeys(68).String = "F10"
    aKeys(69).String = "vNUMLOCK"
    aKeys(70).String = "SCROLL  SCROLL LOCK"
    aKeys(71).String = "7"
    aKeys(72).String = "8"
    aKeys(73).String = "9"
    aKeys(74).String = "-"
    aKeys(75).String = "4"
    aKeys(76).String = "5"
    aKeys(77).String = "6"
    aKeys(78).String = "+"
    aKeys(79).String = "1"
    aKeys(80).String = "2"
    aKeys(81).String = "3"
    aKeys(82).String = "0"
    aKeys(83).String = "."
    aKeys(87).String = "F11"
    aKeys(88).String = "F12"
    aKeys(86).String = "F13"
    aKeys(84).String = "F14"
    aKeys(85).String = "F15"
    aKeys(156).String = "NUMPADENTER"
    aKeys(157).String = "RCONTROL"
    aKeys(91).String = "NUMPADCOMMA Comma on NEC PC98 numeric keypad"
    aKeys(181).String = "/"
    aKeys(183).String = "SYSRQ"
    aKeys(184).String = "Right ALT"
    aKeys(199).String = "HOME"
    aKeys(200).String = "UP  Up arrow"
    aKeys(201).String = "PRIOR  PAGE UP"
    aKeys(203).String = "LEFT  Left arrow"
    aKeys(205).String = "RIGHT  Right arrow"
    aKeys(207).String = "END"
    aKeys(208).String = "DOWN  Down arrow"
    aKeys(209).String = "NEXT  PAGE DOWN"
    aKeys(210).String = "INSERT"
    aKeys(211).String = "DELETE"
    aKeys(116).String = "PAUSE"
End Sub
Public Function GetChatString() As String
GetChatString = ChatString
End Function
