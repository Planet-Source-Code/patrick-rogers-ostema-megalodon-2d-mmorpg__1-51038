Attribute VB_Name = "modNPC"
Private Type NPC
    X As Integer
    Y As Integer
    MovingX As Integer
    MovingY As Integer
    Walking As Byte
    LastMove As Byte
    Active As Boolean
    Step As Byte
    Head As Byte
    Body As Byte
    Weap As Byte
    WeapStep As Byte
End Type
Public MyNPC() As NPC
Public MyIndex As Integer
Public MyMapIndex As Integer
Public Const North = 1
Public Const South = 2
Public Const East = 3
Public Const West = 4
Public NPCCount As Integer
Public shat As String
Private Connected As Boolean
Public Sub WalkUp()
If MyNPC(MyIndex).Y > 43 Then GoTo Skipp
If MyY = 0 And MovingY = 0 Then GoTo Skipp:
    If MovingY = 0 Then
        MovingY = 32
        MyY = MyY - 1
    End If
MovingY = MovingY - MEGASCROLL
Skipp:
If MyNPC(MyIndex).MovingY = 2 Then
    MyNPC(MyIndex).MovingY = 0
    MyNPC(MyIndex).Walking = 0
    Exit Sub
End If
MyNPC(MyIndex).MovingY = MyNPC(MyIndex).MovingY - MEGASCROLL
End Sub

Public Sub WalkDown()
If MyNPC(MyIndex).Y < 6 Then GoTo Skipp
If MyY = 39 Then GoTo Skipp
If MovingY = 30 Then
    MyY = MyY + 1
    If MyY = 38 Then
        MovingY = 0
        GoTo Skipp
    End If
    MovingY = -2
End If
MovingY = MovingY + MEGASCROLL
Skipp:
If MyNPC(MyIndex).MovingY = -2 Then
    MyNPC(MyIndex).MovingY = 0
    MyNPC(MyIndex).Walking = 0
    Exit Sub
End If
MyNPC(MyIndex).MovingY = MyNPC(MyIndex).MovingY + MEGASCROLL
End Sub
Public Sub WalkLeft()
If MyNPC(MyIndex).X > 38 Then GoTo Skipp
If MyX = 0 And MovingX = 0 Then GoTo Skipp
    If MovingX = 0 Then
    MovingX = 32
    MyX = MyX - 1
    End If
MovingX = MovingX - MEGASCROLL
Skipp:
If MyNPC(MyIndex).MovingX = 2 Then
    MyNPC(MyIndex).Walking = 0
    MyNPC(MyIndex).MovingX = 0
    Exit Sub
End If
MyNPC(MyIndex).MovingX = MyNPC(MyIndex).MovingX - MEGASCROLL
End Sub
Public Sub WalkRight()
If MyNPC(MyIndex).X < 10 Then GoTo Skipp
If MyX = 30 Then GoTo Skipp
If MovingX = 30 Then
    MyX = MyX + 1
    If MyX = 29 Then
        MovingX = 0
        GoTo Skipp
    End If
    MovingX = -2
End If
MovingX = MovingX + MEGASCROLL
Skipp:
If MyNPC(MyIndex).MovingX = -2 Then
    MyNPC(MyIndex).MovingX = 0
    MyNPC(MyIndex).Walking = 0
    Exit Sub
End If
MyNPC(MyIndex).MovingX = MyNPC(MyIndex).MovingX + MEGASCROLL
End Sub

Public Sub UpdateInfo(Data As String)
Dim Buffy As String
ResetBuf:
If Buffy <> "" Then
    Data = Buffy
    Buffy = ""
End If
If Left$(Data, 1) = "#" Then Data = Right$(Data, Len(Data) - 1)
If Left$(Data, 1) = "M" Then
UpdateMapInfo Right$(Data, Len(Data) - 1)
ElseIf Left$(Data, 1) = "N" Then
If InStr(1, Data, Chr(0)) > 0 Then
    Buffy = Right$(Data, Len(Data) - InStr(1, Data, Chr(0)))
    Data = Left$(Data, InStr(1, Data, Chr(0)) - 1)
End If
UpdateNPCInfo Right$(Data, Len(Data) - 2)
ElseIf Left$(Data, 1) = "$" Then
If InStr(1, Data, Chr(0)) > 0 Then
    Buffy = Right$(Data, Len(Data) - InStr(1, Data, Chr(0)))
    Data = Left$(Data, InStr(1, Data, Chr(0)) - 1)
End If
If Right$(Data, 1) = Chr(0) Then Data = Left$(Data, Len(Data) - 1)
DXEngine.MoveChat Right$(Data, Len(Data) - 1)
Else
modNPC.shat = "probi:" & Data
End If
If Buffy <> "" Then GoTo ResetBuf
End Sub
Public Sub UpdateNPCs()
Dim i As Integer
For i = 0 To NPCCount
If i = MyIndex Then GoTo Skip
If MyNPC(i).Active = False Then GoTo Skip
If MyNPC(i).Walking = 1 Then
    If MyNPC(i).MovingY = 2 Then
    MyNPC(i).MovingY = 0
    MyNPC(i).Walking = 0
    Exit Sub
    End If
MyNPC(i).MovingY = MyNPC(i).MovingY - MEGASCROLL
End If
If MyNPC(i).Walking = 2 Then
    If MyNPC(i).MovingY = -2 Then
    MyNPC(i).MovingY = 0
    MyNPC(i).Walking = 0
    Exit Sub
    End If
MyNPC(i).MovingY = MyNPC(i).MovingY + MEGASCROLL
End If
If MyNPC(i).Walking = 4 Then
    If MyNPC(i).MovingX = 2 Then
    MyNPC(i).MovingX = 0
    MyNPC(i).Walking = 0
    Exit Sub
    End If
MyNPC(i).MovingX = MyNPC(i).MovingX - MEGASCROLL
End If
If MyNPC(i).Walking = 3 Then
    If MyNPC(i).MovingX = -2 Then
    MyNPC(i).MovingX = 0
    MyNPC(i).Walking = 0
    Exit Sub
    End If
MyNPC(i).MovingX = MyNPC(i).MovingX + MEGASCROLL
End If
If MyNPC(i).Walking = 0 Then
    MyNPC(i).MovingX = 0
    MyNPC(i).MovingY = 0
End If
Skip:
Next
End Sub
Private Sub UpdateMapInfo(Data As String)
Dim StringBuf As String
Dim TempX As Byte
Dim TempY As Byte
If InStr(1, Data, Chr(0)) > 0 Then
    If InStr(1, Data, Chr(0)) < Len(Data) Then StringBuf = Right$(Data, Len(Data) - InStr(1, Data, Chr(0)))
    Data = Left$(Data, InStr(1, Data, Chr(0)) - 1)
End If
Select Case Left(Data, 1)
    Case "C":
    Data = Right$(Data, Len(Data) - InStr(1, Data, "("))
    TempX = Left$(Data, InStr(1, Data, ",") - 1)
    Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
    TempY = Left$(Data, InStr(1, Data, ")") - 1)
    Data = Right$(Data, Len(Data) - InStr(1, Data, ")"))
    DXEngine.AddItem TempX, TempY, Left(Data, InStr(1, Data, ",") - 1)
    Data = ""
    Case "D":
    Data = Right$(Data, Len(Data) - InStr(1, Data, "("))
    TempX = Left$(Data, InStr(1, Data, ",") - 1)
    Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
    TempY = Left$(Data, InStr(1, Data, ")") - 1)
    DXEngine.DestroyItem TempX, TempY
    Data = ""
    Case "N":
    'modNPC.shat = Data
    NPCCount = -1
    Data = Right$(Data, Len(Data) - 1)
    MyMapIndex = Left$(Data, InStr(1, Data, ","))
    DXEngine.Mapload Str$(MyMapIndex)
    MyMapIndex = MyMapIndex - 1
    Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
    If Left(Data, 1) = "/" Then GoTo Skip
    Do
    NPCCount = NPCCount + 1
    ReDim Preserve MyNPC(0 To NPCCount)
    Data = Right$(Data, Len(Data) - 1)
    MyNPC(NPCCount).X = Left$(Data, InStr(1, Data, ",") - 1)
    Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
    MyNPC(NPCCount).Y = Left$(Data, InStr(1, Data, ")") - 1)
    Data = Right$(Data, Len(Data) - InStr(1, Data, ")"))
    MyNPC(NPCCount).Body = Left$(Data, InStr(1, Data, ",") - 1)
    Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
    MyNPC(NPCCount).Head = Left$(Data, InStr(1, Data, ",") - 1)
    Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
    MyNPC(NPCCount).Weap = Left$(Data, InStr(1, Data, ",") - 1)
    Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
    MyNPC(NPCCount).Active = True
    MyNPC(NPCCount).LastMove = 2
    MyNPC(NPCCount).MovingX = 0
    MyNPC(NPCCount).MovingY = 0
    MyNPC(NPCCount).Walking = 0
    MyNPC(NPCCount).WeapStep = 0
    Loop Until Left(Data, 1) = "/"
Skip:
    Data = Right$(Data, Len(Data) - InStr(1, Data, "/"))
    MyIndex = Left$(Data, InStr(1, Data, "*") - 1)
    Data = Right$(Data, Len(Data) - InStr(1, Data, "*"))
    If Left$(Data, 1) = "(" Then
    Do
    Data = Right$(Data, Len(Data) - InStr(1, Data, "("))
    TempX = Left$(Data, InStr(1, Data, ",") - 1)
    Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
    TempY = Left$(Data, InStr(1, Data, ")") - 1)
    Data = Right$(Data, Len(Data) - InStr(1, Data, ")"))
    DXEngine.AddItem TempX, TempY, Left$(Data, InStr(1, Data, ",") - 1)
    Data = Right$(Data, Len(Data) - InStr(1, Data, "(") + 1)
    Loop While Left$(Data, 1) = "("
    'else no items
    End If
    MyNPC(MyIndex).Active = True
    ResetMyX MyNPC(MyIndex).X
    ResetMyY MyNPC(MyIndex).Y
    MyNPC(MyIndex).LastMove = 2
    Connected = True
    DXEngine.MainLoop
End Select
If StringBuf <> "" Then UpdateInfo StringBuf
End Sub

Private Sub UpdateNPCInfo(Data As String)
Dim TempIndex As String
Dim StringBuf As String
Dim TempX As Byte
Dim TempY As Byte
TempIndex = Left$(Data, InStr(1, Data, ")") - 1)
Data = Right$(Data, Len(Data) - Len(TempIndex) - 1)
If InStr(1, Data, Chr(0)) > 0 Then
    If InStr(1, Data, Chr(0)) < Len(Data) Then StringBuf = Right$(Data, Len(Data) - InStr(1, Data, Chr(0)))
    Data = Left$(Data, InStr(1, Data, Chr(0)) - 1)
End If
If Connected = False Then
    If StringBuf <> "" Then UpdateInfo StringBuf
    Exit Sub
End If
Select Case Left(Data, 1)
    Case "X":
    MyNPC(TempIndex).WeapStep = 0
    If TempIndex = MyIndex Then
        If Right$(Data, Len(Data) - 1) > MyNPC(MyIndex).X Then
            DXEngine.LastMove = 4
        Else
            DXEngine.LastMove = 3
        End If
        MyNPC(MyIndex).X = Right$(Data, Len(Data) - 1)
        ResetMyX MyNPC(MyIndex).X
        Exit Sub
    End If
    If MyNPC(TempIndex).X > Right$(Data, Len(Data) - 1) Then
        MyNPC(TempIndex).Walking = 4
        MyNPC(TempIndex).MovingX = 32
        MyNPC(TempIndex).LastMove = 4
    Else
        MyNPC(TempIndex).Walking = 3
        MyNPC(TempIndex).MovingX = -32
        MyNPC(TempIndex).LastMove = 3
    End If
    MyNPC(TempIndex).X = Right$(Data, Len(Data) - 1)
    If DXEngine.LastMove > 0 Then DXEngine.LastMove = 0
    MyNPC(TempIndex).MovingY = 0
    Case "Y":
    MyNPC(TempIndex).WeapStep = 0
    If TempIndex = MyIndex Then
        If Right$(Data, Len(Data) - 1) > MyNPC(MyIndex).Y Then
            DXEngine.LastMove = 1
        Else
            DXEngine.LastMove = 2
        End If
        MyNPC(MyIndex).Y = Right$(Data, Len(Data) - 1)
        ResetMyY MyNPC(MyIndex).Y
        Exit Sub
    End If
    If MyNPC(TempIndex).Y < Right$(Data, Len(Data) - 1) Then
        MyNPC(TempIndex).Walking = 2
        MyNPC(TempIndex).MovingY = -32
        MyNPC(TempIndex).LastMove = 2
    Else
        MyNPC(TempIndex).Walking = 1
        MyNPC(TempIndex).MovingY = 32
        MyNPC(TempIndex).LastMove = 1
    End If
    MyNPC(TempIndex).Y = Right$(Data, Len(Data) - 1)
    If DXEngine.LastMove > 0 Then DXEngine.LastMove = 0
    MyNPC(TempIndex).MovingX = 0
    Case "C":
    If TempIndex > NPCCount Then
        NPCCount = NPCCount + 1
        ReDim Preserve MyNPC(0 To TempIndex)
    End If
    Data = Right$(Data, Len(Data) - 1)
    MyNPC(TempIndex).Body = Left$(Data, InStr(1, Data, ",") - 1)
    Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
    MyNPC(TempIndex).Head = Left$(Data, InStr(1, Data, ",") - 1)
    Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
    MyNPC(TempIndex).Weap = Left$(Data, InStr(1, Data, ",") - 1)
    Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
    MyNPC(TempIndex).LastMove = 2
    MyNPC(TempIndex).X = Left$(Data, InStr(1, Data, ",") - 1)
    Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
    MyNPC(TempIndex).Y = Left$(Data, InStr(1, Data, ",") - 1)
    MyNPC(TempIndex).Active = True
    MyNPC(TempIndex).MovingX = 0
    MyNPC(TempIndex).MovingY = 0
    MyNPC(TempIndex).Walking = 0
    Case "x":
    Data = Right$(Data, Len(Data) - 1)
    TempX = val(Left$(Data, InStr(1, Data, ",") - 1))
    Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
    TempY = val(Left$(Data, InStr(1, Data, ",") - 1))
    Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
    TempIndex = Left$(Data, InStr(1, Data, ",") - 1)
    Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
    DXEngine.ChangeTile TempX, TempY, val(TempIndex), Data
    Case "K":
    MyNPC(TempIndex).Active = False
    DXEngine.LastMove = 0
    Case "L":
    DXEngine.ElListo = Right$(Data, Len(Data) - 1)
    Case "B":
    MyNPC(TempIndex).Body = Right$(Data, Len(Data) - 1)
    Case "W":
    MyNPC(TempIndex).Weap = val(Right$(Data, Len(Data) - 1))
    Case "w":
    Data = Right$(Data, Len(Data) - 1)
    MyNPC(TempIndex).X = Left$(Data, InStr(1, Data, ",") - 1)
    Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
    MyNPC(TempIndex).Y = Left$(Data, InStr(1, Data, ",") - 1)
    MyNPC(TempIndex).MovingX = 0
    MyNPC(TempIndex).MovingY = 0
    MyNPC(TempIndex).Walking = 0
    MyNPC(TempIndex).WeapStep = 0
    If TempIndex = MyIndex Then
        ResetMyX MyNPC(TempIndex).X
        ResetMyY MyNPC(TempIndex).Y
    End If
    Case "F":
    MyNPC(TempIndex).LastMove = val(Right$(Data, Len(Data) - 1))
    MyNPC(TempIndex).WeapStep = 1
    DXEngine.LastMove = 0 '5 controls attack speed
End Select
If StringBuf <> "" Then UpdateInfo StringBuf
End Sub
Private Sub ResetMyX(NPCX As Integer)
MyNPC(MyIndex).MovingX = 0
MyNPC(MyIndex).MovingY = 0
MyNPC(MyIndex).Walking = 0
DXEngine.MovingX = 0
DXEngine.MovingY = 0
If NPCX > 38 Then
DXEngine.MyX = 30
ElseIf NPCX < 10 Then
DXEngine.MyX = 0
Else
DXEngine.MyX = MyNPC(MyIndex).X - 9
End If

End Sub
Private Sub ResetMyY(NPCY As Integer)
MyNPC(MyIndex).MovingX = 0
MyNPC(MyIndex).MovingY = 0
MyNPC(MyIndex).Walking = 0
DXEngine.MovingX = 0
DXEngine.MovingY = 0
If NPCY > 43 Then
DXEngine.MyY = 39
ElseIf NPCY < 6 Then
DXEngine.MyY = 0
Else
DXEngine.MyY = MyNPC(MyIndex).Y - 5
End If
End Sub
