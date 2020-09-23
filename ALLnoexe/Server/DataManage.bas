Attribute VB_Name = "DataManage"
Option Explicit
Private Type Invent
    Index As Integer
    Amount As Integer
    Equipped As Boolean
End Type
Private Type Equipped
    Item As Integer
    BackPIndex As Integer
End Type
Private Type Atty
    HP As Integer
    Str As Integer
    Arm As Integer
    DSk As Integer
    ASk As Integer
    XP As Integer
    MaxHP As Integer
    DeadDrops(1) As Integer
End Type
'Private Enum Equipped
    'Armor = 0
    'Gloves
    'Helmet
    'Key
    'Money
    'Potion
    'Ring
    'Shield
    'Shoes
    'Weapon
'End Enum
Private Type NPC
    X As Integer
    Y As Integer
    Walking As Byte
    Body As Byte
    Head As Byte
    Weap As Byte
    Mobile As Boolean
    BPack(24) As Invent
    Money As Integer
    Namer As String
    Equipment(9) As Equipped
    Attribs As Atty
    AttMods As Atty
    Active As Boolean
    LastMove As Byte
    Attacking As Byte
    NextXP As Integer
    SkillP As Integer
End Type

Private Type CNPC
    X As Integer
    Y As Integer
    RX As Integer
    RY As Integer
    Walking As Byte
    Facing As Byte
    Body As Byte
    Head As Byte
    Mobile As Boolean
    NPCT As Byte
    Attribs As Atty
    Speech As String
    Name As String
    ReSpawn As Boolean
    DeathScript As Integer
End Type
Private Type NPCz
    Index As Integer
    Mobile As Boolean
    Body As Byte
    Head As Byte
    NPCT As Byte
    Attribs As Atty
    Speech As String
    Name As String
    DeathScript As Integer
End Type
Private Type ServMap
    TileProp As Byte
    NPC As NPCz
    GItem As Integer
    Script As Integer
End Type
Private Type ServerMap
    TileProp() As Byte
    NPC() As NPCz
    GItem() As Integer
    Script() As Integer
End Type
Private Type MapExits
    NorthE As Integer
    EastE As Integer
    SouthE As Integer
    WestE As Integer
End Type
Private Type IGObjects
    Name As String
    Value As Integer
    ObjType As Byte
    Power As Integer
    GIndex As Integer
    Description As String
    SIndex As Integer
End Type
Private Type MeNPCz
    MyNPC() As NPC
End Type
Private Type ZGItems
    MrX() As Integer
    MrY() As Integer
    MrItem() As Integer
End Type
Private Type ServerNPCz
    ServNPC() As CNPC
    NPCTotal As Integer
End Type
Public SMapArray(49, 49) As ServerMap
Public ALLMyNPC() As MeNPCz
Public ALLServNPC() As ServerNPCz
Public MyExits() As MapExits
Public MyObjects() As IGObjects
Public MyGitemz() As ZGItems
Public ElMapo As Integer
Public MapTotal As Integer
Private TotalObjects As Integer
Private NPCTimage As Integer
Public Sub DispArray()
Dim X As Byte
Dim Y As Byte
Dim TempMap As Integer
TempMap = CInt(MegaServer.Text1)
If TempMap < 0 Or TempMap > MapTotal Then Exit Sub
MegaServer.NPCText.Text = "NPC Info"
For X = 0 To 49
    For Y = 0 To 49
        If SMapArray(X, Y).NPC(TempMap).Index > -1 Then
            If SMapArray(X, Y).NPC(TempMap).Index <= ALLServNPC(TempMap).NPCTotal Then
                MegaServer.NPCText.Text = MegaServer.NPCText.Text + vbCrLf + _
                "(" + Str(X) + "," + Str(Y) + ")" + (ALLServNPC(TempMap).ServNPC(SMapArray(X, Y).NPC(TempMap).Index).Name)
            Else
                MegaServer.NPCText.Text = MegaServer.NPCText.Text + vbCrLf + _
                "(" + Str(X) + "," + Str(Y) + ")" + (ALLMyNPC(TempMap).MyNPC(SMapArray(X, Y).NPC(TempMap).Index - ALLServNPC(TempMap).NPCTotal - 1).Namer)
            End If
        End If
    Next
Next
End Sub
Public Sub UpdateInfo(Userindex As Integer, Data As String)
Data = Right$(Data, Len(Data) - 1)
If Left$(Data, 1) = "M" Then
Data = Right$(Data, Len(Data) - 1)
HandleSingleData Userindex, Data
Else
Data = Right$(Data, Len(Data) - 1)
HandleMultData Userindex, Data
End If
End Sub

Private Sub HandleMultData(Userindex As Integer, Data As String)
Select Case Left$(Data, 1)
    Case "X":
    NPCMoveX Userindex, Val(Right$(Data, Len(Data) - 1))
    Case "Y":
    NPCMoveY Userindex, Val(Right$(Data, Len(Data) - 1))
    Case "A":
    ALLMyNPC(ElMapo).MyNPC(Userindex).Attacking = 1
    Case "E":
    Equip Userindex, Val(Right$(Data, Len(Data) - 1))
    Case "P":
    PickupGItem Userindex
    Case "L":
    SendItemList Userindex
    Case "U":
    DropInventItem Userindex, Val(Right$(Data, Len(Data) - 1))
    Case "T":
    Data = Right$(Data, Len(Data) - 1)
    Trade Userindex, Val(Data)
    Case "B":
    Data = Right$(Data, Len(Data) - 1)
    Buy Userindex, Val(Left$(Data, 1)), Val(Right$(Data, Len(Data) - 1))
    Case "a":
    Data = Right$(Data, Len(Data) - 1)
    UserAdjSkill Userindex, Val(Data)
    Case "s":
    SendStats Userindex
    Case "S":
    Data = Right$(Data, Len(Data) - 1)
    SellItem Userindex, Val(Left$(Data, 1)), Val(Right$(Data, Len(Data) - 1))
End Select
End Sub
Private Sub DelInventItem(Userindex As Integer, BPackIndex As Byte)
If ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(BPackIndex).Index > -1 Then
    ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(BPackIndex).Amount = ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(BPackIndex).Amount - 1
    If ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(BPackIndex).Amount = 0 Then ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(BPackIndex).Index = -1
End If
End Sub
Private Sub UserAdjSkill(ByRef Userindex As Integer, MrSkill As Integer)
If MrSkill > 0 And ALLMyNPC(ElMapo).MyNPC(Userindex).SkillP > 0 Then
    Select Case MrSkill
        Case 1:
            ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.ASk = ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.ASk + 1
        Case 2:
            ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.DSk = ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.DSk + 1
    End Select
    ALLMyNPC(ElMapo).MyNPC(Userindex).SkillP = ALLMyNPC(ElMapo).MyNPC(Userindex).SkillP - 1
End If
End Sub
Private Sub SendStats(ByRef NPCIndex As Integer)
Dim Stringer As String
Stringer = "#N(" & NPCIndex & ")L"
Stringer = Stringer & "Cash: " & ALLMyNPC(ElMapo).MyNPC(NPCIndex).Money & ";"
Stringer = Stringer & "HP: " & ALLMyNPC(ElMapo).MyNPC(NPCIndex).Attribs.HP & ";"
Stringer = Stringer & "MaxHP: " & ALLMyNPC(ElMapo).MyNPC(NPCIndex).Attribs.MaxHP & ";"
Stringer = Stringer & "Str: " & ALLMyNPC(ElMapo).MyNPC(NPCIndex).Attribs.Str & ";"
Stringer = Stringer & "Arm: " & ALLMyNPC(ElMapo).MyNPC(NPCIndex).Attribs.Arm & ";"
Stringer = Stringer & "Ask: " & ALLMyNPC(ElMapo).MyNPC(NPCIndex).Attribs.ASk & ";"
Stringer = Stringer & "Dsk: " & ALLMyNPC(ElMapo).MyNPC(NPCIndex).Attribs.DSk & ";"
Stringer = Stringer & "XP: " & ALLMyNPC(ElMapo).MyNPC(NPCIndex).Attribs.XP & ";"
Stringer = Stringer & "ReqXP: " & ALLMyNPC(ElMapo).MyNPC(NPCIndex).NextXP & ";"
Stringer = Stringer & "SP: " & ALLMyNPC(ElMapo).MyNPC(NPCIndex).SkillP & ";"
WSock2.SendDataToAClient ElMapo, NPCIndex, Stringer
End Sub
Public Sub TakeInventItem(Userindex As Integer, ItemIndex As Integer)
Dim Temp As Integer
For Temp = 0 To 24
    If ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(Temp).Index = ItemIndex Then
        ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(Temp).Amount = ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(Temp).Amount - 1
        If ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(Temp).Amount = 0 Then ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(Temp).Index = -1
        Exit For
    End If
Next
End Sub
Private Sub HandleSingleData(Userindex As Integer, Data As String)
Dim Loopi As Integer
Dim intXCounter As Byte
Dim intYCounter As Byte
Dim TempS As String
Select Case Left$(Data, 1)
    Case "N":
        Data = Right$(Data, Len(Data) - 1)
        ALLMyNPC(ElMapo).MyNPC(Userindex).Body = Left$(Data, InStr(1, Data, ",") - 1)
        Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
        ALLMyNPC(ElMapo).MyNPC(Userindex).Head = Left$(Data, InStr(1, Data, ",") - 1)
        Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
        ALLMyNPC(ElMapo).MyNPC(Userindex).Weap = Left$(Data, InStr(1, Data, ",") - 1)
        Data = Right$(Data, Len(Data) - InStr(1, Data, ","))
        ALLMyNPC(ElMapo).MyNPC(Userindex).Namer = Data
        ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.Arm = 0
        ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.Str = 1
        ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.ASk = 3
        ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.DSk = 3
        ALLMyNPC(ElMapo).MyNPC(Userindex).Money = 0
        ALLMyNPC(ElMapo).MyNPC(Userindex).NextXP = 200
        ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.HP = 30
        ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.MaxHP = 30
        SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).NPC(ElMapo).Index = Userindex + ALLServNPC(ElMapo).NPCTotal + 1
        TempS = "#MN1,"
        If ALLServNPC(ElMapo).NPCTotal > 0 Then
            For Loopi = 1 To ALLServNPC(ElMapo).NPCTotal
                If ALLServNPC(ElMapo).ServNPC(Loopi).ReSpawn = False Then
                    TempS = TempS & "(" & ALLServNPC(ElMapo).ServNPC(Loopi).X & "," & ALLServNPC(ElMapo).ServNPC(Loopi).Y & ")" & ALLServNPC(ElMapo).ServNPC(Loopi).Body & "," & ALLServNPC(ElMapo).ServNPC(Loopi).Head & ",0,"
                Else
                    TempS = TempS & "(" & ALLServNPC(ElMapo).ServNPC(Loopi).X & "," & ALLServNPC(ElMapo).ServNPC(Loopi).Y & ")" & 5 & "," & ALLServNPC(ElMapo).ServNPC(Loopi).Head & ",0,"
                End If
            Next
        End If
            For Loopi = 0 To MapSock(ElMapo).MaxCon
                If MapSock(ElMapo).Sockers(Loopi).Active Then
                    TempS = TempS & "(" & ALLMyNPC(ElMapo).MyNPC(Loopi).X & "," & ALLMyNPC(ElMapo).MyNPC(Loopi).Y & ")" & ALLMyNPC(ElMapo).MyNPC(Loopi).Body & "," & ALLMyNPC(ElMapo).MyNPC(Loopi).Head & "," & ALLMyNPC(ElMapo).MyNPC(Loopi).Weap & ","
                End If
            Next
            TempS = TempS & "/" & Userindex + ALLServNPC(ElMapo).NPCTotal & "*"
    For intXCounter = 0 To 49
        For intYCounter = 0 To 49
            If SMapArray(intXCounter, intYCounter).GItem(ElMapo) > -1 And SMapArray(intXCounter, intYCounter).TileProp(ElMapo) <> 1 Then
                TempS = TempS & "(" & intXCounter & "," & intYCounter & ")" & MyObjects(SMapArray(intXCounter, intYCounter).GItem(ElMapo)).GIndex & ","
            End If
        Next
    Next
            WSock2.SendDataToAClient 0, Userindex, TempS
            If MapSock(ElMapo).MaxCon > 0 Then SendDataToClients ElMapo, Userindex, "#N(" & Userindex + ALLServNPC(ElMapo).NPCTotal & ")C" & ALLMyNPC(ElMapo).MyNPC(Userindex).Body & "," & ALLMyNPC(ElMapo).MyNPC(Userindex).Head & "," & ALLMyNPC(ElMapo).MyNPC(Userindex).Weap & "," & ALLMyNPC(ElMapo).MyNPC(Userindex).X & "," & ALLMyNPC(ElMapo).MyNPC(Userindex).Y & ","
            For Loopi = 0 To 9
                ALLMyNPC(ElMapo).MyNPC(Userindex).Equipment(Loopi).BackPIndex = -1
                ALLMyNPC(ElMapo).MyNPC(Userindex).Equipment(Loopi).Item = -1
            Next Loopi
            ALLMyNPC(ElMapo).MyNPC(Userindex).Active = True
End Select
MegaServer.RndTimer.Enabled = True
End Sub
Public Sub ServerOpenIt(strMapName As String)
Dim intFreeFile As Integer
Dim intXCounter As Byte
Dim intYCounter As Byte
Dim MapCounter As Integer
Dim TempMap(49, 49) As ServMap
Dim TempNames() As String
'gotta redim all the crap
intFreeFile = FreeFile
Open strMapName For Input As #intFreeFile
    Input #intFreeFile, MapTotal
    ReDim TempNames(0 To MapTotal)
    For intXCounter = 0 To MapTotal
        Input #intFreeFile, intYCounter
        Input #intFreeFile, TempNames(intXCounter)
    Next
Close #intFreeFile
ReDim ALLServNPC(0 To MapTotal)
ReDim MyGitemz(0 To MapTotal)
ReDim MyExits(0 To MapTotal)
ReDim ALLMyNPC(0 To MapTotal)
ReDim MapSock(0 To MapTotal)
For intXCounter = 0 To MapTotal
    MapSock(intXCounter).MaxCon = 0
    ReDim MapSock(intXCounter).Sockers(0 To 0)
    ReDim ALLServNPC(intXCounter).ServNPC(0 To 0)
    ReDim ALLMyNPC(intXCounter).MyNPC(0 To 0)
    ReDim MyGitemz(intXCounter).MrItem(0 To 0)
    ReDim MyGitemz(intXCounter).MrX(0 To 0)
    ReDim MyGitemz(intXCounter).MrY(0 To 0)
Next
For intXCounter = 0 To 49
    For intYCounter = 0 To 49
        ReDim Preserve SMapArray(intXCounter, intYCounter).GItem(0 To MapTotal)
        ReDim Preserve SMapArray(intXCounter, intYCounter).Script(0 To MapTotal)
        ReDim Preserve SMapArray(intXCounter, intYCounter).TileProp(0 To MapTotal)
        ReDim Preserve SMapArray(intXCounter, intYCounter).NPC(0 To MapTotal)
    Next
Next
For MapCounter = 0 To MapTotal
    intFreeFile = FreeFile
    strMapName = App.Path & "\Maps\" & TempNames(MapCounter)
    Open strMapName For Binary As intFreeFile
    For intXCounter = 0 To 49
        For intYCounter = 0 To 49
            Get intFreeFile, , TempMap(intXCounter, intYCounter)
            If TempMap(intXCounter, intYCounter).GItem > -1 Then
                MyGitemz(MapCounter).MrItem(UBound(MyGitemz(MapCounter).MrItem, 1)) = TempMap(intXCounter, intYCounter).GItem
                MyGitemz(MapCounter).MrX(UBound(MyGitemz(MapCounter).MrX, 1)) = intXCounter
                MyGitemz(MapCounter).MrY(UBound(MyGitemz(MapCounter).MrY, 1)) = intYCounter
                ReDim Preserve MyGitemz(MapCounter).MrItem(0 To UBound(MyGitemz(MapCounter).MrItem, 1) + 1)
                ReDim Preserve MyGitemz(MapCounter).MrX(0 To UBound(MyGitemz(MapCounter).MrX, 1) + 1)
                ReDim Preserve MyGitemz(MapCounter).MrY(0 To UBound(MyGitemz(MapCounter).MrY, 1) + 1)
            End If
        Next
    Next
    Get intFreeFile, , ALLServNPC(MapCounter).NPCTotal
    Get intFreeFile, , MyExits(MapCounter).NorthE
    Get intFreeFile, , MyExits(MapCounter).SouthE
    Get intFreeFile, , MyExits(MapCounter).EastE
    Get intFreeFile, , MyExits(MapCounter).WestE
    Close intFreeFile
    For intXCounter = 0 To 49
        For intYCounter = 0 To 49
            SMapArray(intXCounter, intYCounter).TileProp(MapCounter) = TempMap(intXCounter, intYCounter).TileProp
            SMapArray(intXCounter, intYCounter).Script(MapCounter) = TempMap(intXCounter, intYCounter).Script
            SMapArray(intXCounter, intYCounter).NPC(MapCounter) = TempMap(intXCounter, intYCounter).NPC
            SMapArray(intXCounter, intYCounter).GItem(MapCounter) = TempMap(intXCounter, intYCounter).GItem
        Next
    Next
Next
NPCSetup
End Sub
Private Sub NPCSetup()
Dim intXCounter As Byte
Dim intYCounter As Byte
Dim TempIndex As Integer
Dim intTempy As Integer
For intTempy = 0 To UBound(SMapArray(0, 0).TileProp, 1)
    If ALLServNPC(intTempy).NPCTotal > 0 Then
        ReDim ALLServNPC(intTempy).ServNPC(1 To ALLServNPC(intTempy).NPCTotal)
        For intXCounter = 0 To 49
            For intYCounter = 0 To 49
                If SMapArray(intXCounter, intYCounter).NPC(intTempy).Index > -1 Then
                    TempIndex = SMapArray(intXCounter, intYCounter).NPC(intTempy).Index
                    ALLServNPC(intTempy).ServNPC(TempIndex).X = intXCounter
                    ALLServNPC(intTempy).ServNPC(TempIndex).Y = intYCounter
                    ALLServNPC(intTempy).ServNPC(TempIndex).RX = ALLServNPC(intTempy).ServNPC(TempIndex).X
                    ALLServNPC(intTempy).ServNPC(TempIndex).RY = ALLServNPC(intTempy).ServNPC(TempIndex).Y
                    ALLServNPC(intTempy).ServNPC(TempIndex).Body = SMapArray(intXCounter, intYCounter).NPC(intTempy).Body
                    ALLServNPC(intTempy).ServNPC(TempIndex).Head = SMapArray(intXCounter, intYCounter).NPC(intTempy).Head
                    ALLServNPC(intTempy).ServNPC(TempIndex).Mobile = SMapArray(intXCounter, intYCounter).NPC(intTempy).Mobile
                    ALLServNPC(intTempy).ServNPC(TempIndex).Speech = SMapArray(intXCounter, intYCounter).NPC(intTempy).Speech
                    ALLServNPC(intTempy).ServNPC(TempIndex).Name = SMapArray(intXCounter, intYCounter).NPC(intTempy).Name
                    ALLServNPC(intTempy).ServNPC(TempIndex).NPCT = SMapArray(intXCounter, intYCounter).NPC(intTempy).NPCT
                    ALLServNPC(intTempy).ServNPC(TempIndex).ReSpawn = False
                    ALLServNPC(intTempy).ServNPC(TempIndex).Attribs = SMapArray(intXCounter, intYCounter).NPC(intTempy).Attribs
                    ALLServNPC(intTempy).ServNPC(TempIndex).Attribs.MaxHP = ALLServNPC(intTempy).ServNPC(TempIndex).Attribs.HP
                   ALLServNPC(intTempy).ServNPC(TempIndex).DeathScript = SMapArray(intXCounter, intYCounter).NPC(intTempy).DeathScript
                End If
            Next
        Next
    End If
Next
End Sub


Public Sub MoveAllServNPC(Optional SingMap As Integer, Optional SingIndex As Integer, Optional ByPass As Boolean)
Dim i As Integer
Dim CurMapola As Integer
If ByPass Then
    i = SingIndex
    CurMapola = SingMap
    GoTo StartMove
End If
NPCTimage = NPCTimage + 1
For CurMapola = 0 To UBound(ALLServNPC, 1)
    'do client attacks
    If NPCTimage Mod 2 = 0 Then
        'MegaServer.GameChat.Text = MegaServer.GameChat.Text + vbCrLf + "fired at" + Str(NPCTimage)
        For i = 0 To UBound(ALLMyNPC(CurMapola).MyNPC, 1)
            If ALLMyNPC(CurMapola).MyNPC(i).Attacking = 1 And ALLMyNPC(CurMapola).MyNPC(i).Attribs.HP > 0 Then Combat.ClientAttack CurMapola, i
            ALLMyNPC(CurMapola).MyNPC(i).Attacking = 0
        Next
    End If
    'do npc moves/call attacks
    For i = 1 To ALLServNPC(CurMapola).NPCTotal
        If ALLServNPC(CurMapola).ServNPC(i).Mobile Then
            If ALLServNPC(CurMapola).ServNPC(i).NPCT <> 2 And ALLServNPC(CurMapola).ServNPC(i).ReSpawn = False Then
StartMove:
            Select Case Int((8 * Rnd) + 1)
            Case 1:
            If ALLServNPC(CurMapola).ServNPC(i).Y > 1 Then
                If SMapArray(ALLServNPC(CurMapola).ServNPC(i).X, ALLServNPC(CurMapola).ServNPC(i).Y - 1).TileProp(CurMapola) > 2 And SMapArray(ALLServNPC(CurMapola).ServNPC(i).X, ALLServNPC(CurMapola).ServNPC(i).Y - 1).NPC(CurMapola).Index = -1 Then
                    SMapArray(ALLServNPC(CurMapola).ServNPC(i).X, ALLServNPC(CurMapola).ServNPC(i).Y).NPC(CurMapola).Index = -1
                    ALLServNPC(CurMapola).ServNPC(i).Y = ALLServNPC(CurMapola).ServNPC(i).Y - 1
                    SMapArray(ALLServNPC(CurMapola).ServNPC(i).X, ALLServNPC(CurMapola).ServNPC(i).Y).NPC(CurMapola).Index = i
                    WSock2.SendDataToClients CurMapola, -1, "#N(" & i - 1 & ")Y" & ALLServNPC(CurMapola).ServNPC(i).Y
                End If
            End If
            Case 2:
            If ALLServNPC(CurMapola).ServNPC(i).Y < 48 Then
                If SMapArray(ALLServNPC(CurMapola).ServNPC(i).X, ALLServNPC(CurMapola).ServNPC(i).Y + 1).TileProp(CurMapola) > 2 And SMapArray(ALLServNPC(CurMapola).ServNPC(i).X, ALLServNPC(CurMapola).ServNPC(i).Y + 1).NPC(CurMapola).Index = -1 Then
                    SMapArray(ALLServNPC(CurMapola).ServNPC(i).X, ALLServNPC(CurMapola).ServNPC(i).Y).NPC(CurMapola).Index = -1
                    ALLServNPC(CurMapola).ServNPC(i).Y = ALLServNPC(CurMapola).ServNPC(i).Y + 1
                    SMapArray(ALLServNPC(CurMapola).ServNPC(i).X, ALLServNPC(CurMapola).ServNPC(i).Y).NPC(CurMapola).Index = i
                    WSock2.SendDataToClients CurMapola, -1, "#N(" & i - 1 & ")Y" & ALLServNPC(CurMapola).ServNPC(i).Y
                End If
            End If
            Case 3:
            If ALLServNPC(CurMapola).ServNPC(i).X < 48 Then
                If SMapArray(ALLServNPC(CurMapola).ServNPC(i).X + 1, ALLServNPC(CurMapola).ServNPC(i).Y).TileProp(CurMapola) > 2 And SMapArray(ALLServNPC(CurMapola).ServNPC(i).X + 1, ALLServNPC(CurMapola).ServNPC(i).Y).NPC(CurMapola).Index = -1 Then
                    SMapArray(ALLServNPC(CurMapola).ServNPC(i).X, ALLServNPC(CurMapola).ServNPC(i).Y).NPC(CurMapola).Index = -1
                    ALLServNPC(CurMapola).ServNPC(i).X = ALLServNPC(CurMapola).ServNPC(i).X + 1
                    SMapArray(ALLServNPC(CurMapola).ServNPC(i).X, ALLServNPC(CurMapola).ServNPC(i).Y).NPC(CurMapola).Index = i
                    WSock2.SendDataToClients CurMapola, -1, "#N(" & i - 1 & ")X" & ALLServNPC(CurMapola).ServNPC(i).X
                End If
            End If
            Case 4:
            If ALLServNPC(CurMapola).ServNPC(i).X > 1 Then
                If SMapArray(ALLServNPC(CurMapola).ServNPC(i).X - 1, ALLServNPC(CurMapola).ServNPC(i).Y).TileProp(CurMapola) > 2 And SMapArray(ALLServNPC(CurMapola).ServNPC(i).X - 1, ALLServNPC(CurMapola).ServNPC(i).Y).NPC(CurMapola).Index = -1 Then
                    SMapArray(ALLServNPC(CurMapola).ServNPC(i).X, ALLServNPC(CurMapola).ServNPC(i).Y).NPC(CurMapola).Index = -1
                    ALLServNPC(CurMapola).ServNPC(i).X = ALLServNPC(CurMapola).ServNPC(i).X - 1
                    SMapArray(ALLServNPC(CurMapola).ServNPC(i).X, ALLServNPC(CurMapola).ServNPC(i).Y).NPC(CurMapola).Index = i
                    WSock2.SendDataToClients CurMapola, -1, "#N(" & i - 1 & ")X" & ALLServNPC(CurMapola).ServNPC(i).X
                End If
            End If
            Case Else 'nada
            End Select
            If ByPass Then Exit Sub
            Else
                If NPCTimage Mod 2 = 0 Then
                    Combat.AISeek CurMapola, i, True
                Else
                    Combat.AISeek CurMapola, i, False
                End If
            End If
        End If
    Next
Next
If NPCTimage = 420 Then
    RespawnNPC
    NPCTimage = 0
End If
End Sub
Public Sub OpenObjFile(strMapName As String)
Dim intCounter As Byte
Dim intFreeFile As Integer
intFreeFile = FreeFile
Open strMapName For Binary As intFreeFile
    Get intFreeFile, , TotalObjects
    ReDim MyObjects(0 To TotalObjects - 1)
    For intCounter = 0 To TotalObjects - 1
       Get intFreeFile, , MyObjects(intCounter)
    Next
Close intFreeFile
End Sub
Public Sub AddItemToInvent(PlayerID As Integer, ItemID As Integer)
Dim Temp As Integer
If MyObjects(ItemID).ObjType <> 4 Then
    For Temp = 0 To 24
        If ALLMyNPC(ElMapo).MyNPC(PlayerID).BPack(Temp).Index = -1 Or ALLMyNPC(ElMapo).MyNPC(PlayerID).BPack(Temp).Index = ItemID Then
            ALLMyNPC(ElMapo).MyNPC(PlayerID).BPack(Temp).Index = ItemID
            ALLMyNPC(ElMapo).MyNPC(PlayerID).BPack(Temp).Amount = ALLMyNPC(ElMapo).MyNPC(PlayerID).BPack(Temp).Amount + 1
            Exit Sub
        End If
    Next
Else
    ALLMyNPC(ElMapo).MyNPC(PlayerID).Money = ALLMyNPC(ElMapo).MyNPC(PlayerID).Money + MyObjects(ItemID).Value
End If
End Sub
Public Function CheckInvent(PlayerID As Integer, ItemIndex As Integer) As Boolean
Dim Temp As Integer
CheckInvent = False
    For Temp = 0 To 24
        If ALLMyNPC(ElMapo).MyNPC(PlayerID).BPack(Temp).Index = ItemIndex Then
            CheckInvent = True
            Exit Function
        End If
    Next
End Function
Private Sub Trade(Userindex As Integer, Dir As Byte)
Dim i As Integer
Dim Stringy As String
Dim NPCIndex As Integer
Select Case Dir
    Case 1: NPCIndex = SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y - 1).NPC(ElMapo).Index
    Case 2: NPCIndex = SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y + 1).NPC(ElMapo).Index
    Case 3: NPCIndex = SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X + 1, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).NPC(ElMapo).Index
    Case 4: NPCIndex = SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X - 1, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).NPC(ElMapo).Index
End Select
If NPCIndex = -1 Or NPCIndex > ALLServNPC(ElMapo).NPCTotal Then Exit Sub
If ALLServNPC(ElMapo).ServNPC(NPCIndex).NPCT = 1 Then
    Stringy = "#N(" & NPCIndex & ")L"
    i = ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.HP
    Stringy = Stringy & MyObjects(ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.Str).Name & ":" & ((MyObjects(ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.Str).Value * (i / 100))) & ";"
    Stringy = Stringy & MyObjects(ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.Arm).Name & ":" & ((MyObjects(ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.Arm).Value * (i / 100))) & ";"
    Stringy = Stringy & MyObjects(ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.DSk).Name & ":" & ((MyObjects(ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.DSk).Value * (i / 100))) & ";"
    Stringy = Stringy & MyObjects(ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.ASk).Name & ":" & ((MyObjects(ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.ASk).Value * (i / 100))) & ";"
    Stringy = Stringy & MyObjects(ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.XP).Name & ":" & ((MyObjects(ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.XP).Value * (i / 100))) & ";"
    WSock2.SendDataToAClient ElMapo, Userindex, Stringy
End If
If Len(ALLServNPC(ElMapo).ServNPC(NPCIndex).Speech) > 0 Then
    WSock2.SendDataToAClient ElMapo, Userindex, "$" & ALLServNPC(ElMapo).ServNPC(NPCIndex).Name & ": " & ALLServNPC(ElMapo).ServNPC(NPCIndex).Speech
End If
End Sub
Private Sub Buy(Userindex As Integer, Dir As Byte, ListIndex As Byte)
Dim Price As Integer
Dim CNPCTemp As Integer
Select Case Dir
    Case 1: CNPCTemp = SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y - 1).NPC(ElMapo).Index
    Case 2: CNPCTemp = SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y + 1).NPC(ElMapo).Index
    Case 3: CNPCTemp = SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X + 1, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).NPC(ElMapo).Index
    Case 4: CNPCTemp = SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X - 1, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).NPC(ElMapo).Index
End Select
If CNPCTemp = -1 Or CNPCTemp > ALLServNPC(ElMapo).NPCTotal Then Exit Sub
If ALLServNPC(ElMapo).ServNPC(CNPCTemp).NPCT = 1 Then
    Price = GetPrice(ListIndex, CNPCTemp)
    If ALLMyNPC(ElMapo).MyNPC(Userindex).Money >= Price Then
        AddItemToInvent Userindex, GetItem(ListIndex, CNPCTemp)
        ALLMyNPC(ElMapo).MyNPC(Userindex).Money = ALLMyNPC(ElMapo).MyNPC(Userindex).Money - Price
        WSock2.SendDataToAClient ElMapo, Userindex, "$Bought it for " & Price & ".  You have " & ALLMyNPC(ElMapo).MyNPC(Userindex).Money & " left."
    End If
End If
End Sub
Private Sub SellItem(Userindex As Integer, Dir As Integer, ListIndex As Byte)
Dim NPCIndex As Integer
Select Case Dir
    Case 1: NPCIndex = SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y - 1).NPC(ElMapo).Index
    Case 2: NPCIndex = SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y + 1).NPC(ElMapo).Index
    Case 3: NPCIndex = SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X + 1, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).NPC(ElMapo).Index
    Case 4: NPCIndex = SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X - 1, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).NPC(ElMapo).Index
End Select
If NPCIndex = -1 Or NPCIndex > ALLServNPC(ElMapo).NPCTotal Then Exit Sub
If ALLServNPC(ElMapo).ServNPC(NPCIndex).NPCT <> 1 Or ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(ListIndex).Index = -1 Then Exit Sub
If ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(ListIndex).Equipped = True And ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(ListIndex).Amount < 2 Then Exit Sub
ALLMyNPC(ElMapo).MyNPC(Userindex).Money = ALLMyNPC(ElMapo).MyNPC(Userindex).Money + (ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.HP / 100 * MyObjects(ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(ListIndex).Index).Value)
DelInventItem Userindex, ListIndex
WSock2.SendDataToAClient ElMapo, Userindex, "$You now have " & ALLMyNPC(ElMapo).MyNPC(Userindex).Money
End Sub
Private Function GetPrice(ListIndex As Byte, NPCIndex As Integer) As Integer
Dim i As Integer
i = ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.HP
Select Case ListIndex
    Case 0: GetPrice = (MyObjects(ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.Str).Value * (i / 100))
    Case 1: GetPrice = (MyObjects(ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.Arm).Value * (i / 100))
    Case 2: GetPrice = (MyObjects(ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.DSk).Value * (i / 100))
    Case 3: GetPrice = (MyObjects(ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.ASk).Value * (i / 100))
    Case 4: GetPrice = (MyObjects(ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.XP).Value * (i / 100))
End Select
End Function
Private Function GetItem(ListIndex As Byte, NPCIndex As Integer) As Integer
Select Case ListIndex
    Case 0: GetItem = ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.Str
    Case 1: GetItem = ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.Arm
    Case 2: GetItem = ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.DSk
    Case 3: GetItem = ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.ASk
    Case 4: GetItem = ALLServNPC(ElMapo).ServNPC(NPCIndex).Attribs.XP
End Select
End Function

Private Sub SendItemList(Userindex As Integer)
Dim Stringy As String
Dim i As Integer
Stringy = "#N(" & Userindex + ALLServNPC(ElMapo).NPCTotal & ")L"
For i = 0 To 24
    If ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(i).Index > -1 Then
        Stringy = Stringy & MyObjects(ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(i).Index).Name & ":" & ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(i).Amount
        If ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(i).Equipped Then
            Stringy = Stringy & " E"
        End If
        Stringy = Stringy & ";"
    Else
        Stringy = Stringy & "None" & ";"
    End If
Next
WSock2.SendDataToAClient ElMapo, Userindex, Stringy
End Sub

Private Sub DropInventItem(Userindex As Integer, ListIndex As Byte)
If ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(ListIndex).Index > -1 And SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).TileProp(ElMapo) > 1 Then
    If ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(ListIndex).Equipped = True And ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(ListIndex).Amount < 2 Then
        WSock2.SendDataToAClient ElMapo, Userindex, "$You can't drop that equipped item since you only have 1."
        Exit Sub
    End If
    SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).GItem(ElMapo) = ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(ListIndex).Index
    DelInventItem Userindex, ListIndex
    WSock2.SendDataToClients ElMapo, -1, "#MC(" & ALLMyNPC(ElMapo).MyNPC(Userindex).X & "," & ALLMyNPC(ElMapo).MyNPC(Userindex).Y & ")" & MyObjects(SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).GItem(ElMapo)).GIndex & ","
End If
End Sub

Private Sub PickupGItem(Userindex As Integer)
If SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).GItem(ElMapo) > -1 And SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).TileProp(ElMapo) > 1 Then
    AddItemToInvent Userindex, SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).GItem(ElMapo)
    WSock2.SendDataToClients ElMapo, -1, "#MD(" & ALLMyNPC(ElMapo).MyNPC(Userindex).X & "," & ALLMyNPC(ElMapo).MyNPC(Userindex).Y & ")"
    WSock2.SendDataToAClient ElMapo, Userindex, "$*Picked up a " & MyObjects(SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).GItem(ElMapo)).Name & "* "
    SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).GItem(ElMapo) = -1
End If
End Sub

Private Sub NPCMoveX(Userindex As Integer, X As Byte)
If ALLMyNPC(ElMapo).MyNPC(Userindex).X > X Then
    ALLMyNPC(ElMapo).MyNPC(Userindex).LastMove = 4
Else
    ALLMyNPC(ElMapo).MyNPC(Userindex).LastMove = 3
End If
If SMapArray(X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).TileProp(ElMapo) >= 2 And SMapArray(X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).NPC(ElMapo).Index = -1 Then
    If Abs(X - ALLMyNPC(ElMapo).MyNPC(Userindex).X) > 1 Then
        MegaServer.GameChat.Text = MegaServer.GameChat.Text + vbCrLf + MapSock(ElMapo).Sockers(Userindex).Name + "Skipped a space"   'niner
        Warp Userindex, ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y, ElMapo
        Exit Sub
    End If
    SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).NPC(ElMapo).Index = -1
    ALLMyNPC(ElMapo).MyNPC(Userindex).X = X
    SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).NPC(ElMapo).Index = Userindex + ALLServNPC(ElMapo).NPCTotal + 1
    WSock2.SendDataToClients ElMapo, Userindex, "#N(" & Userindex + ALLServNPC(ElMapo).NPCTotal & ")X" & ALLMyNPC(ElMapo).MyNPC(Userindex).X
    If SMapArray(X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).Script(ElMapo) > -1 Then
        ScriptMod.CurScript = SMapArray(X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).Script(ElMapo)
        ScriptMod.PlayerSocket = Userindex
        ScriptMod.Execute
    End If
Else
    If SMapArray(X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).TileProp(ElMapo) = 1 And SMapArray(X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).NPC(ElMapo).Index = -1 Then
        If ALLMyNPC(ElMapo).MyNPC(Userindex).Equipment(3).Item = SMapArray(X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).GItem(ElMapo) Then
            SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).NPC(ElMapo).Index = -1
            ALLMyNPC(ElMapo).MyNPC(Userindex).X = X
            SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).NPC(ElMapo).Index = Userindex + ALLServNPC(ElMapo).NPCTotal + 1
            WSock2.SendDataToClients ElMapo, Userindex, "#N(" & Userindex + ALLServNPC(ElMapo).NPCTotal & ")X" & ALLMyNPC(ElMapo).MyNPC(Userindex).X
            If SMapArray(X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).Script(ElMapo) > -1 Then
                ScriptMod.CurScript = SMapArray(X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).Script(ElMapo)
                ScriptMod.PlayerSocket = Userindex
                ScriptMod.Execute
            End If
            Exit Sub
        End If
    End If
    WSock2.SendDataToAClient ElMapo, Userindex, "#N(" & Userindex + ALLServNPC(ElMapo).NPCTotal & ")X" & ALLMyNPC(ElMapo).MyNPC(Userindex).X
End If
If X = 0 Or X = 49 Then ChangeMap Userindex, False, -1
End Sub

Private Sub NPCMoveY(Userindex As Integer, Y As Byte)
If ALLMyNPC(ElMapo).MyNPC(Userindex).Y > Y Then
    ALLMyNPC(ElMapo).MyNPC(Userindex).LastMove = 1
Else
    ALLMyNPC(ElMapo).MyNPC(Userindex).LastMove = 2
End If
If SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, Y).TileProp(ElMapo) >= 2 And SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, Y).NPC(ElMapo).Index = -1 Then
    If Abs(Y - ALLMyNPC(ElMapo).MyNPC(Userindex).Y) > 1 Then
        MegaServer.GameChat.Text = MegaServer.GameChat.Text + vbCrLf + MapSock(ElMapo).Sockers(Userindex).Name & "Skipped a space"  'niner
        Warp Userindex, ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y, ElMapo
        Exit Sub
    End If
    SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).NPC(ElMapo).Index = -1
    ALLMyNPC(ElMapo).MyNPC(Userindex).Y = Y
    SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).NPC(ElMapo).Index = Userindex + ALLServNPC(ElMapo).NPCTotal + 1
    WSock2.SendDataToClients ElMapo, Userindex, "#N(" & Userindex + ALLServNPC(ElMapo).NPCTotal & ")Y" & ALLMyNPC(ElMapo).MyNPC(Userindex).Y
    If SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, Y).Script(ElMapo) > -1 Then
        ScriptMod.CurScript = SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, Y).Script(ElMapo)
        ScriptMod.PlayerSocket = Userindex
        ScriptMod.Execute
    End If
Else
    If SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, Y).TileProp(ElMapo) = 1 And SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, Y).NPC(ElMapo).Index = -1 Then
        If ALLMyNPC(ElMapo).MyNPC(Userindex).Equipment(3).Item = SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, Y).GItem(ElMapo) Then
            SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).NPC(ElMapo).Index = -1
            ALLMyNPC(ElMapo).MyNPC(Userindex).Y = Y
            SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).NPC(ElMapo).Index = Userindex + ALLServNPC(ElMapo).NPCTotal + 1
            WSock2.SendDataToClients ElMapo, Userindex, "#N(" & Userindex + ALLServNPC(ElMapo).NPCTotal & ")Y" & ALLMyNPC(ElMapo).MyNPC(Userindex).Y
            If SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, Y).Script(ElMapo) > -1 Then
                ScriptMod.CurScript = SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, Y).Script(ElMapo)
                ScriptMod.PlayerSocket = Userindex
                ScriptMod.Execute
            End If
            Exit Sub
        End If
    End If
    WSock2.SendDataToAClient ElMapo, Userindex, "#N(" & Userindex + ALLServNPC(ElMapo).NPCTotal & ")Y" & ALLMyNPC(ElMapo).MyNPC(Userindex).Y
End If
If Y = 0 Or Y = 49 Then ChangeMap Userindex, False, -1
End Sub

Private Sub Equip(Userindex As Integer, UserItemIndex As Integer)
Dim TempType As Integer
Dim ItemIndex As Integer
ItemIndex = ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(UserItemIndex).Index
If ItemIndex = -1 Then Exit Sub
TempType = MyObjects(ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(UserItemIndex).Index).ObjType
If TempType < 4 Or TempType > 5 Then
    If ALLMyNPC(ElMapo).MyNPC(Userindex).Equipment(TempType).BackPIndex > -1 Then
        ModNPCStat Userindex, ALLMyNPC(ElMapo).MyNPC(Userindex).Equipment(TempType).Item, False
        ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(ALLMyNPC(ElMapo).MyNPC(Userindex).Equipment(TempType).BackPIndex).Equipped = False
    End If
    If TempType = 9 Then ALLMyNPC(ElMapo).MyNPC(Userindex).Weap = (MyObjects(ItemIndex).GIndex + 1)
    ALLMyNPC(ElMapo).MyNPC(Userindex).Equipment(TempType).BackPIndex = UserItemIndex
    ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(UserItemIndex).Equipped = True
    ALLMyNPC(ElMapo).MyNPC(Userindex).Equipment(TempType).Item = ItemIndex
    ModNPCStat Userindex, ItemIndex, True
    If MyObjects(ItemIndex).Description <> "" Then
        WSock2.SendDataToAClient ElMapo, Userindex, "$You read the inscription on the " & MyObjects(ItemIndex).Name & " you equipped:;" & MyObjects(ItemIndex).Description
    Else
        WSock2.SendDataToAClient ElMapo, Userindex, "$You equipped the " & MyObjects(ItemIndex).Name
    End If
    If TempType = 9 Then SendDataToClients ElMapo, -1, "#N(" & (Userindex + ALLServNPC(ElMapo).NPCTotal) & ")W" & Str$(MyObjects(ItemIndex).GIndex + 1)
Else
    If TempType = 5 Then
        ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.HP = ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.HP + MyObjects(ItemIndex).Power
        If ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.HP > ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.MaxHP Then ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.HP = ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.MaxHP
        ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(UserItemIndex).Amount = ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(UserItemIndex).Amount - 1
        If ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(UserItemIndex).Amount = 0 Then ALLMyNPC(ElMapo).MyNPC(Userindex).BPack(UserItemIndex).Index = -1
        WSock2.SendDataToAClient ElMapo, Userindex, "$Your health is now " & ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.HP
    End If
End If
'redo all the scripts
RScriptMods Userindex, True
End Sub
Private Sub ModNPCStat(Userindex As Integer, ItemIndex As Integer, Oper As Boolean)
If Oper = True Then
'equipping
    Select Case MyObjects(ItemIndex).ObjType
    Case 0, 1, 2, 7, 8:
        ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.Arm = ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.Arm + MyObjects(ItemIndex).Power
    Case 6:
        'all scripted
    Case 9:
        ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.Str = ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.Str + MyObjects(ItemIndex).Power
    End Select
Else
'taking off
    Select Case MyObjects(ItemIndex).ObjType
    Case 0, 1, 2, 7, 8:
        ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.Arm = ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.Arm - MyObjects(ItemIndex).Power
    Case 6:
        'all scripted
    Case 9:
        ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.Str = ALLMyNPC(ElMapo).MyNPC(Userindex).Attribs.Str - MyObjects(ItemIndex).Power
    End Select
End If
End Sub
Public Sub RespawnNPC()
Dim i As Integer
Dim b As Integer
For b = 0 To UBound(SMapArray(0, 0).TileProp, 1)
    For i = 1 To UBound(ALLServNPC(b).ServNPC, 1)
        If ALLServNPC(b).ServNPC(i).ReSpawn And SMapArray(ALLServNPC(b).ServNPC(i).X, ALLServNPC(b).ServNPC(i).Y).NPC(b).Index = -1 Then
            SendDataToClients b, -1, "#N(" & i - 1 & ")B" & Str$(ALLServNPC(b).ServNPC(i).Body)
            ALLServNPC(b).ServNPC(i).X = ALLServNPC(b).ServNPC(i).RX
            ALLServNPC(b).ServNPC(i).Y = ALLServNPC(b).ServNPC(i).RY
            ALLServNPC(b).ServNPC(i).Attribs.HP = ALLServNPC(b).ServNPC(i).Attribs.MaxHP
            SendDataToClients b, -1, "#N(" & (i - 1) & ")w" & ALLServNPC(b).ServNPC(i).X & "," & ALLServNPC(b).ServNPC(i).Y & ","
            SMapArray(ALLServNPC(b).ServNPC(i).X, ALLServNPC(b).ServNPC(i).Y).NPC(b).Index = i
            ALLServNPC(b).ServNPC(i).ReSpawn = False
        End If
    Next i
Next b
'this part puts all the gitems back out
For b = 0 To UBound(MyGitemz, 1)
    For i = 0 To (UBound(MyGitemz(b).MrItem, 1) - 1)
        SMapArray(MyGitemz(b).MrX(i), MyGitemz(b).MrY(i)).GItem(b) = MyGitemz(b).MrItem(i)
        WSock2.SendDataToClients b, -1, "#MC(" & Str$(MyGitemz(b).MrX(i)) & "," & Str$(MyGitemz(b).MrY(i)) & ")" & Str$(MyObjects(MyGitemz(b).MrItem(i)).GIndex) & ","
    Next i
Next b
End Sub
Public Sub RScriptMods(ByRef PlayerSocket As Integer, ReRunS As Boolean)
Dim TempCount As Integer
ALLMyNPC(ElMapo).MyNPC(PlayerSocket).Attribs.Str = ALLMyNPC(ElMapo).MyNPC(PlayerSocket).Attribs.Str - ALLMyNPC(ElMapo).MyNPC(PlayerSocket).AttMods.Str
ALLMyNPC(ElMapo).MyNPC(PlayerSocket).Attribs.Arm = ALLMyNPC(ElMapo).MyNPC(PlayerSocket).Attribs.Arm - ALLMyNPC(ElMapo).MyNPC(PlayerSocket).AttMods.Arm
ALLMyNPC(ElMapo).MyNPC(PlayerSocket).Attribs.DSk = ALLMyNPC(ElMapo).MyNPC(PlayerSocket).Attribs.DSk - ALLMyNPC(ElMapo).MyNPC(PlayerSocket).AttMods.DSk
ALLMyNPC(ElMapo).MyNPC(PlayerSocket).Attribs.ASk = ALLMyNPC(ElMapo).MyNPC(PlayerSocket).Attribs.ASk - ALLMyNPC(ElMapo).MyNPC(PlayerSocket).AttMods.ASk
ALLMyNPC(ElMapo).MyNPC(PlayerSocket).Attribs.MaxHP = ALLMyNPC(ElMapo).MyNPC(PlayerSocket).Attribs.MaxHP - ALLMyNPC(ElMapo).MyNPC(PlayerSocket).AttMods.MaxHP
ALLMyNPC(ElMapo).MyNPC(PlayerSocket).AttMods.Str = 0
ALLMyNPC(ElMapo).MyNPC(PlayerSocket).AttMods.Arm = 0
ALLMyNPC(ElMapo).MyNPC(PlayerSocket).AttMods.DSk = 0
ALLMyNPC(ElMapo).MyNPC(PlayerSocket).AttMods.ASk = 0
ALLMyNPC(ElMapo).MyNPC(PlayerSocket).AttMods.MaxHP = 0
If ReRunS Then
    For TempCount = 0 To 9
        If ALLMyNPC(ElMapo).MyNPC(PlayerSocket).Equipment(TempCount).Item > -1 Then
            If MyObjects(ALLMyNPC(ElMapo).MyNPC(PlayerSocket).Equipment(TempCount).Item).SIndex > -1 Then
                ScriptMod.CurScript = MyObjects(ALLMyNPC(ElMapo).MyNPC(PlayerSocket).Equipment(TempCount).Item).SIndex
                ScriptMod.PlayerSocket = PlayerSocket
                ScriptMod.Execute
            End If
        End If
    Next
End If
End Sub

Public Sub Warp(i As Integer, NewX As Integer, NewY As Integer, NewM As Integer)
Dim intX As Integer
Dim intY As Integer
SMapArray(ALLMyNPC(ElMapo).MyNPC(i).X, ALLMyNPC(ElMapo).MyNPC(i).Y).NPC(ElMapo).Index = -1
'this is stealing other players index
    For intX = 0 To 15
        For intY = 0 To 15
            If (NewX + intX) < 49 And (NewY + intY) < 49 Then
                If SMapArray(NewX + intX, NewY + intY).NPC(NewM).Index = -1 Then
                    NewX = NewX + intX
                    NewY = NewY + intY
                    GoTo Exiter
                End If
            End If
            If (NewX - intX) > 0 And (NewY - intY) > 0 Then
                If SMapArray(NewX - intX, NewY - intY).NPC(NewM).Index = -1 Then
                    NewX = NewX - intX
                    NewY = NewY - intY
                    GoTo Exiter
                End If
            End If
        Next
    Next
Exiter:
ChangeMap i, True, NewM, NewX, NewY
End Sub
Public Sub ChangeMap(Userindex As Integer, Warpy As Boolean, WM As Integer, Optional WX As Integer, Optional WY As Integer)
Dim Loopi As Integer
Dim intXCounter As Byte
Dim intYCounter As Byte
Dim TempS As String
Dim Oldelmapo As Integer
Dim NewIndex As Integer
Dim i As Integer
SMapArray(ALLMyNPC(ElMapo).MyNPC(Userindex).X, ALLMyNPC(ElMapo).MyNPC(Userindex).Y).NPC(ElMapo).Index = -1
If ElMapo <> WM Then
    SendDataToClients ElMapo, Userindex, "#N(" & Userindex + ALLServNPC(ElMapo).NPCTotal & ")K"
    ALLMyNPC(ElMapo).MyNPC(Userindex).Active = False
    Oldelmapo = ElMapo
    If Warpy Then
        ElMapo = WM
    Else
        If ALLMyNPC(Oldelmapo).MyNPC(Userindex).X = 0 Or ALLMyNPC(Oldelmapo).MyNPC(Userindex).X = 49 Then
            Select Case ALLMyNPC(Oldelmapo).MyNPC(Userindex).X
            Case 0: ElMapo = (MyExits(Oldelmapo).WestE - 1)
            Case 49: ElMapo = (MyExits(Oldelmapo).EastE - 1)
            End Select
        ElseIf ALLMyNPC(Oldelmapo).MyNPC(Userindex).Y = 0 Or ALLMyNPC(Oldelmapo).MyNPC(Userindex).Y = 49 Then
            Select Case ALLMyNPC(Oldelmapo).MyNPC(Userindex).Y
            Case 0: ElMapo = (MyExits(Oldelmapo).NorthE - 1)
            Case 49: ElMapo = (MyExits(Oldelmapo).SouthE - 1)
            End Select
        End If
    End If
    If Oldelmapo <> ElMapo Then
        If MapSock(ElMapo).MaxCon > -1 Then
            For i = 0 To MapSock(ElMapo).MaxCon
                If MapSock(ElMapo).Sockers(i).Active = False Then
                    MapSock(ElMapo).Sockers(i).Socket = MapSock(Oldelmapo).Sockers(Userindex).Socket
                    MapSock(ElMapo).Sockers(i).Active = True
                    MapSock(ElMapo).Sockers(i).Index = i
                    NewIndex = i
                    MapSock(ElMapo).Sockers(i).IP = MapSock(Oldelmapo).Sockers(Userindex).IP
                    MapSock(ElMapo).Sockers(i).Name = MapSock(Oldelmapo).Sockers(Userindex).Name
                    GoTo Skippay:
                End If
            Next
        End If
        MapSock(ElMapo).MaxCon = MapSock(ElMapo).MaxCon + 1
        ReDim Preserve MapSock(ElMapo).Sockers(0 To MapSock(ElMapo).MaxCon)
        MapSock(ElMapo).Sockers(MapSock(ElMapo).MaxCon).Socket = MapSock(Oldelmapo).Sockers(Userindex).Socket
        MapSock(ElMapo).Sockers(MapSock(ElMapo).MaxCon).Active = True
        NewIndex = MapSock(ElMapo).MaxCon
        'MegaServer.GameChat.Text = MegaServer.GameChat.Text + vbCrLf + "created new index"
        ReDim Preserve ALLMyNPC(ElMapo).MyNPC(0 To NewIndex)
    Else
        NewIndex = Userindex
    End If
Skippay:
    'MsgBox "Elmapo:" & ElMapo & "userindex" & newindex
    MapSock(Oldelmapo).Sockers(Userindex).Active = False
    MapSock(ElMapo).Sockers(NewIndex).Active = True
    ALLMyNPC(ElMapo).MyNPC(NewIndex) = ALLMyNPC(Oldelmapo).MyNPC(Userindex)
    If Oldelmapo <> ElMapo Then
        ALLMyNPC(Oldelmapo).MyNPC(Userindex).Body = 6
    End If
    TempS = "#MN"
    If Warpy Then
        ALLMyNPC(ElMapo).MyNPC(NewIndex).X = WX
        ALLMyNPC(ElMapo).MyNPC(NewIndex).Y = WY
        TempS = TempS & (WM + 1) & ","
    Else
        If ALLMyNPC(Oldelmapo).MyNPC(Userindex).X = 0 Then ALLMyNPC(ElMapo).MyNPC(NewIndex).X = 49
        If ALLMyNPC(Oldelmapo).MyNPC(Userindex).X = 49 Then ALLMyNPC(ElMapo).MyNPC(NewIndex).X = 0
        If ALLMyNPC(Oldelmapo).MyNPC(Userindex).Y = 49 Then ALLMyNPC(ElMapo).MyNPC(NewIndex).Y = 0
        If ALLMyNPC(Oldelmapo).MyNPC(Userindex).Y = 0 Then ALLMyNPC(ElMapo).MyNPC(NewIndex).Y = 49
    End If
    If Warpy = False Then
        Select Case ALLMyNPC(Oldelmapo).MyNPC(Userindex).X
            Case 0: TempS = TempS & (MyExits(Oldelmapo).WestE) & ","
            Case 49: TempS = TempS & (MyExits(Oldelmapo).EastE) & ","
        End Select
         Select Case ALLMyNPC(Oldelmapo).MyNPC(Userindex).Y
            Case 0: TempS = TempS & (MyExits(Oldelmapo).NorthE) & ","
            Case 49: TempS = TempS & (MyExits(Oldelmapo).SouthE) & ","
        End Select
    End If
    If ALLServNPC(ElMapo).NPCTotal > 0 Then
        For Loopi = 1 To ALLServNPC(ElMapo).NPCTotal 'fix!!!1
            If ALLServNPC(ElMapo).ServNPC(Loopi).ReSpawn = False Then
                TempS = TempS & "(" & ALLServNPC(ElMapo).ServNPC(Loopi).X & "," & ALLServNPC(ElMapo).ServNPC(Loopi).Y & ")" & ALLServNPC(ElMapo).ServNPC(Loopi).Body & "," & ALLServNPC(ElMapo).ServNPC(Loopi).Head & ",0,"
            Else
                TempS = TempS & "(" & ALLServNPC(ElMapo).ServNPC(Loopi).X & "," & ALLServNPC(ElMapo).ServNPC(Loopi).Y & ")" & 5 & "," & ALLServNPC(ElMapo).ServNPC(Loopi).Head & ",0,"
            End If
        Next
    End If
    For Loopi = 0 To MapSock(ElMapo).MaxCon
        'If MapSock(ElMapo).Sockers(Loopi).Active Then
            TempS = TempS & "(" & ALLMyNPC(ElMapo).MyNPC(Loopi).X & "," & ALLMyNPC(ElMapo).MyNPC(Loopi).Y & ")" & ALLMyNPC(ElMapo).MyNPC(Loopi).Body & "," & ALLMyNPC(ElMapo).MyNPC(Loopi).Head & "," & ALLMyNPC(ElMapo).MyNPC(Loopi).Weap & ","
        'End If
    Next
    TempS = TempS & "/" & NewIndex + ALLServNPC(ElMapo).NPCTotal & "*"
    For intXCounter = 0 To 49
        For intYCounter = 0 To 49
            If SMapArray(intXCounter, intYCounter).GItem(ElMapo) > -1 And SMapArray(intXCounter, intYCounter).TileProp(ElMapo) <> 1 Then
                TempS = TempS & "(" & intXCounter & "," & intYCounter & ")" & MyObjects(SMapArray(intXCounter, intYCounter).GItem(ElMapo)).GIndex & ","
            End If
        Next
    Next
    WSock2.SendDataToAClient ElMapo, NewIndex, TempS
    If MapSock(ElMapo).MaxCon > 0 Then
        SendDataToClients ElMapo, NewIndex, "#N(" & NewIndex + ALLServNPC(ElMapo).NPCTotal & ")C" & ALLMyNPC(ElMapo).MyNPC(NewIndex).Body & "," & ALLMyNPC(ElMapo).MyNPC(NewIndex).Head & "," & ALLMyNPC(ElMapo).MyNPC(NewIndex).Weap & "," & ALLMyNPC(ElMapo).MyNPC(NewIndex).X & "," & ALLMyNPC(ElMapo).MyNPC(NewIndex).Y & ","
    End If
    'MegaServer.GameChat.Text = MegaServer.GameChat.Text + vbCrLf + "Elmapo:" & ElMapo & "userindex" & Str$(NewIndex) + vbCrLf + TempS + vbCrLf + Str$(UBound(ALLMyNPC(ElMapo).MyNPC, 1))
    ALLMyNPC(ElMapo).MyNPC(NewIndex).Active = True
Else
    NewIndex = Userindex
    ALLMyNPC(ElMapo).MyNPC(NewIndex).X = WX
    ALLMyNPC(ElMapo).MyNPC(NewIndex).Y = WY
    SendDataToClients ElMapo, -1, "#N(" & NewIndex + ALLServNPC(ElMapo).NPCTotal & ")w" & WX & "," & WY & ","
    'MegaServer.GameChat.Text = MegaServer.GameChat.Text + vbCrLf + "efficient"
End If
SMapArray(ALLMyNPC(ElMapo).MyNPC(NewIndex).X, ALLMyNPC(ElMapo).MyNPC(NewIndex).Y).NPC(ElMapo).Index = NewIndex + ALLServNPC(ElMapo).NPCTotal + 1
End Sub
