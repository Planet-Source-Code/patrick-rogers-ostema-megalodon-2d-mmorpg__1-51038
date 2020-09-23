Attribute VB_Name = "Combat"
Option Explicit
Private Function GetDist(intX1 As Integer, intY1 As Integer, intX2 As Integer, intY2 As Integer) As Integer
    GetDist = Sqr((intX1 - intX2) ^ 2 + (intY1 - intY2) ^ 2)
End Function
Public Sub AISeek(CurMap As Integer, NPCIndex As Integer, Attackable As Boolean)
Dim TempDist As Integer
Dim PlayerIndex As Integer
If ALLServNPC(CurMap).ServNPC(NPCIndex).Mobile = False Or ALLServNPC(CurMap).ServNPC(NPCIndex).ReSpawn Then Exit Sub
If ALLServNPC(CurMap).ServNPC(NPCIndex).Walking <> 0 Then Exit Sub
For PlayerIndex = 0 To UBound(ALLMyNPC(CurMap).MyNPC, 1)
    TempDist = GetDist(ALLServNPC(CurMap).ServNPC(NPCIndex).X, ALLServNPC(CurMap).ServNPC(NPCIndex).Y, ALLMyNPC(CurMap).MyNPC(PlayerIndex).X, ALLMyNPC(CurMap).MyNPC(PlayerIndex).Y)
    If TempDist <= 4 And ALLMyNPC(CurMap).MyNPC(PlayerIndex).Active Then
        If TempDist > 1 Then
ReDo:
            If ALLMyNPC(CurMap).MyNPC(PlayerIndex).X < ALLServNPC(CurMap).ServNPC(NPCIndex).X Then
                If SMapArray(ALLServNPC(CurMap).ServNPC(NPCIndex).X - 1, ALLServNPC(CurMap).ServNPC(NPCIndex).Y).TileProp(CurMap) > 2 And SMapArray(ALLServNPC(CurMap).ServNPC(NPCIndex).X - 1, ALLServNPC(CurMap).ServNPC(NPCIndex).Y).NPC(CurMap).Index = -1 Then
                    SMapArray(ALLServNPC(CurMap).ServNPC(NPCIndex).X, ALLServNPC(CurMap).ServNPC(NPCIndex).Y).NPC(CurMap).Index = -1
                    ALLServNPC(CurMap).ServNPC(NPCIndex).X = ALLServNPC(CurMap).ServNPC(NPCIndex).X - 1
                    SMapArray(ALLServNPC(CurMap).ServNPC(NPCIndex).X, ALLServNPC(CurMap).ServNPC(NPCIndex).Y).NPC(CurMap).Index = NPCIndex
                    WSock2.SendDataToClients CurMap, -1, "#N(" & NPCIndex - 1 & ")X" & ALLServNPC(CurMap).ServNPC(NPCIndex).X
                    Exit Sub
                End If
            End If
            If ALLMyNPC(CurMap).MyNPC(PlayerIndex).X > ALLServNPC(CurMap).ServNPC(NPCIndex).X Then
                If SMapArray(ALLServNPC(CurMap).ServNPC(NPCIndex).X + 1, ALLServNPC(CurMap).ServNPC(NPCIndex).Y).TileProp(CurMap) > 2 And SMapArray(ALLServNPC(CurMap).ServNPC(NPCIndex).X + 1, ALLServNPC(CurMap).ServNPC(NPCIndex).Y).NPC(CurMap).Index = -1 Then
                    SMapArray(ALLServNPC(CurMap).ServNPC(NPCIndex).X, ALLServNPC(CurMap).ServNPC(NPCIndex).Y).NPC(CurMap).Index = -1
                    ALLServNPC(CurMap).ServNPC(NPCIndex).X = ALLServNPC(CurMap).ServNPC(NPCIndex).X + 1
                    SMapArray(ALLServNPC(CurMap).ServNPC(NPCIndex).X, ALLServNPC(CurMap).ServNPC(NPCIndex).Y).NPC(CurMap).Index = NPCIndex
                    WSock2.SendDataToClients CurMap, -1, "#N(" & NPCIndex - 1 & ")X" & ALLServNPC(CurMap).ServNPC(NPCIndex).X
                    Exit Sub
                End If
            End If
            If ALLMyNPC(CurMap).MyNPC(PlayerIndex).Y < ALLServNPC(CurMap).ServNPC(NPCIndex).Y Then
                If SMapArray(ALLServNPC(CurMap).ServNPC(NPCIndex).X, ALLServNPC(CurMap).ServNPC(NPCIndex).Y - 1).TileProp(CurMap) > 2 And SMapArray(ALLServNPC(CurMap).ServNPC(NPCIndex).X, ALLServNPC(CurMap).ServNPC(NPCIndex).Y - 1).NPC(CurMap).Index = -1 Then
                    SMapArray(ALLServNPC(CurMap).ServNPC(NPCIndex).X, ALLServNPC(CurMap).ServNPC(NPCIndex).Y).NPC(CurMap).Index = -1
                    ALLServNPC(CurMap).ServNPC(NPCIndex).Y = ALLServNPC(CurMap).ServNPC(NPCIndex).Y - 1
                    SMapArray(ALLServNPC(CurMap).ServNPC(NPCIndex).X, ALLServNPC(CurMap).ServNPC(NPCIndex).Y).NPC(CurMap).Index = NPCIndex
                    WSock2.SendDataToClients CurMap, -1, "#N(" & NPCIndex - 1 & ")Y" & ALLServNPC(CurMap).ServNPC(NPCIndex).Y
                    Exit Sub
                End If
            End If
            If ALLMyNPC(CurMap).MyNPC(PlayerIndex).Y > ALLServNPC(CurMap).ServNPC(NPCIndex).Y Then
                If SMapArray(ALLServNPC(CurMap).ServNPC(NPCIndex).X, ALLServNPC(CurMap).ServNPC(NPCIndex).Y + 1).TileProp(CurMap) > 2 And SMapArray(ALLServNPC(CurMap).ServNPC(NPCIndex).X, ALLServNPC(CurMap).ServNPC(NPCIndex).Y + 1).NPC(CurMap).Index = -1 Then
                    SMapArray(ALLServNPC(CurMap).ServNPC(NPCIndex).X, ALLServNPC(CurMap).ServNPC(NPCIndex).Y).NPC(CurMap).Index = -1
                    ALLServNPC(CurMap).ServNPC(NPCIndex).Y = ALLServNPC(CurMap).ServNPC(NPCIndex).Y + 1
                    SMapArray(ALLServNPC(CurMap).ServNPC(NPCIndex).X, ALLServNPC(CurMap).ServNPC(NPCIndex).Y).NPC(CurMap).Index = NPCIndex
                    WSock2.SendDataToClients CurMap, -1, "#N(" & NPCIndex - 1 & ")Y" & ALLServNPC(CurMap).ServNPC(NPCIndex).Y
                    Exit Sub
                End If
            End If
        Else
            If ALLMyNPC(CurMap).MyNPC(PlayerIndex).X = ALLServNPC(CurMap).ServNPC(NPCIndex).X Or ALLMyNPC(CurMap).MyNPC(PlayerIndex).Y = ALLServNPC(CurMap).ServNPC(NPCIndex).Y Then
                If ALLMyNPC(CurMap).MyNPC(PlayerIndex).X = ALLServNPC(CurMap).ServNPC(NPCIndex).X And Attackable Then
                    If ALLMyNPC(CurMap).MyNPC(PlayerIndex).Y < ALLServNPC(CurMap).ServNPC(NPCIndex).Y Then
                        WSock2.SendDataToClients CurMap, -1, "#N(" & NPCIndex - 1 & ")F1"
                        ALLServNPC(CurMap).ServNPC(NPCIndex).Facing = 1
                        AIAttack CurMap, NPCIndex
                    Else
                        WSock2.SendDataToClients CurMap, -1, "#N(" & NPCIndex - 1 & ")F2"
                        ALLServNPC(CurMap).ServNPC(NPCIndex).Facing = 2
                        AIAttack CurMap, NPCIndex
                    End If
                Exit Sub
                End If
                If ALLMyNPC(CurMap).MyNPC(PlayerIndex).Y = ALLServNPC(CurMap).ServNPC(NPCIndex).Y And Attackable Then
                    If ALLMyNPC(CurMap).MyNPC(PlayerIndex).X < ALLServNPC(CurMap).ServNPC(NPCIndex).X Then
                        WSock2.SendDataToClients CurMap, -1, "#N(" & NPCIndex - 1 & ")F4"
                        ALLServNPC(CurMap).ServNPC(NPCIndex).Facing = 4
                        AIAttack CurMap, NPCIndex
                    Else
                        WSock2.SendDataToClients CurMap, -1, "#N(" & NPCIndex - 1 & ")F3"
                        ALLServNPC(CurMap).ServNPC(NPCIndex).Facing = 3
                        AIAttack CurMap, NPCIndex
                    End If
                Exit Sub
                End If
            Exit Sub
            Else
            GoTo ReDo
            End If
        End If
    End If
    Next PlayerIndex
    'PlayerIndex = PlayerIndex - 1
    MoveAllServNPC CurMap, NPCIndex, True
    'WSock2.SendDataToClients CurMap, -1, "$Fail: Dist:" & Str$(TempDist) & " Active: " & Str$(ALLMyNPC(CurMap).MyNPC(PlayerIndex).Active) & " Index: " & Str$(PlayerIndex) & " X:" & Str$(ALLMyNPC(CurMap).MyNPC(PlayerIndex).X) & " Y:" & Str$(ALLMyNPC(CurMap).MyNPC(PlayerIndex).Y)
End Sub
Public Sub ClientAttack(SeniorMap As Integer, Userindex As Integer)
Dim EnemyIndex As Integer
EnemyIndex = -1
If ALLMyNPC(SeniorMap).MyNPC(Userindex).X = 0 Or ALLMyNPC(SeniorMap).MyNPC(Userindex).X = 49 Or ALLMyNPC(SeniorMap).MyNPC(Userindex).Y = 0 Or ALLMyNPC(SeniorMap).MyNPC(Userindex).Y = 49 Then Exit Sub
Select Case ALLMyNPC(SeniorMap).MyNPC(Userindex).LastMove
Case 1:
    If SMapArray(ALLMyNPC(SeniorMap).MyNPC(Userindex).X, ALLMyNPC(SeniorMap).MyNPC(Userindex).Y - 1).NPC(SeniorMap).Index > -1 And SMapArray(ALLMyNPC(SeniorMap).MyNPC(Userindex).X, ALLMyNPC(SeniorMap).MyNPC(Userindex).Y - 1).NPC(SeniorMap).Index <= ALLServNPC(SeniorMap).NPCTotal Then
        EnemyIndex = SMapArray(ALLMyNPC(SeniorMap).MyNPC(Userindex).X, ALLMyNPC(SeniorMap).MyNPC(Userindex).Y - 1).NPC(SeniorMap).Index
        If ALLServNPC(SeniorMap).ServNPC(EnemyIndex).NPCT <> 2 Then Exit Sub
        ALLServNPC(SeniorMap).ServNPC(EnemyIndex).Attribs.HP = ALLServNPC(SeniorMap).ServNPC(EnemyIndex).Attribs.HP - GetAttackP(SeniorMap, Userindex, EnemyIndex)
    End If
Case 2:
    If SMapArray(ALLMyNPC(SeniorMap).MyNPC(Userindex).X, ALLMyNPC(SeniorMap).MyNPC(Userindex).Y + 1).NPC(SeniorMap).Index > -1 And SMapArray(ALLMyNPC(SeniorMap).MyNPC(Userindex).X, ALLMyNPC(SeniorMap).MyNPC(Userindex).Y + 1).NPC(SeniorMap).Index <= ALLServNPC(SeniorMap).NPCTotal Then
        EnemyIndex = SMapArray(ALLMyNPC(SeniorMap).MyNPC(Userindex).X, ALLMyNPC(SeniorMap).MyNPC(Userindex).Y + 1).NPC(SeniorMap).Index
        If ALLServNPC(SeniorMap).ServNPC(EnemyIndex).NPCT <> 2 Then Exit Sub
        ALLServNPC(SeniorMap).ServNPC(EnemyIndex).Attribs.HP = ALLServNPC(SeniorMap).ServNPC(EnemyIndex).Attribs.HP - GetAttackP(SeniorMap, Userindex, EnemyIndex)
    End If
Case 3:
    If SMapArray(ALLMyNPC(SeniorMap).MyNPC(Userindex).X + 1, ALLMyNPC(SeniorMap).MyNPC(Userindex).Y).NPC(SeniorMap).Index > -1 And SMapArray(ALLMyNPC(SeniorMap).MyNPC(Userindex).X + 1, ALLMyNPC(SeniorMap).MyNPC(Userindex).Y).NPC(SeniorMap).Index <= ALLServNPC(SeniorMap).NPCTotal Then
        EnemyIndex = SMapArray(ALLMyNPC(SeniorMap).MyNPC(Userindex).X + 1, ALLMyNPC(SeniorMap).MyNPC(Userindex).Y).NPC(SeniorMap).Index
        If ALLServNPC(SeniorMap).ServNPC(EnemyIndex).NPCT <> 2 Then Exit Sub
        ALLServNPC(SeniorMap).ServNPC(EnemyIndex).Attribs.HP = ALLServNPC(SeniorMap).ServNPC(EnemyIndex).Attribs.HP - GetAttackP(SeniorMap, Userindex, EnemyIndex)
    End If
Case 4:
    If SMapArray(ALLMyNPC(SeniorMap).MyNPC(Userindex).X - 1, ALLMyNPC(SeniorMap).MyNPC(Userindex).Y).NPC(SeniorMap).Index > -1 And SMapArray(ALLMyNPC(SeniorMap).MyNPC(Userindex).X - 1, ALLMyNPC(SeniorMap).MyNPC(Userindex).Y).NPC(SeniorMap).Index <= ALLServNPC(SeniorMap).NPCTotal Then
        EnemyIndex = SMapArray(ALLMyNPC(SeniorMap).MyNPC(Userindex).X - 1, ALLMyNPC(SeniorMap).MyNPC(Userindex).Y).NPC(SeniorMap).Index
        If ALLServNPC(SeniorMap).ServNPC(EnemyIndex).NPCT <> 2 Then Exit Sub
        ALLServNPC(SeniorMap).ServNPC(EnemyIndex).Attribs.HP = ALLServNPC(SeniorMap).ServNPC(EnemyIndex).Attribs.HP - GetAttackP(SeniorMap, Userindex, EnemyIndex)
    End If
End Select
If EnemyIndex > -1 Then
    WSock2.SendDataToClients SeniorMap, -1, "#N(" & (Userindex + ALLServNPC(SeniorMap).NPCTotal) & ")F" & ALLMyNPC(SeniorMap).MyNPC(Userindex).LastMove
    If ALLServNPC(SeniorMap).ServNPC(EnemyIndex).Attribs.HP <= 0 Then KillServNPC SeniorMap, EnemyIndex, Userindex
End If
End Sub
Private Sub KillServNPC(ZMap As Integer, NPCIndex As Integer, KillerIndex As Integer)
ALLServNPC(ZMap).ServNPC(NPCIndex).ReSpawn = True
SMapArray(ALLServNPC(ZMap).ServNPC(NPCIndex).X, ALLServNPC(ZMap).ServNPC(NPCIndex).Y).NPC(ZMap).Index = -1
SendDataToClients ZMap, -1, "#N(" & NPCIndex - 1 & ")B5"
If ALLServNPC(ZMap).ServNPC(NPCIndex).DeathScript > -1 Then
    ScriptMod.PlayerSocket = KillerIndex
    ScriptMod.CurScript = ALLServNPC(ZMap).ServNPC(NPCIndex).DeathScript
    ScriptMod.Execute
End If
If SMapArray(ALLServNPC(ZMap).ServNPC(NPCIndex).X, ALLServNPC(ZMap).ServNPC(NPCIndex).Y).TileProp(ZMap) > 1 Then
    Randomize
    SMapArray(ALLServNPC(ZMap).ServNPC(NPCIndex).X, ALLServNPC(ZMap).ServNPC(NPCIndex).Y).GItem(ZMap) = ALLServNPC(ZMap).ServNPC(NPCIndex).Attribs.DeadDrops(Int((2 * Rnd) + 0))
    WSock2.SendDataToClients ZMap, -1, "#MC(" & ALLServNPC(ZMap).ServNPC(NPCIndex).X & "," & ALLServNPC(ZMap).ServNPC(NPCIndex).Y & ")" & MyObjects(SMapArray(ALLServNPC(ZMap).ServNPC(NPCIndex).X, ALLServNPC(ZMap).ServNPC(NPCIndex).Y).GItem(ZMap)).GIndex & ","
End If
End Sub
Private Sub KillMyNPC(NPCIndex As Integer, ByVal MrMap As Integer)
ALLMyNPC(MrMap).MyNPC(NPCIndex).Attribs.HP = ALLMyNPC(MrMap).MyNPC(NPCIndex).Attribs.MaxHP * 0.75
WSock2.SendDataToAClient MrMap, NPCIndex, "$You dead foo!"
DataManage.ElMapo = MrMap
DataManage.Warp NPCIndex, 24, 25, 0
End Sub
Private Sub CheckXP(ByRef ZMap As Integer, ByRef Userindex As Integer)
If ALLMyNPC(ZMap).MyNPC(Userindex).Attribs.XP >= ALLMyNPC(ZMap).MyNPC(Userindex).NextXP Then
    'get REAL player atts
    DataManage.RScriptMods Userindex, False
    ALLMyNPC(ZMap).MyNPC(Userindex).Attribs.MaxHP = ALLMyNPC(ZMap).MyNPC(Userindex).Attribs.MaxHP + 3
    ALLMyNPC(ZMap).MyNPC(Userindex).Attribs.Str = ALLMyNPC(ZMap).MyNPC(Userindex).Attribs.Str + 1
    ALLMyNPC(ZMap).MyNPC(Userindex).Attribs.XP = 0
    ALLMyNPC(ZMap).MyNPC(Userindex).NextXP = ALLMyNPC(ZMap).MyNPC(Userindex).NextXP * 1.75
    ALLMyNPC(ZMap).MyNPC(Userindex).SkillP = ALLMyNPC(ZMap).MyNPC(Userindex).SkillP + 1
    'rerun the scripts
    DataManage.RScriptMods Userindex, True
    WSock2.SendDataToAClient ZMap, Userindex, "$You gained a level!  Make sure to use your Skill Points"
End If
End Sub
Private Function GetAttackP(ZMap As Integer, Userindex As Integer, EnemyIndex As Integer) As Integer
Dim Chance As Integer
Chance = Int(((ALLMyNPC(ZMap).MyNPC(Userindex).Attribs.ASk) * Rnd) + 1)
If Chance > (ALLServNPC(ZMap).ServNPC(EnemyIndex).Attribs.DSk / 2) Then
    GetAttackP = Int(((ALLMyNPC(ZMap).MyNPC(Userindex).Attribs.Str / 2 + 1) * Rnd) + (ALLMyNPC(ZMap).MyNPC(Userindex).Attribs.Str / 2))
    GetAttackP = GetAttackP - (ALLServNPC(ZMap).ServNPC(EnemyIndex).Attribs.Arm)
    If GetAttackP < 0 Then GetAttackP = 0
    If GetAttackP < ALLServNPC(ZMap).ServNPC(EnemyIndex).Attribs.HP Then
        ALLMyNPC(ZMap).MyNPC(Userindex).Attribs.XP = ALLMyNPC(ZMap).MyNPC(Userindex).Attribs.XP + ((GetAttackP / ALLServNPC(ZMap).ServNPC(EnemyIndex).Attribs.MaxHP) * ALLServNPC(ZMap).ServNPC(EnemyIndex).Attribs.XP)
    Else
        ALLMyNPC(ZMap).MyNPC(Userindex).Attribs.XP = ALLMyNPC(ZMap).MyNPC(Userindex).Attribs.XP + ((ALLServNPC(ZMap).ServNPC(EnemyIndex).Attribs.HP / ALLServNPC(ZMap).ServNPC(EnemyIndex).Attribs.MaxHP) * ALLServNPC(ZMap).ServNPC(EnemyIndex).Attribs.XP)
    End If
    CheckXP ZMap, Userindex
    WSock2.SendDataToAClient ZMap, Userindex, "$You hit " & ALLServNPC(ZMap).ServNPC(EnemyIndex).Name & " for " & Str(GetAttackP) & "!"
Else
    GetAttackP = 0
    WSock2.SendDataToAClient ZMap, Userindex, "$You missed " & ALLServNPC(ZMap).ServNPC(EnemyIndex).Name & "!"
End If
End Function
Private Function GetAIAttackP(ZMap As Integer, Userindex As Integer, EnemyIndex As Integer) As Integer
Dim Chance As Integer
Chance = Int(((ALLServNPC(ZMap).ServNPC(Userindex).Attribs.ASk) * Rnd) + 1)
If Chance > (ALLMyNPC(ZMap).MyNPC(EnemyIndex).Attribs.DSk / 2) Then
    GetAIAttackP = Int(((ALLServNPC(ZMap).ServNPC(Userindex).Attribs.Str / 2 + 1) * Rnd) + (ALLServNPC(ZMap).ServNPC(Userindex).Attribs.Str / 2))
    GetAIAttackP = GetAIAttackP - (ALLMyNPC(ZMap).MyNPC(EnemyIndex).Attribs.Arm)
    If GetAIAttackP < 0 Then GetAIAttackP = 0
    WSock2.SendDataToAClient ZMap, EnemyIndex, "$" & ALLServNPC(ZMap).ServNPC(Userindex).Name & " hit you for " & Str(GetAIAttackP) & "!"
Else
    GetAIAttackP = 0
    WSock2.SendDataToAClient ZMap, EnemyIndex, "$" & ALLServNPC(ZMap).ServNPC(Userindex).Name & " missed you!"
End If
End Function
Private Function AIAttack(MrMap As Integer, NPCIndex As Integer)
Dim EnemyIndex As Integer
Select Case ALLServNPC(MrMap).ServNPC(NPCIndex).Facing
Case 1:
    If SMapArray(ALLServNPC(MrMap).ServNPC(NPCIndex).X, ALLServNPC(MrMap).ServNPC(NPCIndex).Y - 1).NPC(MrMap).Index > -1 And SMapArray(ALLServNPC(MrMap).ServNPC(NPCIndex).X, ALLServNPC(MrMap).ServNPC(NPCIndex).Y - 1).NPC(MrMap).Index > ALLServNPC(MrMap).NPCTotal Then
        EnemyIndex = (SMapArray(ALLServNPC(MrMap).ServNPC(NPCIndex).X, ALLServNPC(MrMap).ServNPC(NPCIndex).Y - 1).NPC(MrMap).Index) - ALLServNPC(MrMap).NPCTotal - 1
        ALLMyNPC(MrMap).MyNPC(EnemyIndex).Attribs.HP = ALLMyNPC(MrMap).MyNPC(EnemyIndex).Attribs.HP - GetAIAttackP(MrMap, NPCIndex, EnemyIndex)
    End If
Case 2:
    If SMapArray(ALLServNPC(MrMap).ServNPC(NPCIndex).X, ALLServNPC(MrMap).ServNPC(NPCIndex).Y + 1).NPC(MrMap).Index > -1 And SMapArray(ALLServNPC(MrMap).ServNPC(NPCIndex).X, ALLServNPC(MrMap).ServNPC(NPCIndex).Y + 1).NPC(MrMap).Index > ALLServNPC(MrMap).NPCTotal Then
        EnemyIndex = (SMapArray(ALLServNPC(MrMap).ServNPC(NPCIndex).X, ALLServNPC(MrMap).ServNPC(NPCIndex).Y + 1).NPC(MrMap).Index) - ALLServNPC(MrMap).NPCTotal - 1
        ALLMyNPC(MrMap).MyNPC(EnemyIndex).Attribs.HP = ALLMyNPC(MrMap).MyNPC(EnemyIndex).Attribs.HP - GetAIAttackP(MrMap, NPCIndex, EnemyIndex)
    End If
Case 3:
    If SMapArray(ALLServNPC(MrMap).ServNPC(NPCIndex).X + 1, ALLServNPC(MrMap).ServNPC(NPCIndex).Y).NPC(MrMap).Index > -1 And SMapArray(ALLServNPC(MrMap).ServNPC(NPCIndex).X + 1, ALLServNPC(MrMap).ServNPC(NPCIndex).Y).NPC(MrMap).Index > ALLServNPC(MrMap).NPCTotal Then
        EnemyIndex = (SMapArray(ALLServNPC(MrMap).ServNPC(NPCIndex).X + 1, ALLServNPC(MrMap).ServNPC(NPCIndex).Y).NPC(MrMap).Index) - ALLServNPC(MrMap).NPCTotal - 1
        ALLMyNPC(MrMap).MyNPC(EnemyIndex).Attribs.HP = ALLMyNPC(MrMap).MyNPC(EnemyIndex).Attribs.HP - GetAIAttackP(MrMap, NPCIndex, EnemyIndex)
    End If
Case 4:
    If SMapArray(ALLServNPC(MrMap).ServNPC(NPCIndex).X - 1, ALLServNPC(MrMap).ServNPC(NPCIndex).Y).NPC(MrMap).Index > -1 And SMapArray(ALLServNPC(MrMap).ServNPC(NPCIndex).X - 1, ALLServNPC(MrMap).ServNPC(NPCIndex).Y).NPC(MrMap).Index > ALLServNPC(MrMap).NPCTotal Then
        EnemyIndex = (SMapArray(ALLServNPC(MrMap).ServNPC(NPCIndex).X - 1, ALLServNPC(MrMap).ServNPC(NPCIndex).Y).NPC(MrMap).Index) - ALLServNPC(MrMap).NPCTotal - 1
        ALLMyNPC(MrMap).MyNPC(EnemyIndex).Attribs.HP = ALLMyNPC(MrMap).MyNPC(EnemyIndex).Attribs.HP - GetAIAttackP(MrMap, NPCIndex, EnemyIndex)
    End If
End Select
If EnemyIndex > -1 Then
    If ALLMyNPC(MrMap).MyNPC(EnemyIndex).Attribs.HP <= 0 Then KillMyNPC EnemyIndex, MrMap
End If
End Function
