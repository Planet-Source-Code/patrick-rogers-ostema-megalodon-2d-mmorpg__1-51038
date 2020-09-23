Attribute VB_Name = "ScriptMod"
'Megalodon Entertainment http://home.kc.rr.com/megalodonsoft
'Created by Patrick Rogers-Ostema
Option Explicit
Private Enum OP_CODE
    OPadd = 1
    OPsub
    OPmul
    OPdiv
    OPass 'assign:)
    OPint
    OPnum
    OPpri
    OPstr
    OPcstr
    OPinp
    OPiff
    OPthn 'then
    OPndi 'end if
    OPlss 'less than
    OPgrt 'greater than
    OPfor
    OPnex
    OPpvar 'player variable
    OPelse
    OPsend
End Enum
Private Type Node
    OP As OP_CODE
    Pointer As Long
End Type
Private Type Script
    MyNodes() As Node
    Name As String
    ConstStr() As String
End Type
Private MyScripts() As Script
Private Ints() As Integer
Private Strs() As String
Public CurScript As Integer
Public PlayerSocket As Integer
Dim TempOP As Byte
Dim TempOP2 As Integer
Public Sub Execute()
Dim i As Integer
Dim TempIndex As Integer
Dim BoolPVar As Boolean
Dim TempOP3() As Integer 'loop
Dim TempOP3Var() As Integer
ReDim TempOP3Var(0 To 0)
ReDim TempOP3(0 To 0)
ReDim Ints(0 To 0)
ReDim Strs(0 To 0)
TempIndex = -1
For i = 0 To UBound(MyScripts(CurScript).MyNodes, 1)
    If TempOP2 < 0 Then
        If MyScripts(CurScript).MyNodes(i).OP = OPndi Or MyScripts(CurScript).MyNodes(i).OP = OPelse Then
            TempOP2 = TempOP2 + 1
        End If
        If MyScripts(CurScript).MyNodes(i).OP = OPiff Then
            TempOP2 = TempOP2 - 1
        End If
    Else
        If MyScripts(CurScript).MyNodes(i).OP = OPelse Then
            TempOP2 = TempOP2 - 1
        End If
    End If
If TempOP2 > -1 Then
'ScriptForm.OutBox.Text = ScriptForm.OutBox.Text & "tempop" & Str(TempOP2) & vbCrLf
    Select Case MyScripts(CurScript).MyNodes(i).OP
    Case OP_CODE.OPint:
        If UBound(Ints, 1) < MyScripts(CurScript).MyNodes(i).Pointer + 1 Then
            ReDim Preserve Ints(0 To MyScripts(CurScript).MyNodes(i).Pointer + 1)
            TempIndex = -1
        Else
            If TempIndex > -1 Then
                If MyScripts(CurScript).MyNodes(TempIndex - 1).OP = OPstr Then
                    If TempOP = OP_CODE.OPadd Then
                        Strs(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Strs(MyScripts(CurScript).MyNodes(TempIndex).Pointer) & Str$(Ints(MyScripts(CurScript).MyNodes(i).Pointer))
                        TempOP = 0
                    ElseIf TempOP = OP_CODE.OPass Then
                        Strs(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Str$(Ints(MyScripts(CurScript).MyNodes(i).Pointer))
                        TempOP = 0
                    End If
                End If
            End If
            If TempIndex > -1 And TempOP = OP_CODE.OPass Then
                If TempOP2 = 0 Or BoolPVar Then
                    If BoolPVar Then
                        SetPVar MyScripts(CurScript).MyNodes(TempIndex).Pointer, (Ints(MyScripts(CurScript).MyNodes(i).Pointer)), i
                        BoolPVar = False
                    Else
                        Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(i).Pointer)
                    End If
                ElseIf TempOP2 = OP_CODE.OPiff Then
                    If Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) <> Ints(MyScripts(CurScript).MyNodes(i).Pointer) Then
                        TempOP2 = -1
                    Else
                        TempOP2 = 0
                    End If
                End If
            End If
            If TempOP = OP_CODE.OPlss Then
                If Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) > Ints(MyScripts(CurScript).MyNodes(i).Pointer) - 1 Then TempOP2 = -1
            End If
            If TempOP = OP_CODE.OPgrt Then
                If Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) < Ints(MyScripts(CurScript).MyNodes(i).Pointer) + 1 Then TempOP2 = -1
            End If
            If TempOP = OP_CODE.OPadd Then
                If CheckOrd(i + 1) = False Then
                    If MyScripts(CurScript).MyNodes(i).OP = OPnum Then
                        Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) + MyScripts(CurScript).MyNodes(i).Pointer
                    Else
                        Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) + Ints(MyScripts(CurScript).MyNodes(i).Pointer)
                    End If
                Else
                    i = i + 1
                    Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) + DoOrd(i)
                End If
            End If
            If TempOP = OP_CODE.OPsub Then
                If CheckOrd(i + 1) = False Then
                    If MyScripts(CurScript).MyNodes(i).OP = OPnum Then
                        Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) - MyScripts(CurScript).MyNodes(i).Pointer
                    Else
                        Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) - Ints(MyScripts(CurScript).MyNodes(i).Pointer)
                    End If
                Else
                    i = i + 1
                    Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) - DoOrd(i)
                End If
            End If
            If TempOP = OP_CODE.OPdiv Then
                Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) / Ints(MyScripts(CurScript).MyNodes(i).Pointer)
            End If
            If TempOP = OP_CODE.OPmul Then
                Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) * Ints(MyScripts(CurScript).MyNodes(i).Pointer)
            End If
            If TempOP = OP_CODE.OPpri Then
                WSock2.SendDataToAClient DataManage.ElMapo, PlayerSocket, "$" & Ints(MyScripts(CurScript).MyNodes(i).Pointer)
                'ScriptForm.OutBox.Text = ScriptForm.OutBox.Text & Ints(MyScripts(CurScript).MyNodes(i).Pointer)
            End If
            If TempOP = OP_CODE.OPinp Then
                'Ints(MyScripts(CurScript).MyNodes(i).Pointer) = InputBox("Enter integer")
            End If
            TempOP = 0
        End If
    Case OP_CODE.OPstr:
        If UBound(Strs, 1) < MyScripts(CurScript).MyNodes(i).Pointer + 1 Then
            ReDim Preserve Strs(0 To MyScripts(CurScript).MyNodes(i).Pointer + 1)
            TempIndex = -1
        Else
            If TempIndex > -1 And TempOP = OP_CODE.OPass Then
                If TempOP2 = 0 Then
                    Strs(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Strs(MyScripts(CurScript).MyNodes(i).Pointer)
                ElseIf TempOP2 = OP_CODE.OPiff Then
                    If Strs(MyScripts(CurScript).MyNodes(TempIndex).Pointer) <> Strs(MyScripts(CurScript).MyNodes(i).Pointer) Then
                        TempOP2 = -1
                    Else
                        TempOP2 = 0
                    End If
                End If
            End If
            If TempOP = OP_CODE.OPpri Then
                WSock2.SendDataToAClient DataManage.ElMapo, PlayerSocket, "$" & Strs(MyScripts(CurScript).MyNodes(i).Pointer)
                'ScriptForm.OutBox.Text = ScriptForm.OutBox.Text & Strs(MyScripts(CurScript).MyNodes(i).Pointer)
            End If
            If TempOP = OP_CODE.OPsend Then
                WSock2.SendDataToAClient DataManage.ElMapo, PlayerSocket, "#N(" & PlayerSocket + ALLServNPC(DataManage.ElMapo).NPCTotal & ")" & Strs(MyScripts(CurScript).MyNodes(i).Pointer)
                'ScriptForm.OutBox.Text = ScriptForm.OutBox.Text & Strs(MyScripts(CurScript).MyNodes(i).Pointer)
            End If
            If TempOP = OP_CODE.OPinp Then
                'Strs(MyScripts(CurScript).MyNodes(i).Pointer) = InputBox("Enter string")
            End If
            If TempOP = OP_CODE.OPadd Then
                Strs(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Strs(MyScripts(CurScript).MyNodes(TempIndex).Pointer) & Strs(MyScripts(CurScript).MyNodes(i).Pointer)
            End If
            TempOP = 0
        End If
    Case OP_CODE.OPcstr:
            If TempIndex > -1 And TempOP = OP_CODE.OPass Then
                If TempOP2 = 0 Then
                    Strs(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = MyScripts(CurScript).ConstStr(MyScripts(CurScript).MyNodes(i).Pointer)
                ElseIf TempOP2 = OP_CODE.OPiff Then
                    If Strs(MyScripts(CurScript).MyNodes(TempIndex).Pointer) <> MyScripts(CurScript).ConstStr(MyScripts(CurScript).MyNodes(i).Pointer) Then
                        TempOP2 = -1
                    Else
                        TempOP2 = 0
                    End If
                End If
            End If
            If TempOP = OP_CODE.OPadd Then
                Strs(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Strs(MyScripts(CurScript).MyNodes(TempIndex).Pointer) & MyScripts(CurScript).ConstStr(MyScripts(CurScript).MyNodes(i).Pointer)
            End If
            If TempOP = OP_CODE.OPpri Then
                WSock2.SendDataToAClient DataManage.ElMapo, PlayerSocket, "$" & MyScripts(CurScript).ConstStr(MyScripts(CurScript).MyNodes(i).Pointer)
                'ScriptForm.OutBox.Text = ScriptForm.OutBox.Text & MyScripts(CurScript).ConstStr(MyScripts(CurScript).MyNodes(i).Pointer)
            End If
            If TempOP = OP_CODE.OPsend Then
                WSock2.SendDataToAClient DataManage.ElMapo, PlayerSocket, "#N(" & PlayerSocket + ALLServNPC(DataManage.ElMapo).NPCTotal & ")" & MyScripts(CurScript).ConstStr(MyScripts(CurScript).MyNodes(i).Pointer)
                'ScriptForm.OutBox.Text = ScriptForm.OutBox.Text & MyScripts(CurScript).ConstStr(MyScripts(CurScript).MyNodes(i).Pointer)
            End If
            TempOP = 0
    Case OP_CODE.OPass:
        TempIndex = i
        TempOP = MyScripts(CurScript).MyNodes(i).OP
    Case OP_CODE.OPnum:
        If TempOP = OP_CODE.OPpri Then
            'ScriptForm.OutBox.Text = ScriptForm.OutBox.Text & MyScripts(CurScript).MyNodes(i).Pointer
        End If
        If TempIndex > -1 And TempOP = OP_CODE.OPass Then
            If TempOP2 = 0 Or BoolPVar Then
                If BoolPVar Then
                    SetPVar MyScripts(CurScript).MyNodes(TempIndex).Pointer, CInt(MyScripts(CurScript).MyNodes(i).Pointer), i
                    BoolPVar = False
                Else
                    Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = MyScripts(CurScript).MyNodes(i).Pointer
                End If
            ElseIf TempOP2 = OP_CODE.OPiff Then
                If Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) <> MyScripts(CurScript).MyNodes(i).Pointer Then
                    TempOP2 = -1
                Else
                    TempOP2 = 0
                End If
            End If
        End If
        If TempOP = OP_CODE.OPlss Then
            If Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) > MyScripts(CurScript).MyNodes(i).Pointer - 1 Then TempOP2 = -1
        End If
        If TempOP = OP_CODE.OPgrt Then
            If Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) < MyScripts(CurScript).MyNodes(i).Pointer + 1 Then TempOP2 = -1
        End If
        If TempOP = OP_CODE.OPadd Then
            If CheckOrd(i + 1) = False Then
                If MyScripts(CurScript).MyNodes(i).OP = OPnum Then
                    Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) + MyScripts(CurScript).MyNodes(i).Pointer
                Else
                    Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) + Ints(MyScripts(CurScript).MyNodes(i).Pointer)
                End If
            Else
                i = i + 1
                Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) + DoOrd(i)
            End If
        End If
        If TempOP = OP_CODE.OPsub Then
            If CheckOrd(i + 1) = False Then
                If MyScripts(CurScript).MyNodes(i).OP = OPnum Then
                    Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) - MyScripts(CurScript).MyNodes(i).Pointer
                Else
                    Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) - Ints(MyScripts(CurScript).MyNodes(i).Pointer)
                End If
            Else
                i = i + 1
                Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) - DoOrd(i)
            End If
        End If
        If TempOP = OP_CODE.OPmul Then
            Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) * MyScripts(CurScript).MyNodes(i).Pointer
        End If
        If TempOP = OP_CODE.OPdiv Then
            Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) / MyScripts(CurScript).MyNodes(i).Pointer
        End If
        TempOP = 0
    Case OP_CODE.OPpri: TempOP = OP_CODE.OPpri
    Case OP_CODE.OPsend: TempOP = OP_CODE.OPsend
    Case OP_CODE.OPadd: TempOP = OP_CODE.OPadd
    Case OP_CODE.OPsub: TempOP = OP_CODE.OPsub
    Case OP_CODE.OPmul: TempOP = OP_CODE.OPmul
    Case OP_CODE.OPdiv: TempOP = OP_CODE.OPdiv
    Case OP_CODE.OPinp: TempOP = OP_CODE.OPinp
    Case OP_CODE.OPfor:
    ReDim Preserve TempOP3(0 To UBound(TempOP3, 1) + 1)
    ReDim Preserve TempOP3Var(0 To UBound(TempOP3Var, 1) + 1)
    TempOP3Var(UBound(TempOP3Var, 1)) = MyScripts(CurScript).MyNodes(i).Pointer
    If MyScripts(CurScript).MyNodes(i + 1).OP = 7 Then
        Ints(TempOP3Var(UBound(TempOP3Var, 1))) = MyScripts(CurScript).MyNodes(i + 1).Pointer
    ElseIf MyScripts(CurScript).MyNodes(i + 1).OP = 6 Then
        Ints(TempOP3Var(UBound(TempOP3Var, 1))) = Ints(MyScripts(CurScript).MyNodes(i + 1).Pointer)
    End If
    If MyScripts(CurScript).MyNodes(i + 2).OP = 7 Then
        TempOP3(UBound(TempOP3, 1)) = MyScripts(CurScript).MyNodes(i + 2).Pointer
    ElseIf MyScripts(CurScript).MyNodes(i + 2).OP = 6 Then
        TempOP3(UBound(TempOP3, 1)) = Ints(MyScripts(CurScript).MyNodes(i + 2).Pointer)
    End If
    i = i + 2
    Case OP_CODE.OPnex:
    Ints(TempOP3Var(UBound(TempOP3Var, 1))) = Ints(TempOP3Var(UBound(TempOP3Var, 1))) + 1
    If TempOP3(UBound(TempOP3, 1)) >= Ints(TempOP3Var(UBound(TempOP3Var, 1))) Then
        i = (MyScripts(CurScript).MyNodes(i).Pointer - 1)
    Else
        ReDim Preserve TempOP3(0 To UBound(TempOP3, 1) - 1)
        ReDim Preserve TempOP3Var(0 To UBound(TempOP3Var, 1) - 1)
    End If
    Case OP_CODE.OPlss:
    TempOP = OP_CODE.OPlss
    TempIndex = i
    Case OP_CODE.OPgrt:
    TempOP = OP_CODE.OPgrt
    TempIndex = i
    Case OP_CODE.OPiff: TempOP2 = OP_CODE.OPiff
    Case OP_CODE.OPthn: If TempOP2 = OP_CODE.OPiff Then TempOP2 = OP_CODE.OPthn
    Case OP_CODE.OPndi: TempOP2 = 0
    Case OP_CODE.OPpvar:
    If TempOP = 0 Then BoolPVar = True
    If TempIndex > -1 Then
        If MyScripts(CurScript).MyNodes(TempIndex - 1).OP = OPstr Then
            If TempOP = OP_CODE.OPadd Then
                Strs(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Strs(MyScripts(CurScript).MyNodes(TempIndex).Pointer) & Str$(GetPVar(MyScripts(CurScript).MyNodes(i).Pointer))
                TempOP = 0
            ElseIf TempOP = OP_CODE.OPass Then
                Strs(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Str$(GetPVar(MyScripts(CurScript).MyNodes(i).Pointer))
                TempOP = 0
            End If
        End If
    End If
    If TempIndex > -1 And TempOP = OP_CODE.OPass Then
        If TempOP2 = 0 Then
            Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = GetPVar(MyScripts(CurScript).MyNodes(i).Pointer)
        ElseIf TempOP2 = OP_CODE.OPiff Then
            If Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) <> GetPVar(MyScripts(CurScript).MyNodes(i).Pointer) Then
                TempOP2 = -1
            Else
                TempOP2 = 0
            End If
        End If
    End If
    If TempOP = OP_CODE.OPpri Then WSock2.SendDataToAClient DataManage.ElMapo, PlayerSocket, "$" & GetPVar(MyScripts(CurScript).MyNodes(i).Pointer)
    If TempOP = OP_CODE.OPadd Then
        If CheckOrd(i + 1) = False Then
            If MyScripts(CurScript).MyNodes(i).OP = OPnum Then
                Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) + MyScripts(CurScript).MyNodes(i).Pointer
            ElseIf MyScripts(CurScript).MyNodes(i).OP = OPint Then
                Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) + Ints(MyScripts(CurScript).MyNodes(i).Pointer)
            ElseIf MyScripts(CurScript).MyNodes(i).OP = OPpvar Then
                Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) + GetPVar(MyScripts(CurScript).MyNodes(i).Pointer)
            End If
        Else
            i = i + 1
            Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) + DoOrd(i)
        End If
    End If
    If TempOP = OP_CODE.OPsub Then
        If CheckOrd(i + 1) = False Then
            If MyScripts(CurScript).MyNodes(i).OP = OPnum Then
                Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) - MyScripts(CurScript).MyNodes(i).Pointer
            ElseIf MyScripts(CurScript).MyNodes(i).OP = OPint Then
                Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) - Ints(MyScripts(CurScript).MyNodes(i).Pointer)
            ElseIf MyScripts(CurScript).MyNodes(i).OP = OPpvar Then
                Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) - GetPVar(MyScripts(CurScript).MyNodes(i).Pointer)
            End If
        Else
            i = i + 1
            Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) - DoOrd(i)
        End If
    End If
    If TempOP = OP_CODE.OPmul Then
        Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) * GetPVar(MyScripts(CurScript).MyNodes(i).Pointer)
    End If
    If TempOP = OP_CODE.OPdiv Then
        Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) = Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) / GetPVar(MyScripts(CurScript).MyNodes(i).Pointer)
    End If
    If TempOP = OP_CODE.OPlss Then
        If Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) > GetPVar(MyScripts(CurScript).MyNodes(i).Pointer) - 1 Then TempOP2 = -1
    End If
    If TempOP = OP_CODE.OPgrt Then
        If Ints(MyScripts(CurScript).MyNodes(TempIndex).Pointer) < GetPVar(MyScripts(CurScript).MyNodes(i).Pointer) + 1 Then TempOP2 = -1
    End If
    TempOP = 0
    End Select
End If
Skip2Next:
Next i
End Sub
Private Function GetPVar(zPointer As Long) As Integer
Select Case zPointer
Case 1: GetPVar = ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Money
Case 2: GetPVar = ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.HP
Case 3: GetPVar = ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.Str
Case 4: GetPVar = ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.Arm
Case 5: GetPVar = ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.DSk
Case 6: GetPVar = ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.ASk
Case 7: GetPVar = ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.MaxHP
End Select
End Function
Private Sub SetPVar(zPointer As Long, NewValue As Integer, ByRef i As Integer)
Select Case zPointer
Case 1: ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Money = NewValue
Case 2:
'ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).AttMods.HP = NewValue - ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.HP
If NewValue <= ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.HP Then
    ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.HP = NewValue
Else
    ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.HP = ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.MaxHP
End If
Case 3:
ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).AttMods.Str = NewValue - ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.Str
ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.Str = NewValue
Case 4:
ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).AttMods.Arm = NewValue - ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.Arm
ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.Arm = NewValue
Case 5:
ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).AttMods.DSk = NewValue - ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.DSk
ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.DSk = NewValue
Case 6:
ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).AttMods.ASk = NewValue - ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.ASk
ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.ASk = NewValue
Case 7:
ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).AttMods.MaxHP = NewValue - ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.MaxHP
ALLMyNPC(DataManage.ElMapo).MyNPC(PlayerSocket).Attribs.MaxHP = NewValue
Case 8:
If DataManage.CheckInvent(PlayerSocket, NewValue) = True And TempOP2 = 12 Then
    TempOP2 = 0
Else
    TempOP2 = -1
End If
Case 9:
DataManage.AddItemToInvent PlayerSocket, NewValue
Case 10:
DataManage.TakeInventItem PlayerSocket, NewValue
Case 11:
DataManage.Warp PlayerSocket, CInt(MyScripts(CurScript).MyNodes(i + 1).Pointer), CInt(MyScripts(CurScript).MyNodes(i + 2).Pointer), NewValue
i = i + 2
End Select
End Sub
Private Function CheckOrd(Index As Integer) As Boolean
CheckOrd = False
If MyScripts(CurScript).MyNodes(Index).OP = OPmul Or MyScripts(CurScript).MyNodes(Index).OP = OPdiv Then CheckOrd = True
End Function
Private Function DoOrd(ByRef Index As Integer) As Long
'this function does order of Z OPS
Dim Flago As Boolean
Flago = False
Reset:
If Flago = False Then
    If MyScripts(CurScript).MyNodes(Index - 1).OP = OPnum Then
        DoOrd = MyScripts(CurScript).MyNodes(Index - 1).Pointer
    ElseIf MyScripts(CurScript).MyNodes(Index - 1).OP = OPint Then
        DoOrd = Ints(MyScripts(CurScript).MyNodes(Index - 1).Pointer)
    ElseIf MyScripts(CurScript).MyNodes(Index - 1).OP = OPpvar Then
        DoOrd = GetPVar(MyScripts(CurScript).MyNodes(Index - 1).Pointer)
    End If
End If
Select Case MyScripts(CurScript).MyNodes(Index).OP
    Case OP_CODE.OPmul:
    If MyScripts(CurScript).MyNodes(Index + 1).OP = OPnum Then
        DoOrd = DoOrd * MyScripts(CurScript).MyNodes(Index + 1).Pointer
    ElseIf MyScripts(CurScript).MyNodes(Index + 1).OP = OPint Then
        DoOrd = DoOrd * Ints(MyScripts(CurScript).MyNodes(Index + 1).Pointer)
    ElseIf MyScripts(CurScript).MyNodes(Index + 1).OP = OPpvar Then
        DoOrd = DoOrd * GetPVar(MyScripts(CurScript).MyNodes(Index + 1).Pointer)
    End If
    Case OP_CODE.OPdiv:
    If MyScripts(CurScript).MyNodes(Index + 1).OP = OPnum Then
        DoOrd = DoOrd / MyScripts(CurScript).MyNodes(Index + 1).Pointer
    ElseIf MyScripts(CurScript).MyNodes(Index + 1).OP = OPint Then
        DoOrd = DoOrd / Ints(MyScripts(CurScript).MyNodes(Index + 1).Pointer)
    ElseIf MyScripts(CurScript).MyNodes(Index + 1).OP = OPpvar Then
        DoOrd = DoOrd / GetPVar(MyScripts(CurScript).MyNodes(Index + 1).Pointer)
    End If
End Select
If CheckOrd(Index + 2) Then
    Index = Index + 2
    Flago = True
    GoTo Reset
End If
End Function
Public Sub OpenScripts()
Dim TokenSize As Integer
Dim ConstStringSize As Integer
Dim Counter As Integer
Dim Counter2 As Integer
Dim Tempint1 As Integer
Dim Tempint2 As Integer
Dim TempString As String
Dim strMapName As String
Dim intFreeFile As Integer
intFreeFile = FreeFile
strMapName = App.Path & "\Scripts\ScriptInit.ini"
Open strMapName For Input As #intFreeFile
        Input #intFreeFile, TokenSize
        ReDim MyScripts(0 To TokenSize)
        For Counter = 0 To TokenSize
            Input #intFreeFile, Tempint1
            Input #intFreeFile, MyScripts(Counter).Name
        Next
Close #intFreeFile
For Counter2 = 0 To UBound(MyScripts, 1)
    intFreeFile = FreeFile
    strMapName = App.Path & "\Scripts\" & MyScripts(Counter2).Name & ".rsb"
    Open strMapName For Binary As intFreeFile
    Get intFreeFile, , TokenSize
    ReDim MyScripts(Counter2).MyNodes(0 To TokenSize)
        For Counter = 0 To TokenSize
            Get intFreeFile, , Tempint1
            Get intFreeFile, , Tempint2
            MyScripts(Counter2).MyNodes(Counter).OP = Tempint1
            MyScripts(Counter2).MyNodes(Counter).Pointer = Tempint2
        Next
    Close intFreeFile
intFreeFile = FreeFile
strMapName = Left$(strMapName, InStr(1, strMapName, ".") - 1)
strMapName = strMapName & ".rsc"
Open strMapName For Random As intFreeFile
        Get intFreeFile, , ConstStringSize
        ReDim MyScripts(Counter2).ConstStr(0 To ConstStringSize)
        For Counter = 0 To ConstStringSize
            Get intFreeFile, , MyScripts(Counter2).ConstStr(Counter)
        Next
Close intFreeFile
Next
End Sub

