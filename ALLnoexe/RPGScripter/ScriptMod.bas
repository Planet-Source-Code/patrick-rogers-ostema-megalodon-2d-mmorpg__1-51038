Attribute VB_Name = "ScriptMod"
'Megalodon Entertainment http://home.kc.rr.com/megalodonsoft
'Created by Patrick Rogers-Ostema
Option Explicit
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
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
Private TokenSize As Integer
Private ConstStringSize As Integer
Private MyNodes() As Node
Private Ints() As Integer
Private IntAlias() As String
Private Strs() As String
Private StrAlias() As String
Private ConstStr() As String
Public Sub CleanCode(AllCode As String)
'This sub cuts all the carraige returns out of the string
Dim CodeSplit As String
Dim CodeSplit2 As String
Do
If InStr(1, AllCode, Chr(13)) >= 1 Then
CodeSplit = Left$(AllCode, InStr(1, AllCode, Chr(13)) - 1)
CodeSplit2 = Right$(AllCode, Len(AllCode) - InStr(1, AllCode, Chr(13)) - 1)
AllCode = CodeSplit & " " & CodeSplit2
End If
Loop Until InStr(1, AllCode, Chr(13)) = 0
'Here we go!
AllCode = AllCode & " "
Compile (AllCode)
End Sub
Private Sub Compile(AllCode As String)
Dim CurCode As String
Dim Temp As Long
Dim Temp2 As Long
Dim TempVar As Integer
Dim LoopIndex() As Integer
Dim i As Integer
ReDim LoopIndex(0 To 0)
ReDim IntAlias(0 To 0)
ReDim MyNodes(0 To 0)
ReDim StrAlias(0 To 0)
ReDim ConstStr(0 To 0)
ScriptForm.OutBox.Text = ""
Do
'Get one alphanumeric value at a time
ResetPos:
CurCode = GetNextString(AllCode)
If CurCode <> "" Then
    ReDim Preserve MyNodes(0 To (UBound(MyNodes, 1) + 1))
    Temp2 = UBound(MyNodes, 1)
    'the big ass select case that turns our code into tokens
    'and other fun stuff
    If IsNumeric(CurCode) = False Then
        Select Case CurCode
        Case "for":
            ReDim Preserve LoopIndex(0 To (UBound(LoopIndex(), 1) + 1))
            MyNodes(Temp2).OP = OPfor
            LoopIndex(UBound(LoopIndex, 1)) = Temp2 + 1
            CurCode = GetNextString(AllCode)
            For i = 0 To UBound(IntAlias, 1)
                If IntAlias(i) = CurCode Then
                    MyNodes(Temp2).Pointer = i
                    Exit For
                End If
            Next i
            If ScriptForm.Check1.Value = 1 Then ScriptForm.OutBox.Text = ScriptForm.OutBox.Text & ("(" & Temp2 & ") " & Str$(MyNodes(Temp2).OP) & ":" & Str$(MyNodes(Temp2).Pointer) & vbCrLf)
            CurCode = GetNextString(AllCode)
            If CurCode = "=" Then
                ReDim Preserve MyNodes(0 To (UBound(MyNodes, 1) + 1))
                Temp2 = UBound(MyNodes, 1)
                CurCode = GetNextString(AllCode)
                If IsNumeric(CurCode) Then
                MyNodes(Temp2).OP = 7
                MyNodes(Temp2).Pointer = Val(CurCode)
                Else
                MyNodes(Temp2).OP = 6
                For i = 0 To UBound(IntAlias, 1)
                    If IntAlias(i) = CurCode Then
                        MyNodes(Temp2).Pointer = i
                        Exit For
                    End If
                Next i
                End If
                CurCode = GetNextString(AllCode)
                If CurCode = "to" Then
                    If ScriptForm.Check1.Value = 1 Then ScriptForm.OutBox.Text = ScriptForm.OutBox.Text & ("(" & Temp2 & ") " & Str$(MyNodes(Temp2).OP) & ":" & Str$(MyNodes(Temp2).Pointer) & vbCrLf)
                    ReDim Preserve MyNodes(0 To (UBound(MyNodes, 1) + 1))
                    Temp2 = UBound(MyNodes, 1)
                    CurCode = GetNextString(AllCode)
                    If IsNumeric(CurCode) Then
                        MyNodes(Temp2).OP = 7
                        MyNodes(Temp2).Pointer = Val(CurCode)
                    Else
                        MyNodes(Temp2).OP = 6
                        For i = 0 To UBound(IntAlias, 1)
                            If IntAlias(i) = CurCode Then
                                MyNodes(Temp2).Pointer = i
                                Exit For
                            End If
                        Next i
                    End If
                End If
            End If
        Case "if":
            MyNodes(Temp2).OP = OPiff
            MyNodes(Temp2).Pointer = 0
        Case "send":
            MyNodes(Temp2).OP = OPsend
            MyNodes(Temp2).Pointer = 0
        Case "pmoney", "php", "pstr", "parm", "pdsk", "pask", "pmaxhp", "hasobject", "giveobject", "takeobject", "warp":
            MyNodes(Temp2).OP = OPpvar
            Select Case CurCode
                Case "pmoney": MyNodes(Temp2).Pointer = 1
                Case "php": MyNodes(Temp2).Pointer = 2
                Case "pstr": MyNodes(Temp2).Pointer = 3
                Case "parm": MyNodes(Temp2).Pointer = 4
                Case "pdsk": MyNodes(Temp2).Pointer = 5
                Case "pask": MyNodes(Temp2).Pointer = 6
                Case "pmaxhp": MyNodes(Temp2).Pointer = 7
                Case "hasobject": MyNodes(Temp2).Pointer = 8
                Case "giveobject": MyNodes(Temp2).Pointer = 9
                Case "takeobject": MyNodes(Temp2).Pointer = 10
                Case "warp": MyNodes(Temp2).Pointer = 11
            End Select
            Temp = MyNodes(Temp2).Pointer
        Case "next":
            MyNodes(Temp2).OP = OPnex
            MyNodes(Temp2).Pointer = LoopIndex(UBound(LoopIndex, 1))
            If UBound(LoopIndex, 1) > 0 Then ReDim Preserve LoopIndex(0 To (UBound(LoopIndex(), 1) - 1))
        Case "<":
            MyNodes(Temp2).OP = OPlss
            MyNodes(Temp2).Pointer = Temp
        Case ">":
            MyNodes(Temp2).OP = OPgrt
            MyNodes(Temp2).Pointer = Temp
        Case "then":
            MyNodes(Temp2).OP = OPthn
            MyNodes(Temp2).Pointer = 0
        Case "else":
            MyNodes(Temp2).OP = OPelse
            MyNodes(Temp2).Pointer = 0
        Case "endif":
            MyNodes(Temp2).OP = OPndi
            MyNodes(Temp2).Pointer = 0
        Case "input":
            MyNodes(Temp2).OP = OPinp
            MyNodes(Temp2).Pointer = 0
        Case "int":
            CurCode = GetNextString(AllCode)
            Temp = UBound(IntAlias, 1)
            IntAlias(Temp) = CurCode
            MyNodes(Temp2).OP = OPint
            MyNodes(Temp2).Pointer = Temp
            ReDim Preserve IntAlias(0 To (UBound(IntAlias, 1) + 1))
        Case "str":
            CurCode = GetNextString(AllCode)
            Temp = UBound(StrAlias, 1)
            StrAlias(Temp) = CurCode
            MyNodes(Temp2).OP = OPstr
            MyNodes(Temp2).Pointer = Temp
            ReDim Preserve StrAlias(0 To (UBound(StrAlias, 1) + 1))
        Case "=":
            MyNodes(Temp2).OP = OPass
            MyNodes(Temp2).Pointer = Temp
        Case "+":
            MyNodes(Temp2).OP = OPadd
            MyNodes(Temp2).Pointer = 0
        Case "-":
            MyNodes(Temp2).OP = OPsub
            MyNodes(Temp2).Pointer = 0
        Case "*":
            MyNodes(Temp2).OP = OPmul
            MyNodes(Temp2).Pointer = 0
        Case "/":
            MyNodes(Temp2).OP = OPdiv
            MyNodes(Temp2).Pointer = 0
        Case "print":
            MyNodes(Temp2).OP = OPpri
            MyNodes(Temp2).Pointer = 0
        Case Else:
            Temp = -1
            For i = 0 To UBound(IntAlias, 1)
                If IntAlias(i) = CurCode Then
                    MyNodes(Temp2).OP = OPint
                    MyNodes(Temp2).Pointer = i
                    Temp = i
                    Exit For
                End If
            Next i
            For i = 0 To UBound(StrAlias, 1)
                If StrAlias(i) = CurCode Then
                    MyNodes(Temp2).OP = OPstr
                    MyNodes(Temp2).Pointer = i
                    Temp = i
                    Exit For
                End If
            Next i
            If Temp = -1 Then
                MyNodes(Temp2).OP = OPcstr
                MyNodes(Temp2).Pointer = -1
                If InStr(1, CurCode, Chr(34)) = 1 Then
                    CurCode = Right(CurCode, Len(CurCode) - 1)
                    If InStr(1, CurCode, Chr(34)) > 0 Then
                        CurCode = Left(CurCode, Len(CurCode) - 1)
                    Else
                        Do
                            CurCode = CurCode & " " & GetNextString(AllCode)
                            TempVar = TempVar + 1
                            If TempVar = 40 Then
                                TempVar = MsgBox("Either you have a very long string or you forgot both quotes, halt operation?", vbYesNo)
                                If TempVar = 6 Then Exit Do
                            End If
                        Loop Until InStr(1, CurCode, Chr(34)) > 0
                        TempVar = 0
                        CurCode = Left(CurCode, InStr(1, CurCode, Chr(34)) - 1)
                    End If
                    For i = 0 To UBound(ConstStr, 1)
                        If CurCode = ConstStr(i) Then
                        'if the constant already exits why make
                        'another one?
                        MyNodes(Temp2).Pointer = i
                        End If
                    Next i
                    If MyNodes(Temp2).Pointer = -1 Then
                        MyNodes(Temp2).Pointer = UBound(ConstStr, 1)
                        ConstStr(UBound(ConstStr, 1)) = CurCode
                        ConstStringSize = MyNodes(Temp2).Pointer
                        ReDim Preserve ConstStr(0 To UBound(ConstStr, 1) + 1)
                    End If
                End If
                If MyNodes(Temp2).Pointer = -1 Then
                    MsgBox "Syntax Error: " & CurCode
                    AllCode = Left$(ScriptForm.CodeBox.Text, InStr(1, ScriptForm.CodeBox.Text, CurCode) - 1)
                    AllCode = AllCode & vbCrLf & "Error-->" & Right$(ScriptForm.CodeBox.Text, Len(ScriptForm.CodeBox.Text) - Len(AllCode))
                    ScriptForm.CodeBox.Text = AllCode
                    Exit Sub
                End If
            End If
        End Select
    Else
            'this runs if the string is a const number
            MyNodes(Temp2).OP = OPnum
            MyNodes(Temp2).Pointer = Val(CurCode)
    End If
    If ScriptForm.Check1.Value = 1 Then ScriptForm.OutBox.Text = ScriptForm.OutBox.Text & ("(" & Temp2 & ") " & Str$(MyNodes(Temp2).OP) & ":" & Str$(MyNodes(Temp2).Pointer) & vbCrLf)
End If
Loop Until AllCode = ""
TokenSize = Temp2
End Sub
Private Function GetNextString(ByRef AllCode As String)
'This is our "Parsing" device
If InStr(1, AllCode, " ") >= 1 Then
    Do
        GetNextString = Left$(AllCode, InStr(1, AllCode, " ") - 1)
        AllCode = Right$(AllCode, Len(AllCode) - InStr(1, AllCode, " "))
    Loop Until GetNextString <> "" Or AllCode = ""
End If
End Function
Public Sub SaveTokens(strMapName As String, intFreeFile As Integer, ProcAll As Boolean)
Dim Counter As Integer
Dim TempString As String
Dim AllZCode As String
If ProcAll Then
    DeleteFile (strMapName)
    Open strMapName For Binary As intFreeFile
            Put intFreeFile, , (TokenSize - 1)
            '1 because first node is null
            For Counter = 1 To TokenSize
                Put intFreeFile, , CInt(MyNodes(Counter).OP)
                Put intFreeFile, , CInt(MyNodes(Counter).Pointer)
            Next
    Close intFreeFile
    intFreeFile = FreeFile
    strMapName = Left$(strMapName, InStr(1, strMapName, ".") - 1)
    strMapName = strMapName & ".rsc"
    DeleteFile (strMapName)
    Open strMapName For Random As intFreeFile
            Put intFreeFile, , ConstStringSize
            For Counter = 0 To ConstStringSize
                Put intFreeFile, , ConstStr(Counter)
            Next
    Close intFreeFile
End If
intFreeFile = FreeFile
strMapName = Left$(strMapName, InStr(1, strMapName, ".") - 1)
strMapName = strMapName & ".rss"
AllZCode = ScriptForm.ScriptCode
'delete the old file so we don't get any extra junk floating around
DeleteFile (strMapName)
Open strMapName For Random As intFreeFile
        While Len(AllZCode) > 0
            If InStr(1, AllZCode, Chr(13)) Then
                TempString = Left$(AllZCode, InStr(1, AllZCode, Chr(13)))
                AllZCode = Right$(AllZCode, Len(AllZCode) - Len(TempString))
            Else
                TempString = AllZCode
                AllZCode = ""
            End If
            Put intFreeFile, , TempString
        Wend
Close intFreeFile
End Sub
Public Sub OpenTokens(strMapName As String, intFreeFile As Integer, ProcAll As Boolean)
Dim Counter As Integer
Dim Tempint1 As Integer
Dim Tempint2 As Integer
Dim TempString As String
If ProcAll Then
    Open strMapName For Binary As intFreeFile
            Get intFreeFile, , TokenSize
            ReDim MyNodes(0 To TokenSize)
            For Counter = 0 To TokenSize
                Get intFreeFile, , Tempint1
                Get intFreeFile, , Tempint2
                MyNodes(Counter).OP = Tempint1
                MyNodes(Counter).Pointer = Tempint2
            Next
    Close intFreeFile
    intFreeFile = FreeFile
    strMapName = Left$(strMapName, InStr(1, strMapName, ".") - 1)
    strMapName = strMapName & ".rsc"
    Open strMapName For Random As intFreeFile
            Get intFreeFile, , ConstStringSize
            ReDim ConstStr(0 To ConstStringSize)
            For Counter = 0 To ConstStringSize
                Get intFreeFile, , ConstStr(Counter)
            Next
    Close intFreeFile
End If
intFreeFile = FreeFile
strMapName = Left$(strMapName, InStr(1, strMapName, ".") - 1)
strMapName = strMapName & ".rss"
TempString = "primer"
Open strMapName For Random As intFreeFile
        While Len(TempString) <> 0
            Get intFreeFile, , TempString
            ScriptForm.CodeBox.Text = ScriptForm.CodeBox.Text & TempString
        Wend
Close intFreeFile
End Sub
