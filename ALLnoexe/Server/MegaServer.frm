VERSION 5.00
Begin VB.Form MegaServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reload Me!"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   7425
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClearInfo 
      Caption         =   "Clear Game Info"
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "AutoScroll"
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   240
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6720
      TabIndex        =   6
      Text            =   "0"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmdUpdateNPC 
      Caption         =   "Update NPC Info on Map:"
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox NPCText 
      Height          =   1815
      Left            =   5640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "MegaServer.frx":0000
      Top             =   840
      Width           =   1695
   End
   Begin VB.Timer RndTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3360
      Top             =   120
   End
   Begin VB.CommandButton WSTrigger 
      Caption         =   "WSTrigger"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox GameMessage 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   5415
   End
   Begin VB.TextBox GameChat 
      Height          =   1695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "MegaServer.frx":0009
      Top             =   840
      Width           =   5415
   End
   Begin VB.ListBox NFOList 
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "MegaServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Megalodon MORPG Server
'Do whatever ya want with it just keep it open source:)
'http://home.kc.rr.com/megalodonsoft
'patrickostema@hotmail.com
Option Explicit

Private Sub cmdClearInfo_Click()
'Dim tempSo As String
'Dim i As Integer
'Dim y As Integer
GameChat.Text = "Megalodon MORPG Server (ALPHA)"
'WSock2.SendDataToAClient 0, -1, "#N(" & 0 + ALLServNPC(DataManage.ElMapo).NPCTotal & ")x25,6,2,0,0,"
'For i = 0 To 2
    'For y = 0 To MapSock(i).MaxCon
        'tempSo = tempSo & "map:" & Str$(i) & "Array:" & Str$(y) & ":" & Str$(MapSock(i).Sockers(y).Active) & "Socket:" & Str$(MapSock(i).Sockers(y).Socket)
        'tempSo = tempSo + vbCrLf
    'Next
'Next
'MsgBox tempSo
End Sub

Private Sub cmdUpdateNPC_Click()
DataManage.DispArray
End Sub
Private Sub Form_Load()
Const MyPort = 123
If WSock2.FireWS Then
    WSock2.Get411
    If WSock2.ServListen(MyPort) Then
        Me.Caption = "Loading Data"
        Me.Show
        DoEvents
        ScriptMod.OpenScripts
        NFOList.AddItem "Port: " & MyPort
        DataManage.ServerOpenIt App.Path & "\Maps\MapInit.ini"
        DataManage.OpenObjFile App.Path & "\test.mof"
        Me.Caption = "Megalodon MMORPG Server(ALPHA)"
    Else
        MsgBox "Killing WS"
        WSock2.KillWS
    End If
Else
    WSock2.KillWS
    MsgBox "Killing WS"
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
WSock2.KillWS
End Sub

Private Sub GameMessage_KeyPress(KeyAscii As Integer)
Dim mCounter As Integer
If KeyAscii = 13 Then
    For mCounter = 0 To DataManage.MapTotal
        If WSock2.SendChatToClients(mCounter, ("$Global Message: " & GameMessage.Text)) Then
            GameChat.Text = GameChat.Text + vbCrLf + "Global Message: " + GameMessage.Text
            GameMessage.Text = ""
        End If
    Next
End If
End Sub
Private Sub RndTimer_Timer()
DataManage.MoveAllServNPC
End Sub
Private Sub WSTrigger_KeyDown(KeyCode As Integer, Shift As Integer)
If WSock2.ServAccept() = False Then GameChat.Text = GameChat.Text + vbCrLf + "User had error attempting to connect."
End Sub
Private Sub WSTrigger_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Long
Dim TempS As String
Dim TempIndex As Integer
Dim StringBuf As String
Dim CMap As Integer
Dim CIndex As Integer
i = KeyCode
TempS = WSock2.Read(i)
If TempS <> "" Then
    CMap = CInt(Left$(TempS, InStr(1, TempS, ",")))  'temp
    TempS = Right$(TempS, Len(TempS) - InStr(1, TempS, ",")) 'temp
    If CMap > -1 And CMap <= UBound(DataManage.ALLServNPC, 1) Then
        CIndex = CInt(Left$(TempS, InStr(1, TempS, ","))) - DataManage.ALLServNPC(CMap).NPCTotal 'temp
    End If
    If CIndex < 0 Then CIndex = 0
    TempS = Right$(TempS, Len(TempS) - InStr(1, TempS, ",")) 'temp
    TempS = Left$(TempS, Len(TempS) - 1)
    If WSock2.CheckZIndex(CMap, CIndex, i) Then
        TempIndex = CIndex
    Else
        If Left$(TempS, 3) = "#MN" Then
            TempIndex = WSock2.GetIndexOnDisc(CMap, i)
        Else
            MegaServer.GameChat.Text = MegaServer.GameChat.Text + vbCrLf + "Err: Illegal player index from client" & TempS & ":" & CMap & ":" & CIndex
            TempIndex = WSock2.GetIndexOnDisc(CMap, i)
            'Exit Sub
        End If
    End If
End If
Recall:
If InStr(1, TempS, Chr(0)) > 0 Then
    If InStr(1, TempS, Chr(0)) < Len(TempS) Then StringBuf = Right(TempS, Len(TempS) - InStr(1, TempS, Chr(0)))
    TempS = Left(TempS, InStr(1, TempS, Chr(0)) - 1)
End If
If 0 < InStr(1, TempS, Chr(0)) Then
    MegaServer.GameChat.Text = MegaServer.GameChat.Text + vbCrLf + "Err:" & TempS
End If
If Left$(TempS, 1) = "$" Then
GameChat.Text = GameChat.Text + vbCrLf + (ALLMyNPC(CMap).MyNPC(TempIndex).Namer & ": " & TempS)
WSock2.SendChatToClients CMap, ("$" & ALLMyNPC(CMap).MyNPC(TempIndex).Namer & ": " & Right$(TempS, Len(TempS) - 1))
End If
If Left$(TempS, 1) = "#" Then
'MegaServer.GameChat.Text = MegaServer.GameChat.Text + vbCrLf + Str$(CMap) + "Incoming:" + TempS
DataManage.ElMapo = CMap
DataManage.UpdateInfo TempIndex, TempS
End If
If TempS = "" Then
TempIndex = WSock2.GetIndexOnDisc(CMap, i)
WSock2.DiscClient CMap, TempIndex, i
End If
If Check1.Value = 1 Then GameChat.SelStart = Len(GameChat.Text)
If StringBuf <> "" Then
    TempS = StringBuf
    StringBuf = ""
    GoTo Recall
End If
End Sub

