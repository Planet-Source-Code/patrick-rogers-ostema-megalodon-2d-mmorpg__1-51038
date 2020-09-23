Attribute VB_Name = "DXEngine"
'MORPG Engine
'Written by Patrick Rogers-Ostema
'http://home.kc.rr.com/megalodonsoft
'Use this code to do anything you want, all I ask
'is that you make it open source:-)
'report any bugs to patrickostema@hotmail.com please
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Type Tiles
    X As Integer
    Y As Integer
End Type
Private Type TileLayers
    Ground As Tiles
    Floor As Tiles
    Sky As Tiles
    GItem As Integer
    TileProp As Byte
End Type
Private Const DT_CALCRECT = &H400
Private Const DT_NOCLIP = &H100&
Dim MegaDX As New DirectX7
Dim MegaDD As DirectDraw7
Dim MegaFrontSurf As DirectDrawSurface7 'Front Surface
Dim MegaBackSurf As DirectDrawSurface7 'Represents our BackBuffers
Dim MegaSurface As DirectDrawSurface7 'Tileset's Surface
Dim MegaChatS As DirectDrawSurface7
Dim MegaSpriteBody As DirectDrawSurface7
Dim MegaSpriteHead As DirectDrawSurface7
Dim MegaSpriteWeap As DirectDrawSurface7
Dim MegaItem As DirectDrawSurface7
Dim MegaRun As Boolean 'Program Flowage
Dim MegaRect As RECT 'Rect of Screen
Dim MegaTimer As Long
Dim MegaFPSCounter As Integer
Dim MegaFPS As Integer 'Contains cycles per seconds
Dim MegaMap(50, 50) As TileLayers 'Our Big Sexy Map Array!
Public MyX As Integer 'X Coordinate of upperleft-hand corner tile
Public MyY As Integer 'Y Coordinate of upperleft-hand corner tile
Public MovingX As Integer 'How many pixels we're currently offset on the X axis
Public MovingY As Integer 'How many pixels we're currently offset on the Y axis
'You can change the MegaScroll variable, BUT
'if you do you must also change the "Moving"
'variables in the MoveIT sub so that
'they cycle through the 32 pixels(of a tile)
'correctly otherwise bad thingies can happen.
Private ServerMessage(6) As String
Public Chatting As Boolean
Public LastMove As Byte
Public ElListo As String
Dim Action As Byte
Dim Pointer As Byte
Public Crap As String
Public Const MEGASCROLL = 2
Option Explicit
Public Sub MapInit(Nfo As String)
Dim FileName As String
Dim intFreeFile As Integer
Dim fileNum As Integer

MyX = 15 'Put us approximately in the middle to start
MyY = 20
MovingY = 0
MovingX = 0
modNPC.MyIndex = 0
MyNPC(MyIndex).X = 24
MyNPC(MyIndex).Y = 25
MyNPC(MyIndex).LastMove = 2
Action = 0
WSock2.SendToServer "#MN" & Nfo
End Sub
Public Sub Mapload(ElMap As String)
Dim FileName As String
Dim intFreeFile As Integer
Dim fileNum As Integer
Dim xCounter As Byte
Dim yCounter As Byte
ElMap = Right$(ElMap, Len(ElMap) - 1)
FileName = App.Path & "\" & ElMap & ".map"
fileNum = FreeFile
Open FileName For Binary As fileNum
For xCounter = 0 To 49
    For yCounter = 0 To 49
        Get fileNum, , MegaMap(xCounter, yCounter)
    Next
Next
Close fileNum
MovingY = 0
MovingX = 0
Action = 0
End Sub
Public Sub Start(Nfo As String)
MegaInput.Init
Crap = Nfo
DXInit Nfo 'Fire it up!
End Sub
Private Sub DXInit(Nfo As String)
'Sub initializes our DirectXage
Dim MegaSurfMain As DDSURFACEDESC2
Dim MegaSurfFlip As DDSURFACEDESC2

    MegaRun = True 'Set our program off to running
    FrmMega.Show
    'Get DirectDraw going!
    Set MegaDD = MegaDX.DirectDrawCreate("")
    'Set up our Display
    MegaDD.SetCooperativeLevel FrmMega.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
    MegaDD.SetDisplayMode 640, 480, 16, 0, DDSDM_DEFAULT
    'Describe the flipping chain we want(notice 2 backbuffers)
    MegaSurfMain.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    MegaSurfMain.lBackBufferCount = 2
    MegaSurfMain.ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_PRIMARYSURFACE
    'Set up our screen surfaces
    Set MegaFrontSurf = MegaDD.CreateSurface(MegaSurfMain)
    MegaSurfFlip.ddsCaps.lCaps = DDSCAPS_BACKBUFFER
    Set MegaBackSurf = MegaFrontSurf.GetAttachedSurface(MegaSurfFlip.ddsCaps)
    'DirectDraw Font
    MegaBackSurf.SetForeColor vbYellow
    MegaBackSurf.SetFontTransparency True
    'Screen Rect
    MegaRect.Bottom = 480
    MegaRect.Right = 640
    'Load Tileset
    LoadSurfs
    'Start the main Loopage
    ReDim MyNPC(0)
    MapInit Nfo
End Sub
Public Sub MainLoop()
'Heart of our Tile-Engine
'Calls all the main subs/functions for the Engine
Do
If LostSurfaces Then LoadSurfs
MoveIT
DrawIT
DoEvents
Loop
End Sub
Public Sub ChangeTile(ByRef X As Byte, ByRef Y As Byte, Layer As Byte, NewData As String)
Select Case Layer
Case 1:
    MegaMap(X, Y).Ground.X = val(Left$(NewData, InStr(1, NewData, ",") - 1))
    NewData = Right$(NewData, Len(NewData) - InStr(1, NewData, ","))
    MegaMap(X, Y).Ground.Y = val(Left$(NewData, InStr(1, NewData, ",") - 1))
Case 2:
    MegaMap(X, Y).Floor.X = val(Left$(NewData, InStr(1, NewData, ",") - 1))
    NewData = Right$(NewData, Len(NewData) - InStr(1, NewData, ","))
    MegaMap(X, Y).Floor.Y = val(Left$(NewData, InStr(1, NewData, ",") - 1))
Case 3:
    MegaMap(X, Y).Sky.X = val(Left$(NewData, InStr(1, NewData, ",") - 1))
    NewData = Right$(NewData, Len(NewData) - InStr(1, NewData, ","))
    MegaMap(X, Y).Sky.Y = val(Left$(NewData, InStr(1, NewData, ",") - 1))
End Select
'modNPC.shat = "Map(" & Str$(X) & "," & Str$(Y) & ")" & Str$(Layer)
End Sub
Private Sub FPS(DC As Long)
'Gets and Shows the FPS
    If MegaTimer + 1000 <= MegaDX.TickCount Then
        MegaTimer = MegaDX.TickCount
        MegaFPS = MegaFPSCounter + 1
        MegaFPSCounter = 0
    Else
        MegaFPSCounter = MegaFPSCounter + 1
    End If
    TextOut DC, 0, 0, "FPS: " & Str$(MegaFPS), Len("FPS: " & Str$(MegaFPS))
    If modNPC.shat <> "" Then TextOut DC, 0, 40, modNPC.shat, Len(modNPC.shat)
    If ElListo <> "" Then DrawList DC
    MegaBackSurf.ReleaseDC DC
    MegaFrontSurf.Flip Nothing, DDFLIP_WAIT    'Flip that hot potatoe!
End Sub
Private Function ExclusiveMode() As Boolean
'Make sure we're still in exclusive mode
Dim lngTestExMode As Long
    lngTestExMode = MegaDD.TestCooperativeLevel
    
    If (lngTestExMode = DD_OK) Then
        ExclusiveMode = True
    Else
        ExclusiveMode = False
    End If
    
End Function
Private Function LostSurfaces() As Boolean
'Restore any lost surfaces
    LostSurfaces = False
    Do Until ExclusiveMode
        DoEvents
        LostSurfaces = True
    Loop
    DoEvents
    If LostSurfaces Then
        MegaDD.RestoreAllSurfaces
    End If
End Function
Private Sub Terminate()
'Clears everything out of memory
'and restores our original dispay settings
    MegaRun = False
    WSock2.KillWS
    MegaInput.Done
    MegaDD.RestoreDisplayMode
    MegaDD.SetCooperativeLevel 0, DDSCL_NORMAL
    Set MegaBackSurf = Nothing
    Set MegaFrontSurf = Nothing
    Set MegaSurface = Nothing
    Set MegaSpriteBody = Nothing
    Set MegaSpriteHead = Nothing
    Set MegaSpriteWeap = Nothing
    Set MegaItem = Nothing
    Set MegaChatS = Nothing
    Set MegaDD = Nothing
    Unload FrmMega
    End
End Sub
Private Sub MoveIT()
'Moves around tiles and creates offsets accordingly
CheckAction
If MegaInput.GetKeyState(1) Then Terminate
If modNPC.MyNPC(MyIndex).Walking = modNPC.North Then modNPC.WalkUp
If modNPC.MyNPC(MyIndex).Walking = modNPC.South Then modNPC.WalkDown
If modNPC.MyNPC(MyIndex).Walking = modNPC.West Then modNPC.WalkLeft
If modNPC.MyNPC(MyIndex).Walking = modNPC.East Then modNPC.WalkRight
End Sub

Private Sub DrawIT()
'Sub that gathers all the necessary tile 411 and
'draws it according to that info
Dim i As Integer
Dim j As Integer
Dim tx As Integer
Dim ty As Integer
MegaBackSurf.BltColorFill MegaRect, 0
For i = 0 To 20
    For j = 0 To 11
        tx = GetX(i)
        ty = GetY(j)
        MegaBackSurf.BltFast tx, ty, MegaSurface, GetRect(i, j), DDBLTFAST_WAIT
        If MegaMap(i + MyX, j + MyY).Floor.X <> 0 Or MegaMap(i + MyX, j + MyY).Floor.Y <> 0 Then MegaBackSurf.BltFast tx, ty, MegaSurface, GetRectF(i, j), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        If MegaMap(i + MyX, j + MyY).GItem > -1 Then MegaBackSurf.BltFast tx, ty, MegaItem, GetIRect(i, j), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Next
Next
CheckNPCs
For i = 0 To 20
    For j = 0 To 11
        If MegaMap(i + MyX, j + MyY).Sky.X <> 0 Or MegaMap(i + MyX, j + MyY).Sky.Y <> 0 Then MegaBackSurf.BltFast GetX(i), GetY(j), MegaSurface, GetRectS(i, j), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Next
Next
DrawChat
End Sub
Private Function GetRect(X As Integer, Y As Integer) As RECT
'Gets the basic 32x32 Rect and sends it to be
'altered if needed
With GetRect
    .Left = (MegaMap(X + MyX, Y + MyY).Ground.X) * 32
    .Right = .Left + 32
    .Top = (MegaMap(X + MyX, Y + MyY).Ground.Y) * 32
    .Bottom = .Top + 32
If X = 0 Or X = 20 Or Y = 0 Or Y = 11 Then GetRect = GetRectMods(X, Y, GetRect)
End With
End Function
Private Function GetIRect(X As Integer, Y As Integer) As RECT
'Gets the basic 32x32 Rect and sends it to be
'altered if needed
With GetIRect
    .Left = 0
    .Right = 32
    .Top = (MegaMap(X + MyX, Y + MyY).GItem) * 32
    .Bottom = .Top + 32
If X = 0 Or X = 20 Or Y = 0 Or Y = 11 Then GetIRect = GetRectMods(X, Y, GetIRect)
End With
End Function
Private Function GetRectF(X As Integer, Y As Integer) As RECT
'Gets the basic 32x32 Rect and sends it to be
'altered if needed
With GetRectF
    .Left = (MegaMap(X + MyX, Y + MyY).Floor.X) * 32
    .Right = .Left + 32
    .Top = (MegaMap(X + MyX, Y + MyY).Floor.Y) * 32
    .Bottom = .Top + 32
If X = 0 Or X = 20 Or Y = 0 Or Y = 11 Then GetRectF = GetRectMods(X, Y, GetRectF)
End With
End Function
Private Function GetRectS(X As Integer, Y As Integer) As RECT
'Gets the basic 32x32 Rect and sends it to be
'altered if needed
With GetRectS
    .Left = (MegaMap(X + MyX, Y + MyY).Sky.X) * 32
    .Right = .Left + 32
    .Top = (MegaMap(X + MyX, Y + MyY).Sky.Y) * 32
    .Bottom = .Top + 32
If X = 0 Or X = 20 Or Y = 0 Or Y = 11 Then GetRectS = GetRectMods(X, Y, GetRectS)
End With
End Function
Private Function GetRectMods(X As Integer, Y As Integer, TempRect As RECT) As RECT
'Crops the tile accordingly
GetRectMods = TempRect
With GetRectMods
If Y = 0 Then
    If MovingY > 0 Then
    .Top = .Top + MovingY
    End If
    If MovingY < 0 Then
    .Top = .Top + Abs(32 + MovingY)
    End If
End If
If Y = 11 Then
    If MovingY > 0 Then
    .Bottom = .Bottom - (32 - MovingY)
    End If
    If MovingY < 0 Then
    .Bottom = .Bottom + MovingY
    End If
End If
If X = 0 Then
    If MovingX < 0 Then
    .Left = .Left + Abs(32 + MovingX)
    End If
    If MovingX > 0 Then
    .Left = .Left + MovingX
    End If
End If
If X = 20 Then
    If MovingX > 0 Then
    .Right = .Right - (32 - MovingX)
    End If
    If MovingX < 0 Then
    .Right = .Right - MovingX
    End If
End If
End With
End Function
Private Function GetX(X As Integer) As Integer
'Decides where to place the Tile
'on the X axis
GetX = MovingX
If X = 0 And MovingX <> 0 Then
    GetX = 0
End If
If X >= 1 And MovingX < 0 Then
    GetX = GetX + 32
End If
GetX = (X * 32) - GetX
End Function
Private Function GetY(Y As Integer) As Integer
'Decides where to place the Tile
'on the Y axis
GetY = MovingY
If Y = 0 And MovingY <> 0 Then
    GetY = 0
End If
If Y >= 1 And MovingY < 0 Then
    GetY = GetY + 32
End If
GetY = (Y * 32) - GetY
End Function
Private Function LoadSurface(sWidth As Integer, sHeight As Integer, sFile As String, UseCkey As Boolean) As DirectDrawSurface7
'Loads our Tile Set
Dim MegaTempSurf As DDSURFACEDESC2
Dim aColorKey As DDCOLORKEY
    MegaTempSurf.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    MegaTempSurf.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    MegaTempSurf.lHeight = sHeight
    MegaTempSurf.lWidth = sWidth
    Set LoadSurface = MegaDD.CreateSurfaceFromFile(App.Path & sFile, MegaTempSurf)
If UseCkey Then
    aColorKey.high = vbBlack
    aColorKey.low = vbBlack
    LoadSurface.SetColorKey DDCKEY_SRCBLT, aColorKey
End If
End Function
Private Sub LoadSurfs()
Set MegaSurface = LoadSurface(512, 768, "\tileset.bmp", True)
Set MegaChatS = LoadSurface(640, 128, "\chatdialog.bmp", False)
Set MegaSpriteBody = LoadSurface(768, 192, "\Body.bmp", True)
Set MegaSpriteHead = LoadSurface(128, 160, "\Head.bmp", True)
Set MegaItem = LoadSurface(32, 864, "\Items.bmp", True)
Set MegaSpriteWeap = LoadSurface(512, 128, "\Weaps.bmp", True)
End Sub

Private Sub DrawChat()
Dim TempRect As RECT
Dim TempDC As Long
With TempRect
    .Left = 0
    .Right = 640
    .Top = 0
    .Bottom = 128
End With
MegaBackSurf.BltFast 0, 352, MegaChatS, TempRect, DDBLTFAST_WAIT
TempDC = MegaBackSurf.GetDC
SetBkMode TempDC, 1
SetTextColor TempDC, 8454143
With TempRect
    .Left = 10
    .Top = 357
    .Right = 50
    .Bottom = 377
End With
If Chatting Then
ServerMessage(0) = MegaInput.GetChatString
DrawText TempDC, ServerMessage(0), Len(ServerMessage(0)), TempRect, DT_NOCLIP
DrawText TempDC, ServerMessage(0), Len(ServerMessage(0)), TempRect, DT_CALCRECT
If TempRect.Right > 620 Then
    WSock2.SendToServer ("$" & ServerMessage(0))
    MegaInput.ClearChat
End If
End If
DrawChatting TempDC
End Sub
Private Sub DrawChatting(TempDC As Long)
Dim i As Byte
For i = 1 To 6
TextOut TempDC, 10, 364 + (i * 15), ServerMessage(i), Len(ServerMessage(i))
Next
ExitOut:
FPS (TempDC)
End Sub
Public Sub MoveChat(NewMessage As String)
Dim TempS2 As String
While NewMessage <> ""
TempS2 = NewMessage
If InStr(1, TempS2, ";") > 0 Then
    TempS2 = Right$(TempS2, Len(TempS2) - InStrRev(TempS2, ";"))
    NewMessage = Left$(NewMessage, Len(NewMessage) - Len(TempS2) - 1)
Else
    NewMessage = ""
End If
Dim i As Byte
i = 6
Do
ServerMessage(i) = ServerMessage(i - 1)
i = i - 1
Loop Until i = 1
ServerMessage(1) = TempS2
Wend
End Sub
Private Sub CheckAction()
MegaInput.CheckAllKeys
If MyNPC(MyIndex).Body = 5 Then Exit Sub
If MovingY = 0 And MovingX = 0 And MyNPC(MyIndex).Walking = 0 Then
    If MegaInput.GetKeyState(29) = False Then MyNPC(MyIndex).WeapStep = 0
    If MegaInput.GetKeyState(200) Then
        If MyNPC(MyIndex).Y > 0 Then
            If LastMove <> 1 And MegaMap(MyNPC(MyIndex).X, MyNPC(MyIndex).Y - 1).TileProp > 0 Then
                MyNPC(MyIndex).Walking = 1
                MyNPC(MyIndex).Y = MyNPC(MyIndex).Y - 1
                MyNPC(MyIndex).MovingY = 32
                SendToServer "#NY" & MyNPC(MyIndex).Y
                LastMove = 0
                MyNPC(MyIndex).LastMove = 1
            End If
        End If
        Exit Sub
    End If
    If MegaInput.GetKeyState(208) Then
        If MyNPC(MyIndex).Y < 49 Then
            If LastMove <> 2 And MegaMap(MyNPC(MyIndex).X, MyNPC(MyIndex).Y + 1).TileProp > 0 Then
            MyNPC(MyIndex).Walking = 2
            MyNPC(MyIndex).Y = MyNPC(MyIndex).Y + 1
            MyNPC(MyIndex).MovingY = -32
            SendToServer "#NY" & MyNPC(MyIndex).Y
            LastMove = 0
            MyNPC(MyIndex).LastMove = 2
            End If
        End If
        Exit Sub
    End If
    If MegaInput.GetKeyState(203) Then
        If MyNPC(MyIndex).X > 0 Then
            If LastMove <> 4 And MegaMap(MyNPC(MyIndex).X - 1, MyNPC(MyIndex).Y).TileProp > 0 Then
            MyNPC(MyIndex).Walking = 4
            MyNPC(MyIndex).X = MyNPC(MyIndex).X - 1
            MyNPC(MyIndex).MovingX = 32
            SendToServer "#NX" & MyNPC(MyIndex).X
            LastMove = 0
            MyNPC(MyIndex).LastMove = 4
            End If
        End If
        Exit Sub
    End If
    If MegaInput.GetKeyState(205) Then
        If MyNPC(MyIndex).X < 49 Then
            If LastMove <> 3 And MegaMap(MyNPC(MyIndex).X + 1, MyNPC(MyIndex).Y).TileProp > 0 Then
            MyNPC(MyIndex).Walking = 3
            MyNPC(MyIndex).X = MyNPC(MyIndex).X + 1
            MyNPC(MyIndex).MovingX = -32
            SendToServer "#NX" & MyNPC(MyIndex).X
            LastMove = 0
            MyNPC(MyIndex).LastMove = 3
            End If
        End If
        Exit Sub
    End If

If MegaInput.GetKeyState(29) And LastMove <> 5 Then
    MegaInput.SetKeyLast 29, True
    SendToServer "#NA"
    LastMove = 5
    Exit Sub
End If
'Use action variable to determine action
If MegaInput.GetKeyState(157) And MegaInput.GetKeyLast(157) = False Then
    MegaInput.SetKeyLast 157, True
    Select Case Action
    Case 0:
        If ElListo = "" Then
            ElListo = "Buy/Talk;Sell;View List/Equip;Pick up;Drop;Stats;"
        Else
            If ElListo = "Buy/Talk;Sell;View List/Equip;Pick up;Drop;Stats;" Then
                Action = Pointer + 1
                If Pointer = 0 Then
                    SendToServer "#NT" & modNPC.MyNPC(MyIndex).LastMove
                    Action = 0
                End If
                If Pointer = 3 Then
                    SendToServer "#NP"
                    Action = 0
                End If
                If Pointer = 1 Or Pointer = 2 Or Pointer = 4 Then SendToServer "#NL"
                If Pointer = 5 Then SendToServer "#Ns"
            Else
                SendToServer "#NB" & MyNPC(MyIndex).LastMove & Pointer
            End If
        End If
    Case 1:
        If ElListo = "Buy/Talk;Sell;View List/Equip;Pick up;Drop;" Then
            SendToServer "#NT" & modNPC.MyNPC(MyIndex).LastMove
        Else
            SendToServer "#NB" & MyNPC(MyIndex).LastMove & Pointer
        End If
    Case 2:
        SendToServer "#NS" & MyNPC(MyIndex).LastMove & Pointer
    Case 3:
        SendToServer "#NE" & Pointer
    Case 5:
        SendToServer "#NU" & Pointer
    Case 6:
        If Pointer = 5 Then SendToServer "#Na1"
        If Pointer = 6 Then SendToServer "#Na2"
        'SendToServer "#Ns"
    End Select
    Exit Sub
End If
    If MegaInput.GetKeyState(30) And MegaInput.GetKeyLast(30) = False Then
        MegaInput.SetKeyLast 30, True
        If ElListo = "" Then Exit Sub
        If Pointer > 0 Then
        Pointer = Pointer - 1
        End If
        Exit Sub
    End If
    If MegaInput.GetKeyState(44) And MegaInput.GetKeyLast(44) = False Then
        MegaInput.SetKeyLast 44, True
        If ElListo = "" Then Exit Sub
        If Pointer < 24 Then
        Pointer = Pointer + 1
        End If
        Exit Sub
    End If
    If MegaInput.GetKeyState(20) And Chatting = False Then
        MegaInput.SetKeyLast 20, True
        Chatting = True
        Exit Sub
    End If
    If MegaInput.GetKeyState(54) And Chatting = False Then
        MegaInput.SetKeyLast 54, True
        ResetAction
        Exit Sub
    End If
End If
End Sub
Private Sub ResetAction()
Action = 0
Pointer = 0
ElListo = ""
LastMove = 0
End Sub
Private Sub CheckNPCs()
Dim i As Integer
If modNPC.NPCCount >= 1 Then modNPC.UpdateNPCs
For i = 0 To modNPC.NPCCount
If MyNPC(i).Active Then
If MyNPC(i).Walking > 0 Then
    If Abs(MyNPC(i).MovingX) > 6 Or Abs(MyNPC(i).MovingY) > 6 Then MyNPC(i).Step = 1
    If Abs(MyNPC(i).MovingX) > 12 Or Abs(MyNPC(i).MovingY) > 12 Then MyNPC(i).Step = 0
    If Abs(MyNPC(i).MovingX) > 18 Or Abs(MyNPC(i).MovingY) > 18 Then MyNPC(i).Step = 2
    If Abs(MyNPC(i).MovingX) > 24 Or Abs(MyNPC(i).MovingY) > 24 Then MyNPC(i).Step = 0
Else
    MyNPC(i).Step = 0
End If
If MyX >= MyNPC(i).X - 20 And MyX <= MyNPC(i).X Then
    If MyY >= MyNPC(i).Y - 15 And MyY <= MyNPC(i).Y Then
            DFS GetX(MyNPC(i).X - MyX) + MyNPC(i).MovingX, GetY(MyNPC(i).Y - MyY) + MyNPC(i).MovingY, i
    End If
End If
'here weapstep
If MyNPC(i).WeapStep >= 1 Then
    MyNPC(i).WeapStep = MegaFPSCounter / 30 + 1
End If
End If
If i <> MyIndex Then
'here's the walk of the screen fixers
'If MyNPC(i).MovingX <> 0 Then
    'If MyX - 1 = MyNPC(i).X And MyNPC(i).Walking = 4 Then MegaBackSurf.BltFast 0, GetY(MyNPC(i).Y - MyY) + MyNPC(i).MovingY, MegaSprites, GetNPCRect(-1, MyNPC(i).Y - MyY, i), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    'If MyX = MyNPC(i).X And MyNPC(i).Walking = 3 Then MegaBackSurf.BltFast 0, GetY(MyNPC(i).Y - MyY) + MyNPC(i).MovingY, MegaSprites, GetNPCRect(-1, MyNPC(i).Y - MyY, i), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    'If MyX + 20 = MyNPC(i).X And MyNPC(i).Walking = 3 Then MegaBackSurf.BltFast 608 + (32 + MyNPC(i).MovingX), GetY(MyNPC(i).Y - MyY) + MyNPC(i).MovingY, MegaSprites, GetNPCRect(21, MyNPC(i).Y - MyY, i), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    'If MyX + 19 = MyNPC(i).X And MyNPC(i).Walking = 4 Then MegaBackSurf.BltFast 640 - (32 - Abs(MyNPC(i).MovingX)), GetY(MyNPC(i).Y - MyY) + MyNPC(i).MovingY, MegaSprites, GetNPCRect(21, MyNPC(i).Y - MyY, i), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
'End If
'If MyNPC(i).MovingY <> 0 Then
    'If MyY - 1 = MyNPC(i).Y And MyNPC(i).Walking = 1 Then MegaBackSurf.BltFast GetX(MyNPC(i).X - MyX) + MyNPC(i).MovingX, 0, MegaSprites, GetNPCRect(MyNPC(i).X - MyX, -1, i), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    'If MyY = MyNPC(i).Y And MyNPC(i).Walking = 2 Then MegaBackSurf.BltFast GetX(MyNPC(i).X - MyX) + MyNPC(i).MovingX, 0, MegaSprites, GetNPCRect(MyNPC(i).X - MyX, -1, i), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
'End If
End If

Next
End Sub
Private Function GetNPCRect(X As Integer, Y As Integer, NPCIndex As Integer) As RECT
With GetNPCRect
    .Top = 32 * MyNPC(NPCIndex).Body
    .Bottom = .Top + 32
    .Left = 3 * ((MyNPC(NPCIndex).LastMove - 1) * 32) + 32 * MyNPC(NPCIndex).Step
    .Right = .Left + 32
End With
If MyNPC(NPCIndex).WeapStep >= 1 Then
    GetNPCRect.Left = 384 + 3 * ((MyNPC(NPCIndex).LastMove) - 1) * 32 + ((MyNPC(NPCIndex).WeapStep - 1) * 32)
    GetNPCRect.Right = GetNPCRect.Left + 32
End If
If MyNPC(NPCIndex).Body = 5 Then
    GetNPCRect.Top = 160
    GetNPCRect.Bottom = 192
    GetNPCRect.Left = 0
    GetNPCRect.Right = 32
End If
If X = -1 Or X = 21 Or Y = -1 Or Y = 16 Then GetNPCRect = GetNPCRectMods(X, Y, GetNPCRect, NPCIndex)
If X = 0 Or X = 20 Or Y = 0 Or Y = 15 Then GetNPCRect = GetRectMods(X, Y, GetNPCRect)
End Function
Private Function GetNPCRectH(X As Integer, Y As Integer, NPCIndex As Integer) As RECT
With GetNPCRectH
    .Top = 32 * MyNPC(NPCIndex).Head
    .Bottom = .Top + 32
    .Left = ((MyNPC(NPCIndex).LastMove - 1) * 32)
    .Right = .Left + 32
End With
If X = -1 Or X = 21 Or Y = -1 Or Y = 16 Then GetNPCRectH = GetNPCRectMods(X, Y, GetNPCRectH, NPCIndex)
If X = 0 Or X = 20 Or Y = 0 Or Y = 15 Then GetNPCRectH = GetRectMods(X, Y, GetNPCRectH)
End Function
Private Function GetNPCRectMods(X As Integer, Y As Integer, TempRect As RECT, NPCIndex As Integer) As RECT
GetNPCRectMods = TempRect
With GetNPCRectMods
If X = -1 Then
    If MyNPC(NPCIndex).Walking = 4 Then
        .Left = .Left + (32 - MyNPC(NPCIndex).MovingX)
    ElseIf MyNPC(NPCIndex).Walking = 3 Then
        .Left = .Left + Abs(MyNPC(NPCIndex).MovingX)
    End If
End If
If X = 21 Then
    If MyNPC(NPCIndex).Walking = 3 Then
        .Right = .Right - (32 + MyNPC(NPCIndex).MovingX)
    ElseIf MyNPC(NPCIndex).Walking = 4 Then
        .Right = .Right - Abs(MyNPC(NPCIndex).MovingX)
    End If
End If
If Y = -1 Then
    If MyNPC(NPCIndex).Walking = 1 Then
        .Top = .Top + (32 - MyNPC(NPCIndex).MovingY)
    ElseIf MyNPC(NPCIndex).Walking = 2 Then
        .Top = .Top + Abs(MyNPC(NPCIndex).MovingY)
    End If
End If
End With
End Function

Private Sub DFS(X As Integer, Y As Integer, NPC As Integer)
'Stands for Draw Full Sprite:)
If MyNPC(NPC).Body = 6 Then Exit Sub
If MyNPC(NPC).Body = 5 Then GoTo DeadBlob
If MyNPC(NPC).Weap > 0 Then
    Select Case MyNPC(NPC).LastMove
        Case 1:
            If MyNPC(NPC).WeapStep > 0 Then MegaBackSurf.BltFast X, Y, MegaSpriteWeap, GetWeapRect(MyNPC(NPC).X - MyX, MyNPC(NPC).Y - MyY, NPC), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            MegaBackSurf.BltFast X, Y, MegaSpriteHead, GetNPCRectH(MyNPC(NPC).X - MyX, MyNPC(NPC).Y - MyY, NPC), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            MegaBackSurf.BltFast X, Y, MegaSpriteBody, GetNPCRect(MyNPC(NPC).X - MyX, MyNPC(NPC).Y - MyY, NPC), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            If MyNPC(NPC).WeapStep = 0 Then MegaBackSurf.BltFast X, Y, MegaSpriteWeap, GetWeapRect(MyNPC(NPC).X - MyX, MyNPC(NPC).Y - MyY, NPC), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        Case 2:
            If MyNPC(NPC).WeapStep = 0 Then MegaBackSurf.BltFast X, Y, MegaSpriteWeap, GetWeapRect(MyNPC(NPC).X - MyX, MyNPC(NPC).Y - MyY, NPC), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            MegaBackSurf.BltFast X, Y, MegaSpriteBody, GetNPCRect(MyNPC(NPC).X - MyX, MyNPC(NPC).Y - MyY, NPC), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            MegaBackSurf.BltFast X, Y, MegaSpriteHead, GetNPCRectH(MyNPC(NPC).X - MyX, MyNPC(NPC).Y - MyY, NPC), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            If MyNPC(NPC).WeapStep > 0 Then MegaBackSurf.BltFast X, Y, MegaSpriteWeap, GetWeapRect(MyNPC(NPC).X - MyX, MyNPC(NPC).Y - MyY, NPC), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        Case 3:
            MegaBackSurf.BltFast X, Y, MegaSpriteHead, GetNPCRectH(MyNPC(NPC).X - MyX, MyNPC(NPC).Y - MyY, NPC), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            MegaBackSurf.BltFast X, Y, MegaSpriteBody, GetNPCRect(MyNPC(NPC).X - MyX, MyNPC(NPC).Y - MyY, NPC), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            MegaBackSurf.BltFast X, Y, MegaSpriteWeap, GetWeapRect(MyNPC(NPC).X - MyX, MyNPC(NPC).Y - MyY, NPC), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        Case 4:
            MegaBackSurf.BltFast X, Y, MegaSpriteBody, GetNPCRect(MyNPC(NPC).X - MyX, MyNPC(NPC).Y - MyY, NPC), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            MegaBackSurf.BltFast X, Y, MegaSpriteWeap, GetWeapRect(MyNPC(NPC).X - MyX, MyNPC(NPC).Y - MyY, NPC), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            MegaBackSurf.BltFast X, Y, MegaSpriteHead, GetNPCRectH(MyNPC(NPC).X - MyX, MyNPC(NPC).Y - MyY, NPC), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    End Select
Else
    If MyNPC(NPC).Body <> 5 Then
        MegaBackSurf.BltFast X, Y, MegaSpriteBody, GetNPCRect(MyNPC(NPC).X - MyX, MyNPC(NPC).Y - MyY, NPC), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        MegaBackSurf.BltFast X, Y, MegaSpriteHead, GetNPCRectH(MyNPC(NPC).X - MyX, MyNPC(NPC).Y - MyY, NPC), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Else
DeadBlob:
        MegaBackSurf.BltFast X, Y, MegaSpriteBody, GetNPCRect(MyNPC(NPC).X - MyX, MyNPC(NPC).Y - MyY, NPC), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    End If
End If
End Sub
Private Function GetWeapRect(X As Integer, Y As Integer, NPCIndex As Integer) As RECT
With GetWeapRect
    .Top = 32 * (MyNPC(NPCIndex).Weap - 1)
    .Bottom = .Top + 32
    .Left = ((MyNPC(NPCIndex).LastMove - 1) * 32)
    .Right = .Left + 32
End With
If MyNPC(NPCIndex).WeapStep >= 1 Then
    GetWeapRect.Left = (((MyNPC(NPCIndex).LastMove) * 3) * 32) + (MyNPC(NPCIndex).WeapStep * 32)
    GetWeapRect.Right = GetWeapRect.Left + 32
End If
If X = -1 Or X = 21 Or Y = -1 Or Y = 16 Then GetWeapRect = GetNPCRectMods(X, Y, GetWeapRect, NPCIndex)
If X = 0 Or X = 20 Or Y = 0 Or Y = 15 Then GetWeapRect = GetRectMods(X, Y, GetWeapRect)
End Function

Public Sub AddItem(X As Byte, Y As Byte, Index As Integer)
MegaMap(X, Y).GItem = Index
End Sub
Public Sub DestroyItem(X As Byte, Y As Byte)
MegaMap(X, Y).GItem = -1
End Sub
Private Sub DrawList(DC As Long)
Dim TempList As String
Dim i As Integer
TempList = ElListo
For i = 0 To 25
If TempList <> "" Then
If Pointer = i Then TextOut DC, 415, i * 15, ">", Len(">")
If i < 25 Then
    TextOut DC, 425, (i * 15), Str$(i) & ":" & Left$(TempList, InStr(1, TempList, ";")), Len(Str$(i) & ":" & Left$(TempList, InStr(1, TempList, ";") - 1))
Else
    TextOut DC, 425, (25 * 15), Left$(TempList, InStr(1, TempList, ";")), Len(Left$(TempList, InStr(1, TempList, ";") - 1))
End If
'If Pointer = i Then TextOut DC, 430, i * 15, ">", Len(">")
'TextOut DC, 450, (i * 15), Str$(i) & ":", Len(Str$(i) & ":")
End If
'TextOut DC, 450, (25 * 15), "$" & Left(TempList, InStr(1, TempList, ";")), Len("$" & Left(TempList, InStr(1, TempList, ";") - 1))
TempList = Right$(TempList, Len(TempList) - InStr(1, TempList, ";"))
Next
End Sub
