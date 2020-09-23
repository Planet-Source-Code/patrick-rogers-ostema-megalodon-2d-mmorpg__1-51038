VERSION 5.00
Begin VB.Form TileForm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   207
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   3600
      Max             =   35
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   32
      TabIndex        =   0
      Top             =   2880
      Width           =   855
   End
End
Attribute VB_Name = "TileForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Client Types
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
'Server Types
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
Public NorthE As Integer
Public EastE As Integer
Public SouthE As Integer
Public WestE As Integer
Private MapArray(49, 49) As TileLayers
Private SMapArray(49, 49) As ServMap
Public ClickOption As Byte
Private Sub Form_Load()
BlankSetup
PlaneForm.TileProp = 3
End Sub
Public Sub FillLayer()
Dim TempX As Byte
Dim TempY As Byte
For TempX = 0 To 49
    For TempY = 0 To 49
        SMapArray(TempX, TempY).TileProp = PlaneForm.TileProp
        MapArray(TempX, TempY).TileProp = PlaneForm.TileProp
        Select Case PlaneForm.Layer
            Case 0:
                MapArray(TempX, TempY).Ground.X = PlaneForm.TileX
                MapArray(TempX, TempY).Ground.Y = PlaneForm.TileY
            Case 1:
                MapArray(TempX, TempY).Floor.X = PlaneForm.TileX
                MapArray(TempX, TempY).Floor.Y = PlaneForm.TileY
            Case 2:
                MapArray(TempX, TempY).Sky.X = PlaneForm.TileX
                MapArray(TempX, TempY).Sky.Y = PlaneForm.TileY
        End Select
    Next
Next
End Sub
Public Sub BlankSetup()
Dim TempX As Byte
Dim TempY As Byte
ClickOption = 0
TileForm.Width = MainForm.ScaleWidth - 3000
TileForm.Height = MainForm.ScaleHeight - 130
TileForm.HScroll1.Width = TileForm.ScaleWidth - 16
TileForm.HScroll1.Top = TileForm.ScaleHeight - 16
TileForm.VScroll1.Left = TileForm.ScaleWidth - 16
TileForm.VScroll1.Height = TileForm.ScaleHeight - 16
NPCForm.DeathScript = -1
NPCForm.NPCTotal = 0
For TempX = 0 To 49
    For TempY = 0 To 49
        SMapArray(TempX, TempY).TileProp = 3
        SMapArray(TempX, TempY).NPC.Index = -1
        SMapArray(TempX, TempY).GItem = -1
        SMapArray(TempX, TempY).Script = -1
        MapArray(TempX, TempY).GItem = -1
        MapArray(TempX, TempY).TileProp = 3
        MapArray(TempX, TempY).Ground.X = 0
        MapArray(TempX, TempY).Ground.Y = 0
        MapArray(TempX, TempY).Floor.X = 0
        MapArray(TempX, TempY).Floor.Y = 0
        MapArray(TempX, TempY).Sky.X = 0
        MapArray(TempX, TempY).Sky.Y = 0
    Next
Next
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MainForm.MrClean = False Then
MsgBox "You must load a new or saved map array before editing"
Exit Sub
End If
If X < 0 Then Exit Sub
If Y < 0 Then Exit Sub
Dim intX As Byte
Dim intY As Byte
Dim loopx As Byte
Dim loopy As Byte
    intX = Int(X / 32) + TileForm.HScroll1.Value
    intY = Int(Y / 32) + TileForm.VScroll1.Value
If intX > 49 Then Exit Sub
If intY > 49 Then Exit Sub
Me.Caption = Str$(intX) & " : " & Str$(intY)
If ClickOption = 0 And Button = 1 Then
Select Case PlaneForm.Layer
Case 0:
MapArray(intX, intY).Ground.X = PlaneForm.TileX
MapArray(intX, intY).Ground.Y = PlaneForm.TileY
Case 1:
MapArray(intX, intY).Floor.X = PlaneForm.TileX
MapArray(intX, intY).Floor.Y = PlaneForm.TileY
Case 2:
MapArray(intX, intY).Sky.X = PlaneForm.TileX
MapArray(intX, intY).Sky.Y = PlaneForm.TileY
End Select
If PlaneForm.TileProp = 4 Then
SMapArray(intX, intY).TileProp = PlaneForm.TileProp
SMapArray(intX, intY).Script = InputBox("Enter Script Index", , SMapArray(intX, intY).Script)
DoEvents
DXMod.DrawALLTiles
DXMod.DrawTileSet
End If
If SMapArray(intX, intY).TileProp = 1 And PlaneForm.TileProp <> 1 Then SMapArray(intX, intY).GItem = -1
    MapArray(intX, intY).TileProp = PlaneForm.TileProp
    SMapArray(intX, intY).TileProp = PlaneForm.TileProp
    DrawTile intX, intY
    If PlaneForm.TileProp = 1 Then
        SMapArray(intX, intY).GItem = InputBox("Enter Key Index")
        DoEvents
        DXMod.DrawALLTiles
    End If
End If
If SMapArray(intX, intY).NPC.Index > -1 And Button = 2 Then
If SMapArray(intX, intY).NPC.Mobile Then
    NPCForm.Check1.Value = 1
Else
    NPCForm.Check1.Value = 0
End If
NPCForm.txtNPCSpeech.Text = SMapArray(intX, intY).NPC.Speech
NPCForm.txtNPCName.Text = SMapArray(intX, intY).NPC.Name
NPCForm.NPCTypeCombo.ListIndex = SMapArray(intX, intY).NPC.NPCT
NPCForm.txtNPCAtts(0).Text = SMapArray(intX, intY).NPC.Attribs.HP
NPCForm.txtNPCAtts(1).Text = SMapArray(intX, intY).NPC.Attribs.Str
NPCForm.txtNPCAtts(2).Text = SMapArray(intX, intY).NPC.Attribs.Arm
NPCForm.txtNPCAtts(3).Text = SMapArray(intX, intY).NPC.Attribs.DSk
NPCForm.txtNPCAtts(4).Text = SMapArray(intX, intY).NPC.Attribs.ASk
NPCForm.txtNPCAtts(5).Text = SMapArray(intX, intY).NPC.Attribs.XP
NPCForm.NPCDeadDropCombo(0).ListIndex = SMapArray(intX, intY).NPC.Attribs.DeadDrops(0)
NPCForm.NPCDeadDropCombo(1).ListIndex = SMapArray(intX, intY).NPC.Attribs.DeadDrops(1)
NPCForm.Label1 = "Index: " & SMapArray(intX, intY).NPC.Index
NPCForm.VScroll1.Value = SMapArray(intX, intY).NPC.Body
NPCForm.VScroll2.Value = SMapArray(intX, intY).NPC.Head
NPCForm.DeathScript = SMapArray(intX, intY).NPC.DeathScript
DXMod.DrawALLTiles
DXMod.DrawNPConForm
Exit Sub
End If
If ClickOption = 1 Then
If SMapArray(intX, intY).NPC.Index = -1 Then NPCForm.NPCTotal = NPCForm.NPCTotal + 1
SMapArray(intX, intY).TileProp = 3
MapArray(intX, intY).TileProp = 3
SMapArray(intX, intY).NPC.Mobile = NPCForm.Check1.Value
SMapArray(intX, intY).NPC.Speech = NPCForm.txtNPCSpeech.Text
SMapArray(intX, intY).NPC.Name = NPCForm.txtNPCName.Text
SMapArray(intX, intY).NPC.NPCT = NPCForm.NPCTypeCombo.ListIndex
SMapArray(intX, intY).NPC.Attribs.HP = Val(NPCForm.txtNPCAtts(0).Text)
SMapArray(intX, intY).NPC.Attribs.Str = Val(NPCForm.txtNPCAtts(1).Text)
SMapArray(intX, intY).NPC.Attribs.Arm = Val(NPCForm.txtNPCAtts(2).Text)
SMapArray(intX, intY).NPC.Attribs.DSk = Val(NPCForm.txtNPCAtts(3).Text)
SMapArray(intX, intY).NPC.Attribs.ASk = Val(NPCForm.txtNPCAtts(4).Text)
SMapArray(intX, intY).NPC.Attribs.XP = Val(NPCForm.txtNPCAtts(5).Text)
SMapArray(intX, intY).NPC.Attribs.DeadDrops(0) = NPCForm.NPCDeadDropCombo(0).ListIndex
SMapArray(intX, intY).NPC.Attribs.DeadDrops(1) = NPCForm.NPCDeadDropCombo(1).ListIndex
SMapArray(intX, intY).NPC.DeathScript = NPCForm.DeathScript
If SMapArray(intX, intY).NPC.Index = -1 Then SMapArray(intX, intY).NPC.Index = NPCForm.NPCTotal
SMapArray(intX, intY).NPC.Body = NPCForm.VScroll1.Value
SMapArray(intX, intY).NPC.Head = NPCForm.VScroll2.Value
NPCForm.Label1 = "Index: " & SMapArray(intX, intY).NPC.Index
DXMod.DrawALLTiles
End If
If ClickOption = 2 Then
If SMapArray(intX, intY).NPC.Index = -1 Then Exit Sub
For loopx = 0 To 49
    For loopy = 0 To 49
        If SMapArray(loopx, loopy).NPC.Index > SMapArray(intX, intY).NPC.Index Then
            SMapArray(loopx, loopy).NPC.Index = SMapArray(loopx, loopy).NPC.Index - 1
        End If
    Next
Next
NPCForm.NPCTotal = NPCForm.NPCTotal - 1
SMapArray(intX, intY).NPC.Index = -1
NPCForm.Label1 = "Index: " & NPCForm.NPCTotal
DXMod.DrawALLTiles
End If
If ClickOption = 3 Then
SMapArray(intX, intY).GItem = ObjectForm.ObjCombo.ListIndex
DXMod.DrawALLTiles
End If
If ClickOption = 4 Then
SMapArray(intX, intY).GItem = -1
DXMod.DrawALLTiles
End If
End Sub
Public Function GetMapItem(X As Byte, Y As Byte) As Integer
GetMapItem = SMapArray(X, Y).GItem
End Function
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button > 0 Then Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = MainForm.ExitProgram()
End Sub
Public Function GetTileY(X As Byte, Y As Byte) As Integer
GetTileY = MapArray(X, Y).Ground.Y
End Function
Public Function GetTileX(X As Byte, Y As Byte) As Integer
GetTileX = MapArray(X, Y).Ground.X
End Function
Public Function GetTileProp(X As Byte, Y As Byte) As Integer
GetTileProp = SMapArray(X, Y).TileProp
End Function
Public Function GetTileFY(X As Byte, Y As Byte) As Integer
GetTileFY = MapArray(X, Y).Floor.Y
End Function
Public Function GetTileFX(X As Byte, Y As Byte) As Integer
GetTileFX = MapArray(X, Y).Floor.X
End Function
Public Function GetTileSY(X As Byte, Y As Byte) As Integer
GetTileSY = MapArray(X, Y).Sky.Y
End Function
Public Function GetTileSX(X As Byte, Y As Byte) As Integer
GetTileSX = MapArray(X, Y).Sky.X
End Function

Public Function GetNPCIndex(X As Byte, Y As Byte) As Integer
GetNPCIndex = SMapArray(X, Y).NPC.Index
End Function
Public Function GetNPCB(X As Byte, Y As Byte) As Byte
GetNPCB = SMapArray(X, Y).NPC.Body
End Function
Public Function GetNPCH(X As Byte, Y As Byte) As Byte
GetNPCH = SMapArray(X, Y).NPC.Head
End Function
Private Sub HScroll1_Change()
Me.Caption = HScroll1.Value
DXMod.DrawALLTiles
End Sub

Private Sub HScroll1_Scroll()
DXMod.DrawALLTiles
End Sub

Private Sub VScroll1_Change()
DXMod.DrawALLTiles
End Sub

Private Sub VScroll1_Scroll()
DXMod.DrawALLTiles
End Sub

Public Sub OpenIt(strMapName As String, intFreeFile As Integer)
Dim intXCounter As Byte
Dim intYCounter As Byte
Open strMapName For Binary As intFreeFile
    For intXCounter = 0 To 49
        For intYCounter = 0 To 49
            Get intFreeFile, , MapArray(intXCounter, intYCounter)
        Next
    Next
Close intFreeFile
End Sub
Public Sub SaveIt(strMapName As String, intFreeFile As Integer)
Dim intXCounter As Byte
Dim intYCounter As Byte
Open strMapName For Binary As intFreeFile
    For intXCounter = 0 To 49
        For intYCounter = 0 To 49
            Put intFreeFile, , MapArray(intXCounter, intYCounter)
        Next
    Next
Close intFreeFile
End Sub
Public Sub ServerSaveIt(strMapName As String, intFreeFile As Integer)
Dim intXCounter As Byte
Dim intYCounter As Byte
Open strMapName For Binary As intFreeFile
    For intXCounter = 0 To 49
        For intYCounter = 0 To 49
            Put intFreeFile, , SMapArray(intXCounter, intYCounter)
        Next
    Next
Put intFreeFile, , NPCForm.NPCTotal
Put intFreeFile, , TileForm.NorthE
Put intFreeFile, , TileForm.SouthE
Put intFreeFile, , TileForm.EastE
Put intFreeFile, , TileForm.WestE
Close intFreeFile
End Sub
Public Sub ServerOpenIt(strMapName As String, intFreeFile As Integer)
Dim intXCounter As Byte
Dim intYCounter As Byte
Dim TempTotal As Integer
Dim TempE As Integer
Open strMapName For Binary As intFreeFile
    For intXCounter = 0 To 49
        For intYCounter = 0 To 49
            Get intFreeFile, , SMapArray(intXCounter, intYCounter)
        Next
    Next
Get intFreeFile, , TempTotal
Get intFreeFile, , TempE
TileForm.NorthE = TempE
Get intFreeFile, , TempE
TileForm.SouthE = TempE
Get intFreeFile, , TempE
TileForm.EastE = TempE
Get intFreeFile, , TempE
TileForm.WestE = TempE
Close intFreeFile
NPCForm.NPCTotal = TempTotal
NPCForm.Label1 = "Index: " & TempTotal
End Sub

