VERSION 5.00
Begin VB.Form PlaneForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plane Options"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   154
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   200
   Begin VB.CommandButton cmdFill 
      Caption         =   "Fill"
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdNPCBlock 
      Caption         =   "NPC Block"
      Height          =   495
      Left            =   1080
      TabIndex        =   11
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox picTileSelect 
      Height          =   495
      Left            =   720
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   2760
      Max             =   19
      TabIndex        =   9
      Top             =   1320
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   285
      Left            =   0
      Max             =   11
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdLayer 
      Caption         =   "Sky"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1050
      Width           =   855
   End
   Begin VB.CommandButton cmdLayer 
      Caption         =   "Floor"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   870
      Width           =   855
   End
   Begin VB.CommandButton cmdLayer 
      Caption         =   "Ground"
      Enabled         =   0   'False
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   690
      Width           =   855
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "Lock"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdWarp 
      Caption         =   "Exits"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdScript 
      Caption         =   "Scripted"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdBlock 
      Caption         =   "Blocked"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdPass 
      Caption         =   "Passable"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "PlaneForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TileX As Integer
Public TileY As Integer
Public Layer As Byte
Public TileProp As Byte
Option Explicit
Private Sub cmdBlock_Click()
TileProp = 0
Me.Caption = "Blocked"
cmdPass.Enabled = True
cmdPass.SetFocus
cmdLock.Enabled = True
cmdBlock.Enabled = False
cmdScript.Enabled = True
cmdWarp.Enabled = True
cmdNPCBlock.Enabled = True
End Sub

Private Sub cmdFill_Click()
TileForm.FillLayer
DoEvents
DXMod.DrawALLTiles
End Sub

Private Sub cmdLayer_Click(Index As Integer)
Select Case Index
Case 0:
cmdLayer(0).Enabled = False
cmdLayer(1).Enabled = True
cmdLayer(2).Enabled = True
Layer = 0
Case 1:
cmdLayer(0).Enabled = True
cmdLayer(1).Enabled = False
cmdLayer(2).Enabled = True
Layer = 1
Case 2:
cmdLayer(0).Enabled = True
cmdLayer(1).Enabled = True
cmdLayer(2).Enabled = False
Layer = 2
End Select
If cmdPass.Enabled Then cmdPass.SetFocus
If cmdBlock.Enabled Then cmdBlock.SetFocus
End Sub

Private Sub cmdLock_Click()
TileProp = 1
Me.Caption = "Lock"
cmdBlock.Enabled = True
cmdPass.Enabled = True
cmdPass.SetFocus
cmdScript.Enabled = True
cmdLock.Enabled = False
cmdWarp.Enabled = True
cmdNPCBlock.Enabled = True
End Sub

Private Sub cmdNPCBlock_Click()
TileProp = 2
Me.Caption = "NPC Block"
cmdBlock.Enabled = True
cmdPass.Enabled = True
cmdPass.SetFocus
cmdLock.Enabled = True
cmdScript.Enabled = True
cmdWarp.Enabled = True
cmdNPCBlock.Enabled = False
End Sub

Private Sub cmdPass_Click()
TileProp = 3
Me.Caption = "Passable"
cmdBlock.Enabled = True
cmdPass.Enabled = False
cmdLock.Enabled = True
cmdScript.Enabled = True
cmdWarp.Enabled = True
cmdNPCBlock.Enabled = True
End Sub

Private Sub cmdScript_Click()
TileProp = 4
Me.Caption = "Scripted"
cmdBlock.Enabled = True
cmdPass.Enabled = True
cmdPass.SetFocus
cmdLock.Enabled = True
cmdScript.Enabled = False
cmdWarp.Enabled = True
cmdNPCBlock.Enabled = True
End Sub

Private Sub cmdWarp_Click()
Dim Zexit As Integer
TileProp = 5
Me.Caption = "Exits"
Zexit = InputBox("which exit do u want to set?")
Select Case Zexit
Case 1: TileForm.NorthE = InputBox("North exit", , TileForm.NorthE)
Case 2: TileForm.SouthE = InputBox("South exit", , TileForm.SouthE)
Case 3: TileForm.EastE = InputBox("East exit", , TileForm.EastE)
Case 4: TileForm.WestE = InputBox("West exit", , TileForm.WestE)
End Select
DXMod.DrawALLTiles
DXMod.DrawTileSet
End Sub

Private Sub Form_Load()
PlaneForm.Left = TileForm.Width
PlaneForm.Width = MainForm.Width - TileForm.Width - 180
PlaneForm.Top = MainForm.Top + 60
PlaneForm.Height = 5 * (MainForm.Height / 10)
HScroll1.Top = PlaneForm.ScaleHeight - 16
HScroll1.Width = PlaneForm.ScaleWidth - 16
VScroll1.Left = HScroll1.Left + HScroll1.Width - 16
VScroll1.Height = (PlaneForm.ScaleHeight - 16) - VScroll1.Top
picTileSelect.Left = 0
picTileSelect.Width = VScroll1.Left
picTileSelect.Top = VScroll1.Top
picTileSelect.Height = VScroll1.Height

End Sub



Private Sub Form_Unload(Cancel As Integer)
Cancel = MainForm.ExitProgram()
End Sub

Private Sub HScroll1_Change()
DXMod.DrawTileSet
End Sub

Private Sub HScroll1_Scroll()
DXMod.DrawTileSet
End Sub
Private Sub picTileSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MainForm.MrClean = False Then
MsgBox "You must load a new or saved map array before editing"
Exit Sub
End If
If TileForm.ClickOption <> 0 Then Me.Caption = "Draw Mode Re-enabled"
TileForm.ClickOption = 0
If X < 0 Then Exit Sub
If Y < 0 Then Exit Sub
Dim intX As Integer
Dim intY As Integer
    intX = Int(X / 32)
    intY = Int(Y / 32)
TileY = intY + VScroll1.Value
TileX = intX + HScroll1.Value
DXMod.DrawTileSet
End Sub

Private Sub VScroll1_Change()
DXMod.DrawTileSet
End Sub

Private Sub VScroll1_Scroll()
DXMod.DrawTileSet
End Sub

