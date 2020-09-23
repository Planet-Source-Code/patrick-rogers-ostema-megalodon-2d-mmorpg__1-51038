VERSION 5.00
Begin VB.Form NPCForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NPCForm"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   193
   Begin VB.CheckBox Check1 
      Caption         =   "Mobile"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   855
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   540
      Left            =   840
      Max             =   4
      TabIndex        =   18
      Top             =   0
      Width           =   255
   End
   Begin VB.ComboBox NPCDeadDropCombo 
      Height          =   315
      Index           =   1
      ItemData        =   "NPCForm.frx":0000
      Left            =   2160
      List            =   "NPCForm.frx":0002
      TabIndex        =   17
      Top             =   2040
      Width           =   735
   End
   Begin VB.ComboBox NPCDeadDropCombo 
      Height          =   315
      Index           =   0
      ItemData        =   "NPCForm.frx":0004
      Left            =   2160
      List            =   "NPCForm.frx":0006
      TabIndex        =   16
      Top             =   1710
      Width           =   735
   End
   Begin VB.TextBox txtNPCAtts 
      Height          =   285
      Index           =   5
      Left            =   1800
      TabIndex        =   15
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtNPCAtts 
      Height          =   285
      Index           =   4
      Left            =   1440
      TabIndex        =   14
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtNPCAtts 
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   13
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtNPCAtts 
      Height          =   285
      Index           =   2
      Left            =   720
      TabIndex        =   12
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtNPCAtts 
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtNPCAtts 
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtNPCSpeech 
      Height          =   735
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton cmdNPCScript 
      Caption         =   "Script"
      Height          =   315
      Left            =   2160
      TabIndex        =   7
      Top             =   570
      Width           =   735
   End
   Begin VB.ComboBox NPCTypeCombo 
      Height          =   315
      ItemData        =   "NPCForm.frx":0008
      Left            =   960
      List            =   "NPCForm.frx":0018
      TabIndex        =   6
      Text            =   "NPC Type"
      Top             =   570
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemoveNPC 
      Caption         =   "Remove"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   270
      Width           =   750
   End
   Begin VB.CommandButton cmdPlaceNPC 
      Caption         =   "Place"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   15
      Width           =   750
   End
   Begin VB.TextBox txtNPCName 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "NPC Name"
      Top             =   0
      Width           =   1095
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   540
      Left            =   540
      Max             =   4
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   0
      Width           =   540
   End
   Begin VB.Label Label2 
      Caption         =   " HP    Str   Arm DSk  ASk  XP "
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1710
      Width           =   2295
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Index: 0"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   330
      Width           =   1095
   End
End
Attribute VB_Name = "NPCForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NPCTotal As Integer
Public DeathScript As Integer
Private Sub cmdNPCScript_Click()
DeathScript = InputBox("Enter death script", , DeathScript)
DoEvents
DXMod.DrawALLTiles
DXMod.DrawTileSet
End Sub

Private Sub cmdPlaceNPC_Click()
TileForm.ClickOption = 1
PlaneForm.Caption = "Draw Mode Disabled"
End Sub

Private Sub cmdRemoveNPC_Click()
TileForm.ClickOption = 2
PlaneForm.Caption = "Draw Mode Disabled"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = MainForm.ExitProgram()
End Sub
Private Sub Form_Load()
NPCForm.Height = 3 * (MainForm.Height / 10) + 60
NPCForm.Left = TileForm.Width
NPCForm.Width = MainForm.Width - TileForm.Width - 180
NPCForm.Top = PlaneForm.Height + ObjectForm.Height
End Sub

Private Sub VScroll1_Change()
DXMod.DrawNPConForm
End Sub

Private Sub VScroll2_Change()
DXMod.DrawNPConForm
End Sub
Public Sub Setupdrops()
For i = 0 To ObjectForm.TotalObjects - 1
    NPCDeadDropCombo(0).AddItem ObjectForm.ObjCombo.List(i)
    NPCDeadDropCombo(1).AddItem ObjectForm.ObjCombo.List(i)
Next
End Sub
