VERSION 5.00
Begin VB.Form ObjectForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ObjectForm"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   Begin VB.CommandButton cmdRemoveObj 
      Caption         =   "Remove"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   285
      Width           =   975
   End
   Begin VB.CommandButton cmdPlaceObj 
      Caption         =   "Place"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   45
      Width           =   975
   End
   Begin VB.ComboBox ObjCombo 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   1215
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
End
Attribute VB_Name = "ObjectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type IGObjects
    Name As String
    Value As Integer
    ObjType As Byte
    Power As Integer
    GIndex As Integer
    Description As String
    SIndex As Integer
End Type
Dim MyObjects() As IGObjects
Public TotalObjects As Integer

Private Sub cmdPlaceObj_Click()
TileForm.ClickOption = 3
PlaneForm.Caption = "Draw Mode Disabled"
End Sub

Private Sub cmdRemoveObj_Click()
TileForm.ClickOption = 4
PlaneForm.Caption = "Draw Mode Disabled"
End Sub

Private Sub Form_Load()
ObjectForm.Height = (MainForm.Height / 10) + 60
ObjectForm.Left = TileForm.Width
ObjectForm.Width = MainForm.Width - TileForm.Width - 180
ObjectForm.Top = PlaneForm.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = MainForm.ExitProgram()
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
ObjCombo.Clear
For i = 0 To TotalObjects - 1
    ObjCombo.AddItem i & ": " & MyObjects(i).Name
Next
ObjCombo.ListIndex = 0
NPCForm.Setupdrops
Me.Caption = TotalObjects
DXMod.DrawItem
End Sub
Public Function GetItemGraphic(ObjectIndex As Integer) As Integer
GetItemGraphic = MyObjects(ObjectIndex).GIndex
End Function

Private Sub ObjCombo_Change()
DXMod.DrawItem
End Sub
Private Sub ObjCombo_Click()
DXMod.DrawItem
End Sub
