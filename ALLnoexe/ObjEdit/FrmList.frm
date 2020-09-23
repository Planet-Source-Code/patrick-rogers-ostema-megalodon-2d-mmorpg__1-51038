VERSION 5.00
Begin VB.Form FrmList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   141
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   209
   Begin VB.ListBox ObjList 
      Height          =   450
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "FrmList"
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
Private Sub Form_Unload(Cancel As Integer)
Cancel = MDIMain.ExitProg()
End Sub
Public Function GetSIndex(index As Integer)
GetSIndex = MyObjects(index).SIndex
End Function

Public Sub ObjList_DblClick()
'add
If ObjList.List(ObjList.ListIndex) = "Add Object" Then
FrmEdit.Caption = ObjList.ListIndex
FrmEdit.txtDesc = ""
FrmEdit.txtName = ""
FrmEdit.txtPower.Text = "0"
FrmEdit.txtVal.Text = "0"
FrmEdit.ScriptIndex = -1
FrmEdit.VScroll1.Value = 0
FrmEdit.HScroll1.Value = 0
FrmEdit.cmbNPCType.ListIndex = -1
Else
FrmEdit.Caption = ObjList.ListIndex
FrmEdit.txtDesc = MyObjects(ObjList.ListIndex).Description
FrmEdit.txtName = MyObjects(ObjList.ListIndex).Name
FrmEdit.txtPower = MyObjects(ObjList.ListIndex).Power
FrmEdit.txtVal = MyObjects(ObjList.ListIndex).Value
FrmEdit.VScroll1.Value = MyObjects(ObjList.ListIndex).GIndex
FrmEdit.HScroll1.Value = MyObjects(ObjList.ListIndex).Power
FrmEdit.cmbNPCType.ListIndex = MyObjects(ObjList.ListIndex).ObjType
End If
End Sub
Public Sub AddObject(index As Integer, Name As String, Value As Integer, NPCType As Byte, Power As Integer, GIndex As Integer, Description As String, SIndex As Integer)
Dim i As Integer
ReDim Preserve MyObjects(0 To TotalObjects)
With MyObjects(index)
    .Name = Name
    .Value = Value
    .ObjType = NPCType
    .Power = Power
    .GIndex = GIndex
    .Description = Description
    .SIndex = SIndex
End With
ObjList.Clear
TotalObjects = TotalObjects + 1
For i = 0 To TotalObjects - 1
    ObjList.AddItem i & ": " & MyObjects(i).Name
Next
ObjList.AddItem "Add Object"
Me.Caption = TotalObjects
End Sub
Public Sub RemoveObject(index As Integer)
Dim i As Integer
If index < TotalObjects - 1 Then
For i = index To TotalObjects - 2
    MyObjects(i) = MyObjects(i + 1)
Next
End If
TotalObjects = TotalObjects - 1
If TotalObjects > 0 Then ReDim Preserve MyObjects(0 To TotalObjects - 1)
ObjList.Clear
For i = 0 To TotalObjects - 1
    ObjList.AddItem i & ": " & MyObjects(i).Name
Next
ObjList.AddItem "Add Object"
Me.Caption = TotalObjects
End Sub

Public Sub SaveIt(strMapName As String)
Dim intCounter As Byte
Dim intFreeFile As Integer
intFreeFile = FreeFile
Open strMapName For Binary As intFreeFile
    Put intFreeFile, , TotalObjects
    For intCounter = 0 To TotalObjects - 1
        Put intFreeFile, , MyObjects(intCounter)
    Next
Close intFreeFile
End Sub
Public Sub OpenIt(strMapName As String)
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
ObjList.Clear
For i = 0 To TotalObjects - 1
    ObjList.AddItem i & ": " & MyObjects(i).Name
Next
ObjList.AddItem "Add Object"
Me.Caption = TotalObjects
End Sub
Public Sub UpdateObject(index As Integer, Name As String, Value As Integer, NPCType As Byte, Power As Integer, GIndex As Integer, Description As String, SIndex As Integer)
With MyObjects(index)
    .Name = Name
    .Value = Value
    .ObjType = NPCType
    .Power = Power
    .GIndex = GIndex
    .Description = Description
    .SIndex = SIndex
End With
End Sub
