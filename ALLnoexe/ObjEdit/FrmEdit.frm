VERSION 5.00
Begin VB.Form FrmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   142
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   162
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   30
      TabIndex        =   12
      Top             =   1680
      Width           =   615
   End
   Begin VB.PictureBox ObjPic 
      Height          =   495
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   11
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtDesc 
      Height          =   735
      Left            =   690
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "FrmEdit.frx":0000
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdScript 
      Caption         =   "Script"
      Height          =   375
      Left            =   30
      TabIndex        =   9
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtPower 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   840
      Max             =   9999
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   510
      LargeChange     =   5
      Left            =   480
      Max             =   26
      TabIndex        =   5
      Top             =   705
      Width           =   255
   End
   Begin VB.TextBox txtVal 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.ComboBox cmbNPCType 
      Height          =   315
      ItemData        =   "FrmEdit.frx":000C
      Left            =   0
      List            =   "FrmEdit.frx":002E
      TabIndex        =   2
      Text            =   "Obj Type"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Power:"
      Height          =   255
      Left            =   810
      TabIndex        =   8
      Top             =   1005
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Value:"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   405
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   495
   End
End
Attribute VB_Name = "FrmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ScriptIndex As Integer
Private Sub cmdAdd_Click()
If cmbNPCType.ListIndex = -1 Then
MsgBox "You must select Object type before proceeding."
Exit Sub
End If
If FrmList.ObjList.List(Me.Caption) = "Add Object" Then
FrmList.AddObject Me.Caption, txtName.Text, txtVal.Text, cmbNPCType.ListIndex, txtPower.Text, VScroll1.Value, txtDesc.Text, ScriptIndex
Else
FrmList.UpdateObject Me.Caption, txtName.Text, txtVal.Text, cmbNPCType.ListIndex, txtPower.Text, VScroll1.Value, txtDesc.Text, ScriptIndex
End If
End Sub
Private Sub cmdScript_Click()
Dim blah As Variant
blah = InputBox("Enter Script Index", , FrmList.GetSIndex(Me.Caption))
If IsNumeric(blah) Then
ScriptIndex = CInt(blah)
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = MDIMain.ExitProg()
End Sub
Private Sub HScroll1_Change()
txtPower.Text = HScroll1.Value
End Sub
Private Sub txtPower_Change()
HScroll1.Value = Val(txtPower.Text)
End Sub
Private Sub VScroll1_Change()
DXMod.DrawObject
End Sub
