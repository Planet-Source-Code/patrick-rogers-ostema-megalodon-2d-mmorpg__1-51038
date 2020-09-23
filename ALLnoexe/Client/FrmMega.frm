VERSION 5.00
Begin VB.Form FrmMega 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "FrmMega.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton WSTrigger 
      Caption         =   "WSTrigger"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "FrmMega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MORPG Engine
'Written by Patrick Rogers-Ostema
'http://home.kc.rr.com/megalodonsoft
'Use this code to do anything you want, all I ask
'is that you make it open source:-)
'patrickostema@hotmail.com please
Option Explicit
Private Sub Form_Load()
Dim IPString As String
Dim TempPort As Long
Dim TempBody As String
modNPC.MyMapIndex = -1
IPString = InputBox("Enter Server IP", , "127.0.0.1")
TempPort = 123
TempBody = InputBox("Enter Body(0-1)")
TempBody = TempBody & "," & InputBox("Enter Head(0-1)")
TempBody = TempBody & ",0," & InputBox("name?")
If WSock2.FireWS Then
    If WSock2.Connecter(IPString, , TempPort) Then
        DXEngine.Start TempBody
    Else
        MsgBox "Error Connecting"
        WSock2.KillWS
        Unload FrmMega
    End If
End If
End Sub

Private Sub WSTrigger_KeyUp(KeyCode As Integer, Shift As Integer)
WSock2.Read
End Sub


