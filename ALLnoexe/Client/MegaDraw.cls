VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBInput Class
'
Option Explicit
'DirectX Variables
Private oDX As New DirectX7             ' the main DirectX object
Private oDI As DirectInput              ' our input object
Private oDIDEV As DirectInputDevice     ' our keyboard input device object
Private oDIState As DIKEYBOARDSTATE     ' Keyboard state structure
Private aKeys(211) As Boolean           ' array to hold all keystates
'
' Function return the state of a key (press=true, unpress=false) according to
' the key-constants in VBGlobal.bas
'
Public Function GetKeyState(lKey As Long) As Boolean
    GetKeyState = aKeys(lKey)
End Function
'
' Initializes the DirectInput objects (keyboard & mouse)
' Parameter lHwnd must be the window-handle of your main application window
'
Public Function Init(lHwnd As Long) As Boolean

    On Error GoTo ErrorHandler
    '
    ' Create a DirectInput object
    '
    Set oDI = oDX.DirectInputCreate()
    '
    ' Create a keyboard as input-device
    '
    Set oDIDEV = oDI.CreateDevice("GUID_SysKeyboard")
    '
    ' Set the exchange-data format to the keyboard structure format (DIKEYBOARDSTATE)
    '
    oDIDEV.SetCommonDataFormat DIFORMAT_KEYBOARD
    '
    ' Set the nonexclusively mode, when in foreground
    '
    oDIDEV.SetCooperativeLevel lHwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    '
    ' Aquire our input device
    '
    oDIDEV.Acquire
    '
    ' Create a mouse as input-device
    
    Init = True
    Exit Function
ErrorHandler:
    Init = False
End Function
'
' Function updates all keystates
'
Public Sub CheckAllKeys()
    Dim i As Integer
    '
    ' Get the all keystates from the input device
    '
    oDIDEV.GetDeviceStateKeyboard oDIState
    
    '
    ' transfer all keystates in our storage array
    '
    For i = 1 To 211
        If oDIState.Key(i) <> 0 Then
            aKeys(i) = True             ' key is pressed
        Else
            aKeys(i) = False            ' key is unpressed
        End If
    Next
    
End Sub
Public Sub Done()
    oDIDEV.Unacquire
End Sub





