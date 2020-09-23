VERSION 5.00
Begin VB.MDIForm MainForm 
   BackColor       =   &H8000000C&
   Caption         =   "Megalodon MORPG Map Editor"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4935
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuslide 
         Caption         =   "Client File"
         Begin VB.Menu mnuFileOpen 
            Caption         =   "Open"
         End
         Begin VB.Menu mnuFileSaveAs 
            Caption         =   "Save As"
         End
         Begin VB.Menu mnuFileSave 
            Caption         =   "Save"
         End
      End
      Begin VB.Menu mnuslide2 
         Caption         =   "Server File"
         Begin VB.Menu mnuServerOpen 
            Caption         =   "Open"
         End
         Begin VB.Menu mnuServerSaveAs 
            Caption         =   "Save As"
         End
         Begin VB.Menu mnuServerSave 
            Caption         =   "Save"
         End
      End
      Begin VB.Menu mnuObjectFile 
         Caption         =   "Object File"
         Begin VB.Menu mnuObjectLoad 
            Caption         =   "Load"
         End
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MORPG Map Editor
'Do anything ya want with it just please keep it open
'source so others can learn:)
'http://home.kc.rr.com/megalodonsoft
'patrickostema@hotmail.com
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (lpOpenFilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (lpOpenFilename As OPENFILENAME) As Long
Public MrClean As Boolean
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As String
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As LoadPictureColorConstants
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Dim CurSave As String
Dim CurSave2 As String
Dim CurSave3 As String
Option Explicit
Private Sub MDIForm_Load()
Me.Show
TileForm.Show
PlaneForm.Show
ObjectForm.Show
NPCForm.Show
DXMod.Init
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Cancel = ExitProgram()
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub
Public Function ExitProgram() As Boolean
Dim intRetVal As Integer
        intRetVal = MsgBox("Are you sure you want to quit?", vbYesNo)
        If intRetVal = vbYes Then
            KillDX
            End
        Else
            ExitProgram = vbCancel
            Exit Function
        End If
End Function

Private Sub mnuFileNew_Click()
CurSave = ""
CurSave2 = ""
TileForm.BlankSetup
MrClean = True
mnuObjectLoad_Click
DoEvents
DXMod.DrawTileSet
DXMod.DrawNPConForm
DXMod.DrawALLTiles
End Sub

Private Sub mnuFileOpen_Click()
Dim rc As Long
Dim pOpenfilename As OPENFILENAME
Dim strMapName As String
Dim intFreeFile As Integer
Dim TempTile As Byte
Const MAX_BUFFER_LENGTH = 256
With pOpenfilename
    .hwndOwner = MainForm.hWnd
    .hInstance = App.hInstance
    .lpstrTitle = "Open Map File"
    .lpstrInitialDir = App.Path
    .lpstrFilter = "Map Files" & Chr$(0) & "*.map*" & Chr$(0)
    .nFilterIndex = 1
    .lpstrDefExt = "map"
    .lpstrFile = String(MAX_BUFFER_LENGTH, 0)
    .nMaxFile = MAX_BUFFER_LENGTH - 1
    .lpstrFileTitle = .lpstrFile
    .nMaxFileTitle = MAX_BUFFER_LENGTH
    .lStructSize = Len(pOpenfilename)
End With
rc = GetOpenFileName(pOpenfilename)
If rc Then
    MrClean = True
    strMapName = Left$(pOpenfilename.lpstrFile, pOpenfilename.nMaxFile)
    CurSave = strMapName
    intFreeFile = FreeFile
    TileForm.OpenIt strMapName, intFreeFile
DXMod.DrawTileSet
Else
    Me.Caption = "Open Canceled"
End If
If MrClean Then
TileForm.Refresh
DXMod.DrawALLTiles
DXMod.DrawNPConForm
End If
End Sub

Private Sub mnuFileSave_Click()
Dim intFreeFile As Integer
Dim strMapName As String
If CurSave <> "" Then
    intFreeFile = FreeFile
    TileForm.SaveIt CurSave, intFreeFile
    If MrClean Then
        TileForm.Refresh
        DXMod.DrawALLTiles
    End If
    MsgBox "Map Saved!", vbOKOnly
Else: mnuFileSaveAs_Click
End If
If MrClean Then
TileForm.Refresh
DXMod.DrawALLTiles
End If
End Sub

Private Sub mnuFileSaveAs_Click()
Dim rc As Long
Dim pOpenfilename As OPENFILENAME
Dim intFreeFile As Integer
Dim strMapName As String
Dim intXCounter As Byte
Dim intYCounter As Byte
Const MAX_BUFFER_LENGTH = 256
With pOpenfilename
    .hwndOwner = MainForm.hWnd
    .hInstance = App.hInstance
    .lpstrTitle = "Save As"
    .lpstrInitialDir = App.Path
    .lpstrFilter = "Map Files" & Chr$(0) & "*.map*" & Chr$(0)
    .nFilterIndex = 1
    .lpstrDefExt = "map"
    .lpstrFile = String(MAX_BUFFER_LENGTH, 0)
    .nMaxFile = MAX_BUFFER_LENGTH - 1
    .lpstrFileTitle = .lpstrFile
    .nMaxFileTitle = MAX_BUFFER_LENGTH
    .lStructSize = Len(pOpenfilename)
End With
rc = GetSaveFileName(pOpenfilename)
If rc Then
    strMapName = Left$(pOpenfilename.lpstrFile, pOpenfilename.nMaxFile)
    CurSave = strMapName
    intFreeFile = FreeFile
    TileForm.SaveIt strMapName, intFreeFile
    If MrClean Then
        TileForm.Refresh
        DXMod.DrawALLTiles
    End If
    MsgBox "Map Saved!", vbOKOnly
Else
    Me.Caption = "Canceled Save As"
End If
If MrClean Then
TileForm.Refresh
DXMod.DrawALLTiles
End If
End Sub

Private Sub mnuObjectLoad_Click()
Dim rc As Long
Dim pOpenfilename As OPENFILENAME
Dim strMapName As String
Dim intFreeFile As Integer
Dim TempTile As Byte
Const MAX_BUFFER_LENGTH = 256
If CurSave3 <> "" Then Exit Sub
With pOpenfilename
    .hwndOwner = MainForm.hWnd
    .hInstance = App.hInstance
    .lpstrTitle = "Open Object File"
    .lpstrInitialDir = App.Path
    .lpstrFilter = "Object Files" & Chr$(0) & "*.mof*" & Chr$(0)
    .nFilterIndex = 1
    .lpstrDefExt = "mof"
    .lpstrFile = String(MAX_BUFFER_LENGTH, 0)
    .nMaxFile = MAX_BUFFER_LENGTH - 1
    .lpstrFileTitle = .lpstrFile
    .nMaxFileTitle = MAX_BUFFER_LENGTH
    .lStructSize = Len(pOpenfilename)
End With
rc = GetOpenFileName(pOpenfilename)
If rc Then
    strMapName = Left$(pOpenfilename.lpstrFile, pOpenfilename.nMaxFile)
    CurSave3 = strMapName
    ObjectForm.OpenObjFile strMapName
    DXMod.DrawTileSet
Else
    Me.Caption = "Open Canceled"
End If
If MrClean Then
TileForm.Refresh
DXMod.DrawALLTiles
DXMod.DrawNPConForm
End If
End Sub

Private Sub mnuServerOpen_Click()
Dim rc As Long
Dim pOpenfilename As OPENFILENAME
Dim strMapName As String
Dim intFreeFile As Integer
Dim TempTile As Byte
Const MAX_BUFFER_LENGTH = 256
mnuObjectLoad_Click
With pOpenfilename
    .hwndOwner = MainForm.hWnd
    .hInstance = App.hInstance
    .lpstrTitle = "Open  Server Map File"
    .lpstrInitialDir = App.Path
    .lpstrFilter = "Server Map Files" & Chr$(0) & "*.smf*" & Chr$(0)
    .nFilterIndex = 1
    .lpstrDefExt = "smf"
    .lpstrFile = String(MAX_BUFFER_LENGTH, 0)
    .nMaxFile = MAX_BUFFER_LENGTH - 1
    .lpstrFileTitle = .lpstrFile
    .nMaxFileTitle = MAX_BUFFER_LENGTH
    .lStructSize = Len(pOpenfilename)
End With
rc = GetOpenFileName(pOpenfilename)
If rc Then
    MrClean = True
    strMapName = Left$(pOpenfilename.lpstrFile, pOpenfilename.nMaxFile)
    CurSave2 = strMapName
    intFreeFile = FreeFile
    TileForm.ServerOpenIt strMapName, intFreeFile
DXMod.DrawTileSet
Else
    Me.Caption = "Open Canceled"
End If
If MrClean Then
TileForm.Refresh
DXMod.DrawALLTiles
DXMod.DrawNPConForm
End If
End Sub

Private Sub mnuServerSave_Click()
Dim intFreeFile As Integer
Dim strMapName As String
Dim intXCounter As Byte
Dim intYCounter As Byte
If CurSave2 <> "" Then
    intFreeFile = FreeFile
    TileForm.ServerSaveIt CurSave2, intFreeFile
    If MrClean Then
        TileForm.Refresh
        DXMod.DrawALLTiles
    End If
    MsgBox "Map Saved!", vbOKOnly
Else: mnuServerSaveAs_Click
End If
If MrClean Then
TileForm.Refresh
DXMod.DrawALLTiles
End If
End Sub

Private Sub mnuServerSaveAs_Click()
Dim rc As Long
Dim pOpenfilename As OPENFILENAME
Dim intFreeFile As Integer
Dim strMapName As String
Dim intXCounter As Byte
Dim intYCounter As Byte
Const MAX_BUFFER_LENGTH = 256
With pOpenfilename
    .hwndOwner = MainForm.hWnd
    .hInstance = App.hInstance
    .lpstrTitle = "Save As"
    .lpstrInitialDir = App.Path
    .lpstrFilter = "Server Map Files" & Chr$(0) & "*.smf*" & Chr$(0)
    .nFilterIndex = 1
    .lpstrDefExt = "smf"
    .lpstrFile = String(MAX_BUFFER_LENGTH, 0)
    .nMaxFile = MAX_BUFFER_LENGTH - 1
    .lpstrFileTitle = .lpstrFile
    .nMaxFileTitle = MAX_BUFFER_LENGTH
    .lStructSize = Len(pOpenfilename)
End With
rc = GetSaveFileName(pOpenfilename)
If rc Then
    strMapName = Left$(pOpenfilename.lpstrFile, pOpenfilename.nMaxFile)
    CurSave2 = strMapName
    intFreeFile = FreeFile
    TileForm.ServerSaveIt strMapName, intFreeFile
    If MrClean Then
        TileForm.Refresh
        DXMod.DrawALLTiles
    End If
    MsgBox "Map Saved!", vbOKOnly
Else
    Me.Caption = "Canceled Save As"
End If
If MrClean Then
TileForm.Refresh
DXMod.DrawALLTiles
End If
End Sub
