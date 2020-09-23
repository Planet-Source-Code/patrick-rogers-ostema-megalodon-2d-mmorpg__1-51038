VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "Mega's MMORPG Object Editor"
   ClientHeight    =   2805
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4875
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuFile 
      Caption         =   "File"
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
   Begin VB.Menu mnuObject 
      Caption         =   "Object"
      Begin VB.Menu mnuObjectAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuObjectRemove 
         Caption         =   "Remove"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (lpOpenFilename As OPENFILENAME) As Long
'Megalodon MMORPG Object Editor
'http://home.kc.rr.com/megalodonsoft
'Use this code however u like just make
'it open source please so people can learn
'just like you did:)
'~Patrick Rogers-Ostema
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (lpOpenFilename As OPENFILENAME) As Long
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
Private Sub MDIForm_Load()
MDIForm_Resize
DXMod.Init
DXMod.DrawObject
FrmList.ObjList.AddItem "Add Object"
End Sub

Private Sub MDIForm_Resize()
If Me.Width > 2665 Then
FrmList.Left = 0
FrmList.Width = Me.Width - 2665
FrmList.Top = 0
FrmList.Height = Me.Height - 750
FrmList.ObjList.Top = 0
FrmList.ObjList.Height = (FrmList.ScaleHeight) + 8
FrmList.ObjList.Left = 0
FrmList.ObjList.Width = (FrmList.ScaleWidth)
FrmEdit.Left = FrmList.Width
FrmEdit.Width = 2505
FrmEdit.Top = 0
FrmEdit.Height = Me.Height - 750
End If
End Sub
Public Function ExitProg() As Integer
If MsgBox("Exit Megalodon MMORPG Object Editor?", vbYesNo) = vbYes Then
ExitProg = 0
DXMod.KillDX
End
Else
ExitProg = 1
End If
End Function

Private Sub MDIForm_Unload(Cancel As Integer)
Cancel = ExitProg()
End Sub

Private Sub mnuFileOpen_Click()
Dim rc As Long
Dim pOpenfilename As OPENFILENAME
Dim strMapName As String
Dim intFreeFile As Integer
Dim TempTile As Byte
Const MAX_BUFFER_LENGTH = 256
With pOpenfilename
    .hwndOwner = MDIMain.hWnd
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
    CurSave = strMapName
    FrmList.OpenIt strMapName
Else
    Me.Caption = "Open Canceled"
End If
End Sub

Private Sub mnuFileSave_Click()
If CurSave <> "" Then
    FrmList.SaveIt CurSave
    MsgBox "Objects Saved!", vbOKOnly
Else: mnuFileSaveAs_Click
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
    .hwndOwner = MDIMain.hWnd
    .hInstance = App.hInstance
    .lpstrTitle = "Save As"
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
rc = GetSaveFileName(pOpenfilename)
If rc Then
    strMapName = Left$(pOpenfilename.lpstrFile, pOpenfilename.nMaxFile)
    CurSave = strMapName
    FrmList.SaveIt strMapName
    MsgBox "Object Saved!", vbOKOnly
Else
    Me.Caption = "Canceled Save As"
End If
End Sub

Private Sub mnuObjectAdd_Click()
MsgBox "Double click on Add Object in the list box."
End Sub

Private Sub mnuObjectRemove_Click()
FrmList.RemoveObject FrmList.ObjList.ListIndex
End Sub
