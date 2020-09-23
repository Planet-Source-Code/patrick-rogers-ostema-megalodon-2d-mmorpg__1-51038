Attribute VB_Name = "DXMod"
Dim DX As New DirectX7
Dim DD As DirectDraw7
Dim picBuffer As DirectDrawSurface7
Dim Primary As DirectDrawSurface7
Dim Secondary As DirectDrawSurface7
Dim ddsd As DDSURFACEDESC2
Dim ddClipper As DirectDrawClipper
Dim WinRect As RECT
Option Explicit
Public Sub Init()
Dim ScreenPropDesc As DDSURFACEDESC2
On Error GoTo ErrHandler:
    Set DD = DX.DirectDrawCreate("")
    Call DD.SetCooperativeLevel(FrmEdit.hWnd, DDSCL_NORMAL)
    Call DD.GetDisplayMode(ScreenPropDesc)
    ddsd.lFlags = DDSD_CAPS
    'This surface is the primary surface (what is visible to the user)
    ddsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    'You're now creating the primary surface with the surface description you just set
    Set Primary = DD.CreateSurface(ddsd)
    'Now let's set the second surface description
    ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    'This is going to be a plain off-screen surface - ie, to hold a bitmap
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd.lHeight = 864
    ddsd.lWidth = 32
    'Now we create the off-screen surface from the pre-rendered picture
    Set picBuffer = DD.CreateSurfaceFromFile(App.Path & "\Items.bmp", ddsd)
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd.lWidth = ScreenPropDesc.lWidth
    ddsd.lHeight = ScreenPropDesc.lHeight
    Set Secondary = DD.CreateSurface(ddsd)
    Set ddClipper = DD.CreateClipper(0)
    ddClipper.SetHWnd FrmEdit.ObjPic.hWnd
    Primary.SetClipper ddClipper
Exit Sub
ErrHandler:
MsgBox "Unable to initialize DirectDraw - Closing program", vbInformation, "error"
End
End Sub
Public Sub KillDX()
Call Secondary.BltColorFill(WinRect, 0)
Set picBuffer = Nothing
Set Primary = Nothing
Set Secondary = Nothing
Set DD = Nothing
Set DX = Nothing
End Sub
Public Sub DrawObject()
Dim TempSDesc As DDSURFACEDESC2
Dim BlitRect As RECT
Dim WinRect As RECT
Dim FullRect As RECT
Dim DestRect As RECT
    Primary.SetClipper ddClipper
    'Gets the bounding rect for the entire window handle, stores in r1
    Call DX.GetWindowRect(FrmEdit.ObjPic.hWnd, WinRect)
    Call Secondary.GetSurfaceDesc(TempSDesc)
    FullRect.Right = TempSDesc.lWidth
    FullRect.Bottom = TempSDesc.lHeight
    With BlitRect
    .Left = 0
    .Right = 32
    .Top = (FrmEdit.VScroll1.Value * 32)
    .Bottom = .Top + 32
    End With
    With DestRect
    .Left = WinRect.Left
    .Top = WinRect.Top
    End With
    Call Secondary.BltFast(DestRect.Left, DestRect.Top, picBuffer, BlitRect, DDBLTFAST_WAIT)
    Call Primary.Blt(FullRect, Secondary, FullRect, DDBLT_WAIT)
End Sub
