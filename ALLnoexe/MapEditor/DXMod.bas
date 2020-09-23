Attribute VB_Name = "DXMod"
Dim DX As New DirectX7
Dim DD As DirectDraw7
Dim picBuffer As DirectDrawSurface7
Dim BodyBuf As DirectDrawSurface7
Dim HeadBuf As DirectDrawSurface7
Dim ItemBuf As DirectDrawSurface7
Dim Primary As DirectDrawSurface7
Dim Secondary As DirectDrawSurface7
Dim ddsd1 As DDSURFACEDESC2
Dim ddsd2 As DDSURFACEDESC2
Dim ddsd3 As DDSURFACEDESC2
Dim ScreenPropDesc As DDSURFACEDESC2
Dim ddClipper As DirectDrawClipper
Dim ddClipper2 As DirectDrawClipper
Dim ddClipper3 As DirectDrawClipper
Dim ddClipper4 As DirectDrawClipper
Dim WinRect As RECT
Option Explicit
Public Sub Init()
Dim AColorKey As DDCOLORKEY
Dim ScreenPropDesc As DDSURFACEDESC2
AColorKey.high = vbBlack
AColorKey.low = vbBlack
    Set DD = DX.DirectDrawCreate("")
    Call DD.SetCooperativeLevel(MainForm.hWnd, DDSCL_NORMAL)
    Call DD.GetDisplayMode(ScreenPropDesc)
    'Indicate that the ddsCaps member is valid in this type
    ddsd1.lFlags = DDSD_CAPS
    'This surface is the primary surface (what is visible to the user)
    ddsd1.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    'You're now creating the primary surface with the surface description you just set
    Set Primary = DD.CreateSurface(ddsd1)
    'Now let's set the second surface description
    ddsd2.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    'This is going to be a plain off-screen surface - ie, to hold a bitmap
    ddsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd2.lHeight = 768
    ddsd2.lWidth = 512
    'Now we create the off-screen surface from the pre-rendered picture
    Set picBuffer = DD.CreateSurfaceFromFile(App.Path & "\tileset.bmp", ddsd2)
    picBuffer.SetColorKey DDCKEY_SRCBLT, AColorKey
    ddsd2.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    'This is going to be a plain off-screen surface - ie, to hold a bitmap
    ddsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd2.lHeight = 864
    ddsd2.lWidth = 32
    Set ItemBuf = DD.CreateSurfaceFromFile(App.Path & "\Items.bmp", ddsd2)
    ItemBuf.SetColorKey DDCKEY_SRCBLT, AColorKey
    ddsd2.lHeight = 160
    ddsd2.lWidth = 128
    Set HeadBuf = DD.CreateSurfaceFromFile(App.Path & "\Head.bmp", ddsd2)
    HeadBuf.SetColorKey DDCKEY_SRCBLT, AColorKey
    ddsd3.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd2.lHeight = 160
    ddsd2.lWidth = 768
    Set BodyBuf = DD.CreateSurfaceFromFile(App.Path & "\Body.bmp", ddsd2)
    BodyBuf.SetColorKey DDCKEY_SRCBLT, AColorKey
    'This is going to be a plain off-screen surface - ie, to hold a bitmap
    ddsd3.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd3.lWidth = ScreenPropDesc.lWidth
    ddsd3.lHeight = ScreenPropDesc.lHeight
    Set Secondary = DD.CreateSurface(ddsd3)
    
    Set ddClipper = DD.CreateClipper(0)
    ddClipper.SetHWnd TileForm.hWnd
    Set ddClipper2 = DD.CreateClipper(0)
    ddClipper2.SetHWnd PlaneForm.picTileSelect.hWnd
    Set ddClipper3 = DD.CreateClipper(0)
    ddClipper3.SetHWnd NPCForm.Picture1.hWnd
    Set ddClipper4 = DD.CreateClipper(0)
    ddClipper4.SetHWnd ObjectForm.Picture1.hWnd
    Primary.SetClipper ddClipper2
    
Exit Sub
ErrHandler:
MsgBox "Unable to initialize DirectDraw - Closing program", vbInformation, "error"
End
End Sub
Public Sub KillDX()
Call Secondary.BltColorFill(WinRect, 0)
Set picBuffer = Nothing
Set HeadBuf = Nothing
Set BodyBuf = Nothing
Set ItemBuf = Nothing
Set Primary = Nothing
Set Secondary = Nothing
Set DD = Nothing
Set DX = Nothing
End Sub
Public Sub DrawTileSet()
Dim TempSDesc As DDSURFACEDESC2
Dim BlitRect As RECT
Dim WinRect As RECT
Dim FullRect As RECT
Dim DestRect As RECT
    Primary.SetClipper ddClipper2
    'Gets the bounding rect for the entire window handle, stores in r1
    Call DX.GetWindowRect(PlaneForm.picTileSelect.hWnd, WinRect)
    Call Secondary.GetSurfaceDesc(TempSDesc)
    FullRect.Right = TempSDesc.lWidth
    FullRect.Bottom = TempSDesc.lHeight
    With BlitRect
    .Left = (PlaneForm.HScroll1.Value * 32)
    .Right = .Left + 160
    .Top = (PlaneForm.VScroll1.Value * 32)
    .Bottom = .Top + 160
    End With
    With DestRect
    .Left = WinRect.Left
    .Top = WinRect.Top
    End With
    Call Secondary.BltColorFill(WinRect, 0)
    Call Secondary.BltFast(DestRect.Left, DestRect.Top, picBuffer, BlitRect, DDBLTFAST_WAIT)
    With BlitRect
    .Left = 320
    .Right = 352
    .Top = 0
    .Bottom = 32
    End With
    With DestRect
    .Left = WinRect.Left + (PlaneForm.TileX - PlaneForm.HScroll1.Value) * 32
    .Right = .Left + 32
    .Top = WinRect.Top + (PlaneForm.TileY - PlaneForm.VScroll1.Value) * 32
    .Bottom = .Top + 32
    End With
    If DestRect.Left >= WinRect.Left Then
        Call Secondary.BltFast(DestRect.Left, DestRect.Top, picBuffer, BlitRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    End If
    Call Primary.Blt(FullRect, Secondary, FullRect, DDBLT_WAIT)
End Sub
Public Sub DrawTile(X As Byte, Y As Byte)
Dim TempSDesc As DDSURFACEDESC2
Dim BlitRect As RECT
Dim WinRect As RECT
Dim FullRect As RECT
Dim DestRect As RECT
    Primary.SetClipper ddClipper
    'Gets the bounding rect for the entire window handle, stores in r1
    Call DX.GetWindowRect(TileForm.hWnd, WinRect)
    Call Secondary.GetSurfaceDesc(TempSDesc)
    FullRect.Right = TempSDesc.lWidth
    FullRect.Bottom = TempSDesc.lHeight
    'Draw first layer
    With BlitRect
    .Left = (TileForm.GetTileX(X, Y) * 32)
    .Right = .Left + 32
    .Top = (TileForm.GetTileY(X, Y) * 32)
    .Bottom = .Top + 32
    End With
    With DestRect
    .Left = WinRect.Left + (X - TileForm.HScroll1.Value) * 32
    .Top = WinRect.Top + (Y - TileForm.VScroll1.Value) * 32
    End With
    Call Secondary.BltFast(DestRect.Left + 3, DestRect.Top + 22, picBuffer, BlitRect, DDBLTFAST_WAIT)
    'Draw second layer
    If TileForm.GetTileFX(X, Y) <> 0 Or TileForm.GetTileFY(X, Y) <> 0 Then
    With BlitRect
    .Left = (TileForm.GetTileFX(X, Y) * 32)
    .Right = .Left + 32
    .Top = (TileForm.GetTileFY(X, Y) * 32)
    .Bottom = .Top + 32
    End With
    Call Secondary.BltFast(DestRect.Left + 3, DestRect.Top + 22, picBuffer, BlitRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    End If
    'draw third layer
    If TileForm.GetTileSX(X, Y) <> 0 Or TileForm.GetTileSY(X, Y) <> 0 Then
    With BlitRect
    .Left = (TileForm.GetTileSX(X, Y) * 32)
    .Right = .Left + 32
    .Top = (TileForm.GetTileSY(X, Y) * 32)
    .Bottom = .Top + 32
    End With
    Call Secondary.BltFast(DestRect.Left + 3, DestRect.Top + 22, picBuffer, BlitRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    End If
    Call Primary.Blt(FullRect, Secondary, FullRect, DDBLT_WAIT)
End Sub
Public Sub DrawALLTiles()
Dim TempSDesc As DDSURFACEDESC2
Dim BlitRect As RECT
Dim WinRect As RECT
Dim FullRect As RECT
Dim DestRect As RECT
Dim TempX As Byte
Dim TempY As Byte
    Primary.SetClipper ddClipper
    'Gets the bounding rect for the entire window handle, stores in r1
    Call DX.GetWindowRect(TileForm.hWnd, WinRect)
    Call Secondary.GetSurfaceDesc(TempSDesc)
    FullRect.Right = TempSDesc.lWidth
    FullRect.Bottom = TempSDesc.lHeight
    For TempX = 0 To 17
        For TempY = 0 To 14
            With BlitRect
            .Left = (TileForm.GetTileX(TempX + TileForm.HScroll1.Value, TempY + TileForm.VScroll1.Value)) * 32
            .Right = .Left + 32
            .Top = (TileForm.GetTileY(TempX + TileForm.HScroll1.Value, TempY + TileForm.VScroll1.Value)) * 32
            .Bottom = .Top + 32
            End With
            With DestRect
            .Left = WinRect.Left + (TempX * 32)
            .Top = WinRect.Top + (TempY * 32)
            End With
            Call Secondary.BltFast(DestRect.Left + 3, DestRect.Top + 22, picBuffer, BlitRect, DDBLTFAST_WAIT)
            'Draw second layer
            If TileForm.GetTileFX(TempX + TileForm.HScroll1.Value, TempY + TileForm.VScroll1.Value) <> 0 Or TileForm.GetTileFY(TempX + TileForm.HScroll1.Value, TempY + TileForm.VScroll1.Value) <> 0 Then
            With BlitRect
            .Left = (TileForm.GetTileFX(TempX + TileForm.HScroll1.Value, TempY + TileForm.VScroll1.Value) * 32)
            .Right = .Left + 32
            .Top = (TileForm.GetTileFY(TempX + TileForm.HScroll1.Value, TempY + TileForm.VScroll1.Value) * 32)
            .Bottom = .Top + 32
            End With
            Call Secondary.BltFast(DestRect.Left + 3, DestRect.Top + 22, picBuffer, BlitRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
            End If
            'Draw Ground Items
            If TileForm.GetMapItem(TempX + TileForm.HScroll1.Value, TempY + TileForm.VScroll1.Value) > -1 Then
                If TileForm.GetTileProp(TempX + TileForm.HScroll1.Value, TempY + TileForm.VScroll1.Value) = 1 Then
                    With BlitRect
                        .Left = 7 * 32
                        .Right = .Left + 32
                        .Top = 8 * 32
                        .Bottom = .Top + 32
                    End With
                    Call Secondary.BltFast(DestRect.Left + 3, DestRect.Top + 22, picBuffer, BlitRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
                    GoTo SkipAge
                End If
            With BlitRect
            .Left = 0
            .Right = 32
            .Top = ObjectForm.GetItemGraphic(TileForm.GetMapItem(TempX + TileForm.HScroll1.Value, TempY + TileForm.VScroll1.Value)) * 32
            .Bottom = .Top + 32
            End With
            Call Secondary.BltFast(DestRect.Left + 3, DestRect.Top + 22, ItemBuf, BlitRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
            End If
            'Draw NPC if there is one
            If TileForm.GetNPCIndex(TempX + TileForm.HScroll1.Value, TempY + TileForm.VScroll1.Value) > -1 Then
            With BlitRect
            .Left = 96
            .Right = 128
            .Top = TileForm.GetNPCB(TempX + TileForm.HScroll1.Value, TempY + TileForm.VScroll1.Value) * 32
            .Bottom = .Top + 32
            End With
            Call Secondary.BltFast(DestRect.Left + 3, DestRect.Top + 22, BodyBuf, BlitRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
            With BlitRect
            .Left = 32
            .Right = 64
            .Top = TileForm.GetNPCH(TempX + TileForm.HScroll1.Value, TempY + TileForm.VScroll1.Value) * 32
            .Bottom = .Top + 32
            End With
            Call Secondary.BltFast(DestRect.Left + 3, DestRect.Top + 22, HeadBuf, BlitRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
            End If
            'draw third layer
            If TileForm.GetTileSX(TempX + TileForm.HScroll1.Value, TempY + TileForm.VScroll1.Value) <> 0 Or TileForm.GetTileSY(TempX + TileForm.HScroll1.Value, TempY + TileForm.VScroll1.Value) <> 0 Then
            With BlitRect
            .Left = (TileForm.GetTileSX(TempX + TileForm.HScroll1.Value, TempY + TileForm.VScroll1.Value) * 32)
            .Right = .Left + 32
            .Top = (TileForm.GetTileSY(TempX + TileForm.HScroll1.Value, TempY + TileForm.VScroll1.Value) * 32)
            .Bottom = .Top + 32
            End With
            Call Secondary.BltFast(DestRect.Left + 3, DestRect.Top + 22, picBuffer, BlitRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
            End If
            
SkipAge:
        Next
    Next
Call Primary.Blt(FullRect, Secondary, FullRect, DDBLT_WAIT)
End Sub
Public Sub DrawNPConForm()
Dim TempSDesc As DDSURFACEDESC2
Dim BlitRect As RECT
Dim WinRect As RECT
Dim FullRect As RECT
Dim DestRect As RECT
    Primary.SetClipper ddClipper3
    'Gets the bounding rect for the entire window handle, stores in r1
    Call DX.GetWindowRect(NPCForm.Picture1.hWnd, WinRect)
    Call Secondary.GetSurfaceDesc(TempSDesc)
    FullRect.Right = TempSDesc.lWidth
    FullRect.Bottom = TempSDesc.lHeight
    With BlitRect
    .Left = 96
    .Right = 128
    .Top = (NPCForm.VScroll1.Value * 32)
    .Bottom = .Top + 32
    End With
    With DestRect
    .Left = WinRect.Left
    .Top = WinRect.Top
    End With
    Call Secondary.BltColorFill(WinRect, 0)
    Call Secondary.BltFast(DestRect.Left, DestRect.Top, BodyBuf, BlitRect, DDBLTFAST_WAIT)
    With BlitRect
    .Left = 32
    .Right = 64
    .Top = (NPCForm.VScroll2.Value * 32)
    .Bottom = .Top + 32
    End With
    Call Secondary.BltFast(DestRect.Left, DestRect.Top, HeadBuf, BlitRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    Call Primary.Blt(FullRect, Secondary, FullRect, DDBLT_WAIT)
End Sub
Public Sub DrawItem()
Dim TempSDesc As DDSURFACEDESC2
Dim BlitRect As RECT
Dim WinRect As RECT
Dim FullRect As RECT
Dim DestRect As RECT
    Primary.SetClipper ddClipper4
    'Gets the bounding rect for the entire window handle, stores in r1
    Call DX.GetWindowRect(ObjectForm.Picture1.hWnd, WinRect)
    Call Secondary.GetSurfaceDesc(TempSDesc)
    FullRect.Right = TempSDesc.lWidth
    FullRect.Bottom = TempSDesc.lHeight
    With BlitRect
    .Left = 0
    .Right = 32
    .Top = (ObjectForm.GetItemGraphic(ObjectForm.ObjCombo.ListIndex) * 32)
    .Bottom = .Top + 32
    End With
    With DestRect
    .Left = WinRect.Left
    .Top = WinRect.Top
    End With
    Call Secondary.BltFast(DestRect.Left, DestRect.Top, ItemBuf, BlitRect, DDBLTFAST_WAIT)
    Call Primary.Blt(FullRect, Secondary, FullRect, DDBLT_WAIT)
End Sub
