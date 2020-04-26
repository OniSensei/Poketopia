Attribute VB_Name = "modDirectDraw7"
Option Explicit
' **********************
' ** Renders graphics **
' **********************
' DirectDraw7 Object
Public DD As DirectDraw7
' Clipper object
Public DD_Clip As DirectDrawClipper

' primary surface
Public DDS_Primary As DirectDrawSurface7
Public DDSD_Primary As DDSURFACEDESC2

' back buffer
Public DDS_BackBuffer As DirectDrawSurface7
Public DDSD_BackBuffer As DDSURFACEDESC2

' Used for pre-rendering paperdolls
Public DDS_Player() As DirectDrawSurface7
Public DDSD_Player() As DDSURFACEDESC2

'GDI Images
Public GDIImage(1 To 255) As GDIpImage
Public GDILoaded(1 To 255) As Boolean
Public PokeImg As GDIpImage
Public EnemyPokeImg As GDIpImage
Public BattleImg As GDIpImage
Public battleBaseImg As GDIpImage

Public PokeImgLoaded As Boolean
Public EnemyPokeImgLoaded As Boolean
Public BattleImgLoaded As Boolean
Public battleBaseImgloaded As Boolean

' gfx buffers
Public DDS_Item() As DirectDrawSurface7
Public DDS_Character() As DirectDrawSurface7
Public DDS_Paperdoll() As DirectDrawSurface7
Public DDS_Tileset() As DirectDrawSurface7
Public DDS_Misc As DirectDrawSurface7
Public DDS_Resource() As DirectDrawSurface7
Public DDS_Door As DirectDrawSurface7
Public DDS_Blood As DirectDrawSurface7
Public DDS_Animation() As DirectDrawSurface7
Public DDS_SpellIcon() As DirectDrawSurface7
Public DDS_DownFrame() As DirectDrawSurface7
Public DDS_DownFrame2() As DirectDrawSurface7
Public DDS_UpFrame() As DirectDrawSurface7
Public DDS_UpFrame2() As DirectDrawSurface7
Public DDS_LeftFrame() As DirectDrawSurface7
Public DDS_LeftFrame2() As DirectDrawSurface7
Public DDS_RightFrame() As DirectDrawSurface7
Public DDS_RightFrame2() As DirectDrawSurface7
Public DDS_Weather As DirectDrawSurface7
Public DDS_Menu As DirectDrawSurface7
Public DDS_Chat As DirectDrawSurface7
Public DDS_Frame As DirectDrawSurface7
' descriptions
Public DDSD_Temp As DDSURFACEDESC2
Public DDSD_Item() As DDSURFACEDESC2
Public DDSD_Character() As DDSURFACEDESC2
Public DDSD_Paperdoll() As DDSURFACEDESC2
Public DDSD_Tileset() As DDSURFACEDESC2
Public DDSD_Misc As DDSURFACEDESC2
Public DDSD_Resource() As DDSURFACEDESC2
Public DDSD_Door As DDSURFACEDESC2
Public DDSD_Blood As DDSURFACEDESC2
Public DDSD_Animation() As DDSURFACEDESC2
Public DDSD_SpellIcon() As DDSURFACEDESC2
Public DDSD_DownFrame() As DDSURFACEDESC2
Public DDSD_DownFrame2() As DDSURFACEDESC2
Public DDSD_UpFrame() As DDSURFACEDESC2
Public DDSD_UpFrame2() As DDSURFACEDESC2
Public DDSD_LeftFrame() As DDSURFACEDESC2
Public DDSD_LeftFrame2() As DDSURFACEDESC2
Public DDSD_RightFrame() As DDSURFACEDESC2
Public DDSD_RightFrame2() As DDSURFACEDESC2
Public DDSD_Weather As DDSURFACEDESC2
Public DDSD_Menu As DDSURFACEDESC2
Public DDSD_Chat As DDSURFACEDESC2
Public DDSD_Frame As DDSURFACEDESC2
' timers
Public Const SurfaceTimerMax As Long = 10000
Public CharacterTimer() As Long
Public PaperdollTimer() As Long
Public ItemTimer() As Long
Public ResourceTimer() As Long
Public DoorTimer As Long
Public BloodTimer As Long
Public AnimationTimer() As Long
Public SpellIconTimer() As Long

' Number of graphic files
Public NumTileSets As Long
Public NumCharacters As Long
Public NumPaperdolls As Long
Public NumItems As Long
Public NumResources As Long
Public NumAnimations As Long
Public NumSpellIcons As Long

' ********************
' ** Initialization **
' ********************
Public Function InitDirectDraw() As Boolean

    On Error GoTo ErrorHandle

    Call DestroyDirectDraw ' clear out everything
    ' Initialize direct draw
    Set DD = DX7.DirectDrawCreate(vbNullString) ' empty string forces primary device
    ' dictates how we access the screen and how other programs
    ' running at the same time will be allowed to access the screen as well.
    Call DD.SetCooperativeLevel(frmMainGame.hwnd, DDSCL_NORMAL)

    ' Init type and set the primary surface
    With DDSD_Primary
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End With

    Set DDS_Primary = DD.CreateSurface(DDSD_Primary)
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    ' Associate the picture hwnd with the clipper
    Call DD_Clip.SetHWnd(frmMainGame.picScreen.hwnd)
    ' Have the blits to the screen clipped to the picture box
    Call DDS_Primary.SetClipper(DD_Clip) ' method attaches a clipper object to, or deletes one from, a surface.
    
    Call InitBackBuffer
    
    InitDirectDraw = True
    
    Exit Function
ErrorHandle:

    Select Case Err.number
        Case 91
            Call MsgBox("DirectX7 master object not created.")
    End Select

    InitDirectDraw = False
End Function

Private Sub InitBackBuffer()
    Dim rec As DxVBLib.RECT

    ' Initialize back buffer
    With DDSD_BackBuffer
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
        .lWidth = (MAX_MAPX + 3) * PIC_X
        .lHeight = (MAX_MAPY + 3) * PIC_Y
    End With

    Set DDS_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
    Call DDS_BackBuffer.BltColorFill(rec, 0)
End Sub

' This sub gets the mask color from the surface loaded from a bitmap image
Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal x As Long, ByVal y As Long)
    Dim TmpR As RECT
    Dim TmpDDSD As DDSURFACEDESC2
    Dim TmpColorKey As DDCOLORKEY

    With TmpR
        .Left = x
        .Top = y
        .Right = x
        .Bottom = y
    End With

    TheSurface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0

    With TmpColorKey
        .Low = TheSurface.GetLockedPixel(x, y)
        .High = .Low
    End With

    TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey
    TheSurface.Unlock TmpR
End Sub

' Initializing a surface, using a bitmap
Public Sub InitDDSurf(FileName As String, ByRef SurfDesc As DDSURFACEDESC2, ByRef Surf As DirectDrawSurface7)

    On Error GoTo ErrorHandle

    ' Set path
    FileName = App.Path & GFX_PATH & FileName & GFX_EXT

    ' Destroy surface if it exist
    If Not Surf Is Nothing Then
        Set Surf = Nothing
        Call ZeroMemory(ByVal VarPtr(SurfDesc), LenB(SurfDesc))
    End If

    ' set flags
    SurfDesc.lFlags = DDSD_CAPS
    SurfDesc.ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
    ' init object
    Set Surf = DD.CreateSurfaceFromFile(FileName, SurfDesc)
    'Call Surf.SetColorKey(DDCKEY_SRCBLT, Key) ' MASK_COLOR
    Call SetMaskColorFromPixel(Surf, 0, 0)
    Exit Sub
ErrorHandle:

    Select Case Err.number
            ' File not found
        Case 53
            MsgBox "missing file: " & FileName
            Call DestroyGame
            ' DirectDraw does not have enough memory to perform the operation.
        Case DDERR_OUTOFMEMORY
            MsgBox "Out of system memory"
            Call DestroyGame
            ' DirectDraw does not have enough display memory to perform the operation.
        Case DDERR_OUTOFVIDEOMEMORY
            MsgBox "Out of video memory, attempting to re-initialize using system memory"
            DDSD_Temp.lFlags = DDSD_CAPS
            DDSD_Temp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
            Call ReInitDD
    End Select

End Sub

Public Function CheckSurfaces() As Boolean

    On Error GoTo ErrorHandle

    ' Check if we need to restore surfaces
    If NeedToRestoreSurfaces Then
        DD.RestoreAllSurfaces
    End If

    CheckSurfaces = True
    Exit Function
ErrorHandle:
    Call ReInitDD
    CheckSurfaces = False
End Function

Private Function NeedToRestoreSurfaces() As Boolean

    If Not DD.TestCooperativeLevel = DD_OK Then
        NeedToRestoreSurfaces = True
    End If

End Function

Public Sub ReInitDD()
    Call InitDirectDraw

    If InMapEditor Then
        Call InitDDSurf("misc", DDSD_Misc, DDS_Misc)
    End If

End Sub

Public Sub DestroyDirectDraw()
    Dim i As Long
    
    ' Unload DirectDraw
    Set DDS_Misc = Nothing
    
    For i = 1 To NumTileSets
        Set DDS_Tileset(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Tileset(i)), LenB(DDSD_Tileset(i))
    Next

    For i = 1 To NumItems
        Set DDS_Item(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Item(i)), LenB(DDSD_Item(i))
    Next

    For i = 1 To NumCharacters
        Set DDS_Character(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Character(i)), LenB(DDSD_Character(i))
    Next
    
    For i = 1 To NumPaperdolls
        Set DDS_Paperdoll(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Paperdoll(i)), LenB(DDSD_Paperdoll(i))
    Next
    
    For i = 1 To NumResources
        Set DDS_Resource(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Resource(i)), LenB(DDSD_Resource(i))
    Next
    
    For i = 1 To NumAnimations
        Set DDS_Animation(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Animation(i)), LenB(DDSD_Animation(i))
    Next
    
    For i = 1 To NumSpellIcons
        Set DDS_SpellIcon(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_SpellIcon(i)), LenB(DDSD_SpellIcon(i))
    Next
    
    
    
    Set DDS_Blood = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Blood), LenB(DDSD_Blood)
    Set DDS_Door = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Door), LenB(DDSD_Door)
    
    Set DDS_Weather = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Weather), LenB(DDSD_Weather)
    
    Set DDS_Menu = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Menu), LenB(DDSD_Menu)
    
    Set DDS_Chat = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Chat), LenB(DDSD_Chat)
    
    Set DDS_Frame = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Frame), LenB(DDSD_Frame)

    Set DDS_BackBuffer = Nothing
    Set DDS_Primary = Nothing
    Set DD_Clip = Nothing
    Set DD = Nothing
End Sub

' **************
' ** Blitting **
' **************
Public Sub Engine_BltFast(ByVal dx As Long, ByVal dy As Long, ByRef ddS As DirectDrawSurface7, srcRECT As RECT, trans As CONST_DDBLTFASTFLAGS)

    On Error GoTo ErrorHandle:

    If Not ddS Is Nothing Then
        Call DDS_BackBuffer.BltFast(dx, dy, ddS, srcRECT, trans)
    End If

    Exit Sub
ErrorHandle:

    Select Case Err.number
        Case 5
            Call DevMsg("Attempting to copy from a surface thats not initialized.", BrightRed)
    End Select

End Sub

Public Function Engine_BltToDC(ByRef Surface As DirectDrawSurface7, sRECT As DxVBLib.RECT, dRECT As DxVBLib.RECT, ByRef picbox As VB.picturebox, Optional Clear As Boolean = True) As Boolean

    On Error GoTo ErrorHandle

    If Clear Then
        picbox.Cls
    End If

    Call Surface.BltToDC(picbox.hDC, sRECT, dRECT)
    picbox.Refresh
    Engine_BltToDC = True
    Exit Function
ErrorHandle:
    ' returns false on error
    Engine_BltToDC = False
End Function

Sub DrawGDIImage(ByVal img As GDIpImage, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal useEffects As Boolean = False, Optional ByVal effect As GDIpEffects)
Dim aDC As Long
aDC = DDS_BackBuffer.GetDC
'If useEffects Then
'Dim rs As RENDERSTYLESTRUCT2
'Set rs.Effects = effect
'PaintPictureGDIplus img, aDC, X, Y, width, height, , , , , rs
'Else
'PaintPictureGDIplus img, aDC, X, Y, width, height
'End If
If img Is Nothing Then MsgBox "Error - passed Img not assigned"
img.Render aDC, x, y, Width, Height
DDS_BackBuffer.ReleaseDC aDC

End Sub




Public Sub BltMapTile(ByVal x As Long, ByVal y As Long)
    Dim rec As DxVBLib.RECT
    Dim i As Long
    If FlashLight = True Then
If x > Player(MyIndex).x + 3 Or x < Player(MyIndex).x - 3 Or y > Player(MyIndex).y + 3 Or y < Player(MyIndex).y - 3 Then Exit Sub
End If
    With map.Tile(x, y)
        For i = MapLayer.Ground To MapLayer.Mask2
            ' skip tile if tileset isn't set
            'GROUND-----------------------------------
            If i = MapLayer.Ground Then
            If InMapEditor = True And GroundUnvisible = False Then
            If .Layer(i).tileset > 0 Then
                ' sort out rec
                rec.Top = .Layer(i).y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).x * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), DDS_Tileset(.Layer(i).tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
            Else
            If InMapEditor = False Then
            If .Layer(i).tileset > 0 Then
                ' sort out rec
                rec.Top = .Layer(i).y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).x * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), DDS_Tileset(.Layer(i).tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
            End If
            End If
            End If
            'MASK-------------------------------------
             If i = MapLayer.mask Then
            If InMapEditor = True And MaskUnvisible = False Then
            If .Layer(i).tileset > 0 Then
                ' sort out rec
                rec.Top = .Layer(i).y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).x * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), DDS_Tileset(.Layer(i).tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
            Else
            If InMapEditor = False Then
            If .Layer(i).tileset > 0 Then
                ' sort out rec
                rec.Top = .Layer(i).y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).x * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), DDS_Tileset(.Layer(i).tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
            End If
            End If
            End If
            'MASK2-------------------------------------
             If i = MapLayer.Mask2 Then
            If InMapEditor = True And Mask2Unvisible = False Then
            If .Layer(i).tileset > 0 Then
                ' sort out rec
                rec.Top = .Layer(i).y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).x * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), DDS_Tileset(.Layer(i).tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
            Else
            If InMapEditor = False Then
            If .Layer(i).tileset > 0 Then
                ' sort out rec
                rec.Top = .Layer(i).y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).x * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), DDS_Tileset(.Layer(i).tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
            End If
            End If
            End If
        Next
    End With
End Sub


Public Sub DrawMenu()

Dim rec As RECT, x As Long, y As Long

If DDS_Menu Is Nothing Then

     Call InitDDSurf("backfinal", DDSD_Menu, DDS_Menu)

End If

With rec

.Top = 0

.Bottom = DDSD_Menu.lHeight

.Left = 0

.Right = DDSD_Menu.lWidth

End With
x = ConvertMapX(Player(MyIndex).x * PIC_X - 148 + Player(MyIndex).XOffset)
y = ConvertMapY(Player(MyIndex).y * PIC_Y - 148 + Player(MyIndex).YOffset)
'X = Camera.Left
'Y = Camera.Bottom - DDSD_Menu.lHeight
Call Engine_BltFast(x, y, DDS_Menu, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

End Sub


Public Sub DrawGDI()

'night
'If GDILoaded(GDI_IMAGE_NIGHT) = False Then
'Set GDIImage(GDI_IMAGE_NIGHT) = LoadPictureGDIplus(App.Path & "\Data Files\graphics\stage1.png")
'GDILoaded(GDI_IMAGE_NIGHT) = True
'End If

'DrawGDIImage GDIImage(GDI_IMAGE_NIGHT), 0, 0, GDIImage(GDI_IMAGE_NIGHT).width, GDIImage(GDI_IMAGE_NIGHT).height

If InIntro Then
If GDILoaded(GDI_IMAGE_OAK) = False Then
Set GDIImage(GDI_IMAGE_OAK) = LoadPictureGDIplus(App.Path & "\Data Files\pictures\1.png")
GDILoaded(GDI_IMAGE_OAK) = True
End If
If GDILoaded(GDI_IMAGE_EEVEE) = False Then
Set GDIImage(GDI_IMAGE_EEVEE) = LoadPictureGDIplus(App.Path & "\Data Files\pictures\2.gif")
GDILoaded(GDI_IMAGE_EEVEE) = True
End If
If DrawOak Then
DrawGDIImage GDIImage(GDI_IMAGE_OAK), Camera.Left + (frmMainGame.picScreen.Width - GDIImage(GDI_IMAGE_OAK).Width), Camera.Top + (frmMainGame.picScreen.Height - GDIImage(GDI_IMAGE_OAK).Height), GDIImage(GDI_IMAGE_OAK).Width, GDIImage(GDI_IMAGE_OAK).Height
End If
If DrawEevee Then
DrawGDIImage GDIImage(GDI_IMAGE_EEVEE), Camera.Left + (frmMainGame.picScreen.Width - GDIImage(GDI_IMAGE_OAK).Width), Camera.Top + (frmMainGame.picScreen.Height - GDIImage(GDI_IMAGE_EEVEE).Height * 2), GDIImage(GDI_IMAGE_EEVEE).Width, GDIImage(GDI_IMAGE_EEVEE).Height
End If
End If


If Dialogs > 0 And CurrentDialog > 0 Then
If DialogImage(CurrentDialog) > 0 Then

If GDILoaded(DialogImage(CurrentDialog)) = False Then
Set GDIImage(DialogImage(CurrentDialog)) = LoadPictureGDIplus(App.Path & "\Data Files\pictures\" & DialogImage(CurrentDialog) & ".png")
GDILoaded(DialogImage(CurrentDialog)) = True
End If

DrawGDIImage GDIImage(DialogImage(CurrentDialog)), Camera.Left + (frmMainGame.picScreen.Width - GDIImage(DialogImage(CurrentDialog)).Width), Camera.Top + (frmMainGame.picScreen.Height - GDIImage(DialogImage(CurrentDialog)).Height), GDIImage(DialogImage(CurrentDialog)).Width, GDIImage(DialogImage(CurrentDialog)).Height

End If
End If


'BATTLE
If inBattle = True Then
If PokeImgLoaded = True And EnemyPokeImgLoaded = True Then

If BattleImgLoaded = False Then
Set BattleImg = LoadPictureGDIplus(App.Path & "\Data Files\graphics\battleBase.png")
BattleImgLoaded = True
End If
If battleBaseImgloaded = False Then
Set battleBaseImg = LoadPictureGDIplus(App.Path & "\Data Files\graphics\base.png")
battleBaseImgloaded = True
End If

DrawGDIImage BattleImg, Camera.Left + 18, Camera.Top + 91, 497, 285
If BattlePokemon > 0 Then
Call DDS_BackBuffer.SetFillColor(RGB(0, 0, 0))
        Call DDS_BackBuffer.DrawBox(Camera.Left + 18 + 45, Camera.Top + 91 + 140, Camera.Left + 18 + 45 + 64, Camera.Top + 91 + 4 + 140)
        Call DDS_BackBuffer.SetFillColor(QBColor(BrightRed))
        Call DDS_BackBuffer.DrawBox(Camera.Left + 18 + 45, Camera.Top + 91 + 140, Camera.Left + 18 + 45 + ((PokemonInstance(BattlePokemon).HP / PokemonInstance(BattlePokemon).MaxHp) * 64), Camera.Top + 91 + 4 + 140)
'
Call DDS_BackBuffer.SetFillColor(RGB(0, 0, 0))
        Call DDS_BackBuffer.DrawBox(Camera.Left + 18 + 300, Camera.Top + 91 + 45, Camera.Left + 18 + 300 + 64, Camera.Top + 91 + 4 + 45)
        Call DDS_BackBuffer.SetFillColor(QBColor(BrightRed))
        If enemyPokemon.MaxHp > 0 Then
        Call DDS_BackBuffer.DrawBox(Camera.Left + 18 + 300, Camera.Top + 91 + 45, Camera.Left + 18 + 300 + ((enemyPokemon.HP / enemyPokemon.MaxHp) * 64), Camera.Top + 91 + 4 + 45)
        End If
        End If
DrawGDIImage battleBaseImg, Camera.Left + 18 + 350 - (battleBaseImg.Width / 3) + 5, Camera.Top + 91 + 60 + EnemyPokeImg.Height - (battleBaseImg.Height / 2), battleBaseImg.Width, battleBaseImg.Height
DrawGDIImage EnemyPokeImg, Camera.Left + 18 + 350, Camera.Top + 91 + 60, EnemyPokeImg.Width, EnemyPokeImg.Height
DrawGDIImage battleBaseImg, Camera.Left + 18 + 45 - (battleBaseImg.Width / 3) + 5 + (PokeImg.Width / 4), Camera.Top + 91 + 150 + PokeImg.Height - (battleBaseImg.Height / 2), battleBaseImg.Width, battleBaseImg.Height
DrawGDIImage PokeImg, Camera.Left + 18 + 45, Camera.Top + 91 + 150, PokeImg.Width, PokeImg.Height
End If
End If

End Sub





Public Sub DrawChaty()

Dim rec As RECT, x As Long, y As Long, i As Long

If DDS_Chat Is Nothing Then

     Call InitDDSurf("Chat", DDSD_Chat, DDS_Chat)

End If

With rec

.Top = 0

.Bottom = DDSD_Chat.lHeight

.Left = 0

.Right = DDSD_Chat.lWidth

End With



x = Camera.Left

y = Camera.Bottom - DDSD_Chat.lHeight
'X = 75
'Y = Camera.Bottom - 15 - DDSD_Chat.lHeight

Engine_BltFast x, y, DDS_Chat, rec, DDBLTFAST_SRCCOLORKEY

End Sub
''''''''''''''''''''''
Public Sub DrawFrames()
Dim rec As RECT, x As Long, y As Long, i As Long

If DDS_Frame Is Nothing Then

     Call InitDDSurf("Frame", DDSD_Frame, DDS_Frame)

End If

With rec

.Top = 0

.Bottom = DDSD_Frame.lHeight

.Left = 0

.Right = DDSD_Frame.lWidth

End With



For i = 0 To 5
x = Camera.Left + DDSD_Chat.lWidth + 210 + i * 8 + i * DDSD_Frame.lHeight
y = Camera.Bottom - 35
Engine_BltFast x, y, DDS_Frame, rec, DDBLTFAST_NOCOLORKEY
Next



End Sub




Public Sub BltMouseTile(ByVal Index As Long, ByVal l As Long, Optional ByVal xx As Long, Optional ByVal yy As Long)
 Dim rec As DxVBLib.RECT
    Dim i As Long
    Dim a As Long
    Dim x As Long
    Dim y As Long
    x = CurX
    y = CurY
If l = 1 Then
    rec.Top = 3 * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = 2 * PIC_X
    rec.Right = rec.Left + PIC_X
   ' render
    Call Engine_BltFast(ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), DDS_Tileset(8), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
Else
'Battle
If l = 2 Then
rec.Top = 3 * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = 3 * PIC_X
    rec.Right = rec.Left + PIC_X
   ' render
    Call Engine_BltFast(ConvertMapX(GetPlayerX(Index) * PIC_X), ConvertMapY((GetPlayerY(Index) - 2) * PIC_Y), DDS_Tileset(8), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End If
'GM
If l = 3 Then
rec.Top = 18 * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = 7 * PIC_X
    rec.Right = rec.Left + PIC_X
   ' render
    Call Engine_BltFast(ConvertMapX(GetPlayerX(Index) * PIC_X), ConvertMapY((GetPlayerY(Index) - 2) * PIC_Y), DDS_Tileset(8), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End If
'Mood
If l = 4 Then
Select Case xx
Case 0
rec.Top = 221 * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = 7 * PIC_X
    rec.Right = rec.Left + PIC_X
Case 1
    rec.Top = 219 * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = 7 * PIC_X
    rec.Right = rec.Left + PIC_X
End Select
   ' render
    Call Engine_BltFast(ConvertMapX(GetPlayerX(Index) * PIC_X), ConvertMapY((GetPlayerY(Index) - 2) * PIC_Y), DDS_Tileset(8), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End If
'END
End If
End Sub



Public Sub BltWeatherTile(ByVal Index As Long)
 Dim rec As DxVBLib.RECT
    Dim i As Long
    Dim a As Long
    Dim x As Long
    Dim y As Long
    x = CurX
    y = CurY
    
    rec.Top = 2 * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = 3 * PIC_X
    rec.Right = rec.Left + PIC_X
   ' render
    Call Engine_BltFast(ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), DDS_Tileset(8), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    



End Sub

Public Sub BltGM(ByVal Index As Long)
 Dim rec As DxVBLib.RECT
    Dim i As Long
    Dim x As Long
    Dim y As Long
    x = CurX
    y = CurY
    rec.Top = 3 * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = 2 * PIC_X
    rec.Right = rec.Left + PIC_X
   ' render
    Call Engine_BltFast(ConvertMapX(GetPlayerX(Index)), ConvertMapY(GetPlayerY(Index)), DDS_Tileset(9), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

End Sub

Public Sub BltWeather(ByVal weathernum As Long, ByVal ImageName As String)
Dim rec As RECT, x As Long, y As Long, i As Long

x = Weather(weathernum).pics_x(1)

y = Weather(weathernum).pics_Y(1)
BltCustomMapTile 5, 5, 9, 1, 1
 

End Sub


Public Sub BltCustomMapTile(ByVal x As Long, ByVal y As Long, ByVal tileset As Long, ByVal tilex As Long, ByVal tiley As Long)
    Dim rec As DxVBLib.RECT
    Dim i As Long
    
    
            ' skip tile if tileset isn't set
            If tileset > 0 Then
                ' sort out rec
                rec.Top = tiley * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = tilex * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), DDS_Tileset(tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
       
End Sub

Public Sub BltMapFringeTile(ByVal x As Long, ByVal y As Long)
    Dim rec As DxVBLib.RECT
    Dim i As Long
    If FlashLight = True Then
If x > Player(MyIndex).x + 3 Or x < Player(MyIndex).x - 3 Or y > Player(MyIndex).y + 3 Or y < Player(MyIndex).y - 3 Then Exit Sub
End If
    With map.Tile(x, y)
        For i = MapLayer.Fringe To MapLayer.Fringe2
            ' skip tile if tileset isn't set
            'FRINGE
            If i = MapLayer.Fringe Then
             If InMapEditor = True And FringeUnvisible = False Then
             If .Layer(i).tileset > 0 Then
                ' sort out rec
                rec.Top = .Layer(i).y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).x * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), DDS_Tileset(.Layer(i).tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
             End If
             Else
             If InMapEditor = False Then
             If .Layer(i).tileset > 0 Then
                ' sort out rec
                rec.Top = .Layer(i).y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).x * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), DDS_Tileset(.Layer(i).tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
            End If
            End If
            'FRINGE 2
            If i = MapLayer.Fringe2 Then
             If InMapEditor = True And Fringe2Unvisible = False Then
             If .Layer(i).tileset > 0 Then
                ' sort out rec
                rec.Top = .Layer(i).y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).x * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), DDS_Tileset(.Layer(i).tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
             End If
             Else
             If InMapEditor = False Then
             If .Layer(i).tileset > 0 Then
                ' sort out rec
                rec.Top = .Layer(i).y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).x * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), DDS_Tileset(.Layer(i).tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
            End If
            End If
        Next
    End With
End Sub

Public Sub BltDoor(ByVal x As Long, ByVal y As Long)
    Dim rec As DxVBLib.RECT
    Dim x2 As Long, y2 As Long
    
    ' sort out animation
    With TempTile(x, y)
        If .DoorAnimate = 1 Then ' opening
            If .DoorTimer + 100 < GetTickCount Then
                If .DoorFrame < 4 Then
                    .DoorFrame = .DoorFrame + 1
                Else
                    .DoorAnimate = 2 ' set to closing
                End If
                .DoorTimer = GetTickCount
            End If
        ElseIf .DoorAnimate = 2 Then ' closing
            If .DoorTimer + 100 < GetTickCount Then
                If .DoorFrame > 1 Then
                    .DoorFrame = .DoorFrame - 1
                Else
                    .DoorAnimate = 0 ' end animation
                End If
                .DoorTimer = GetTickCount
            End If
        End If
        
        If .DoorFrame = 0 Then .DoorFrame = 1
    End With


    If DDS_Door Is Nothing Then
        Call InitDDSurf("door", DDSD_Door, DDS_Door)
    End If

    With rec
        .Top = 0
        .Bottom = DDSD_Door.lHeight
        .Left = ((TempTile(x, y).DoorFrame - 1) * (DDSD_Door.lWidth / 4))
        .Right = .Left + (DDSD_Door.lWidth / 4)
    End With

    x2 = (x * PIC_X)
    y2 = (y * PIC_Y) - (DDSD_Door.lHeight / 2) + 4
    'Call Engine_BltFast(ConvertMapX(x2), ConvertMapY(y2), DDS_Door, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Call DDS_BackBuffer.BltFast(ConvertMapX(x2), ConvertMapY(y2), DDS_Door, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub



'/////////////// AROUND MAPS //////////////////////////////////
Public Sub BltMapUpTile(ByVal x As Long, ByVal y As Long, ByVal drawX As Long, ByVal drawY As Long)
On Error Resume Next
    Dim rec As DxVBLib.RECT
    Dim i As Long
    
    With UpMap.Tile(x, y)
        For i = MapLayer.Ground To MapLayer.Mask2
            If .Layer(i).tileset > 0 Then
                ' sort out rec
                rec.Top = .Layer(i).y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).x * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(drawX * PIC_X), ConvertMapY(drawY * PIC_Y), DDS_Tileset(.Layer(i).tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        Next
    End With
End Sub


Public Sub BltMapUpFringeTile(ByVal x As Long, ByVal y As Long, ByVal drawX As Long, ByVal drawY As Long)
On Error Resume Next
    Dim rec As DxVBLib.RECT
    Dim i As Long
    
    With UpMap.Tile(x, y)
        For i = MapLayer.Fringe To MapLayer.Fringe2
            If .Layer(i).tileset > 0 Then
                ' sort out rec
                rec.Top = .Layer(i).y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).x * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(drawX * PIC_X), ConvertMapY(drawY * PIC_Y), DDS_Tileset(.Layer(i).tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        Next
    End With
End Sub

'////// DOWN /////////
Public Sub BltMapDownTile(ByVal x As Long, ByVal y As Long, ByVal drawX As Long, ByVal drawY As Long)
    Dim rec As DxVBLib.RECT
    Dim i As Long
    
    With DownMap.Tile(x, y)
        For i = MapLayer.Ground To MapLayer.Mask2
            
            If .Layer(i).tileset > 0 Then
                ' sort out rec
                rec.Top = .Layer(i).y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).x * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(drawX * PIC_X), ConvertMapY(drawY * PIC_Y), DDS_Tileset(.Layer(i).tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
       
        Next
    End With
End Sub


Public Sub BltMapDownFringeTile(ByVal x As Long, ByVal y As Long, ByVal drawX As Long, ByVal drawY As Long)
    Dim rec As DxVBLib.RECT
    Dim i As Long
    
    With DownMap.Tile(x, y)
        For i = MapLayer.Fringe To MapLayer.Fringe2
         
            If .Layer(i).tileset > 0 Then
                ' sort out rec
                rec.Top = .Layer(i).y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).x * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(drawX * PIC_X), ConvertMapY(drawY * PIC_Y), DDS_Tileset(.Layer(i).tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
       
        Next
    End With
End Sub
'////////RIGHT//////////////
Public Sub BltMapRightTile(ByVal x As Long, ByVal y As Long, ByVal drawX As Long, ByVal drawY As Long)
On Error Resume Next
    Dim rec As DxVBLib.RECT
    Dim i As Long
    
    With RightMap.Tile(x, y)
        For i = MapLayer.Ground To MapLayer.Mask2
            
            If .Layer(i).tileset > 0 Then
                ' sort out rec
                rec.Top = .Layer(i).y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).x * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(drawX * PIC_X), ConvertMapY(drawY * PIC_Y), DDS_Tileset(.Layer(i).tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
       
        Next
    End With
End Sub


Public Sub BltMapRightFringeTile(ByVal x As Long, ByVal y As Long, ByVal drawX As Long, ByVal drawY As Long)
On Error Resume Next
    Dim rec As DxVBLib.RECT
    Dim i As Long
    
    With RightMap.Tile(x, y)
        For i = MapLayer.Fringe To MapLayer.Fringe2
         
            If .Layer(i).tileset > 0 Then
                ' sort out rec
                rec.Top = .Layer(i).y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).x * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(drawX * PIC_X), ConvertMapY(drawY * PIC_Y), DDS_Tileset(.Layer(i).tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
       
        Next
    End With
End Sub
'///////////LEFT////////

Public Sub BltMapLeftTile(ByVal x As Long, ByVal y As Long, ByVal drawX As Long, ByVal drawY As Long)
On Error Resume Next
    Dim rec As DxVBLib.RECT
    Dim i As Long
    
    With LeftMap.Tile(x, y)
        For i = MapLayer.Ground To MapLayer.Mask2
            
            If .Layer(i).tileset > 0 Then
                ' sort out rec
                rec.Top = .Layer(i).y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).x * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(drawX * PIC_X), ConvertMapY(drawY * PIC_Y), DDS_Tileset(.Layer(i).tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
       
        Next
    End With
End Sub


Public Sub BltMapLeftFringeTile(ByVal x As Long, ByVal y As Long, ByVal drawX As Long, ByVal drawY As Long)
On Error Resume Next
    Dim rec As DxVBLib.RECT
    Dim i As Long
    
    With LeftMap.Tile(x, y)
        For i = MapLayer.Fringe To MapLayer.Fringe2
         
            If .Layer(i).tileset > 0 Then
                ' sort out rec
                rec.Top = .Layer(i).y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).x * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(drawX * PIC_X), ConvertMapY(drawY * PIC_Y), DDS_Tileset(.Layer(i).tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
       
        Next
    End With
End Sub











Public Sub BltBlood(ByVal Index As Long)
    Dim rec As DxVBLib.RECT
    
    With Blood(Index)
        ' check if we should be seeing it
        If .Timer + 20000 < GetTickCount Then Exit Sub
        
        ' re-load graphics if need be
        If DDS_Blood Is Nothing Then
            Call InitDDSurf("Blood", DDSD_Blood, DDS_Blood)
        End If
        
        rec.Top = 0
        rec.Bottom = PIC_Y
        rec.Left = (.Sprite - 1) * PIC_X
        rec.Right = rec.Left + PIC_X
        
        Engine_BltFast ConvertMapX(.x * PIC_X), ConvertMapY(.y * PIC_Y), DDS_Blood, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    End With
End Sub
Sub drawpokemon(ByVal pokemonnum As Long)
Dim img As image
End Sub
Public Sub BltAnimation(ByVal Index As Long, ByVal Layer As Long)
    Dim Sprite As Integer
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    Dim i As Long
    Dim Width As Long, Height As Long
    Dim looptime As Long
    Dim FrameCount As Long
    Dim x As Long, y As Long
    Dim lockindex As Long
    
    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If
    
    Sprite = Animation(AnimInstance(Index).Animation).Sprite(Layer)
    
    If Sprite < 1 Or Sprite > NumAnimations Then Exit Sub
    
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
    
    AnimationTimer(Sprite) = GetTickCount + SurfaceTimerMax
    
    If DDS_Animation(Sprite) Is Nothing Then
        Call InitDDSurf("animations\" & Sprite, DDSD_Animation(Sprite), DDS_Animation(Sprite))
    End If
    
    ' total width divided by frame count
    Width = DDSD_Animation(Sprite).lWidth / FrameCount
    Height = DDSD_Animation(Sprite).lHeight
    
    sRECT.Top = 0
    sRECT.Bottom = Height
    sRECT.Left = (AnimInstance(Index).FrameIndex(Layer) - 1) * Width
    sRECT.Right = sRECT.Left + Width
    
    ' change x or y if locked
    If AnimInstance(Index).LockType > TARGET_TYPE_NONE Then ' if <> none
        ' is a player
        If AnimInstance(Index).LockType = TARGET_TYPE_PLAYER Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if is ingame
            If IsPlaying(lockindex) Then
                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    x = (GetPlayerX(lockindex) * PIC_X) + 16 - (Width / 2) + Player(lockindex).XOffset
                    y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (Height / 2) + Player(lockindex).YOffset
                End If
            End If
        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_NPC Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if NPC exists
            If MapNpc(lockindex).num > 0 Then
                ' check if alive
                If MapNpc(lockindex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    x = (MapNpc(lockindex).x * PIC_X) + 16 - (Width / 2) + MapNpc(lockindex).XOffset
                    y = (MapNpc(lockindex).y * PIC_Y) + 16 - (Height / 2) + MapNpc(lockindex).YOffset
                Else
                    ' npc not alive anymore, kill the animation
                    ClearAnimInstance Index
                    Exit Sub
                End If
            Else
                ' npc not alive anymore, kill the animation
                ClearAnimInstance Index
                Exit Sub
            End If
        End If
    Else
        ' no lock, default x + y
            Select Case GetPlayerDir(Index)
            Case DIR_UP
                y = (AnimInstance(Index).y * 32) + 16 - (Height / 2)
            Case Else
                y = (AnimInstance(Index).y * 32) + 16 - (Height / 2)
             End Select
        x = (AnimInstance(Index).x * 32) + 16 - (Width / 2)
    End If
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)

    ' Clip to screen
    If y < 0 Then

        With sRECT
            .Top = .Top - y
        End With

        y = 0
    End If

    If x < 0 Then

        With sRECT
            .Left = .Left - x
        End With

        x = 0
    End If

    If y + Height > DDSD_BackBuffer.lHeight Then
        sRECT.Bottom = sRECT.Bottom - (y + Height - DDSD_BackBuffer.lHeight)
    End If

    If x + Width > DDSD_BackBuffer.lWidth Then
        sRECT.Right = sRECT.Right - (x + Width - DDSD_BackBuffer.lWidth)
    End If
    
    Call Engine_BltFast(x, y, DDS_Animation(Sprite), sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub BltItem(ByVal itemnum As Long)
    Dim PicNum As Integer
    Dim rec As DxVBLib.RECT
    Dim MaxFrames As Byte
     If FlashLight = True Then
If MapItem(itemnum).x > Player(MyIndex).x + 3 Or MapItem(itemnum).x < Player(MyIndex).x - 3 Or MapItem(itemnum).y > Player(MyIndex).y + 3 Or MapItem(itemnum).y < Player(MyIndex).y - 3 Then Exit Sub
End If
    PicNum = Item(MapItem(itemnum).num).pic
    

    If PicNum < 1 Or PicNum > NumItems Then Exit Sub
    ItemTimer(PicNum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(PicNum) Is Nothing Then
        Call InitDDSurf("items\" & PicNum, DDSD_Item(PicNum), DDS_Item(PicNum))
    End If

    If DDSD_Item(PicNum).lWidth > 64 Then ' has more than 1 frame

        With rec
            .Top = 0
            .Bottom = 32
            .Left = (MapItem(itemnum).frame * 32)
            .Right = .Left + 32
        End With

    Else

        With rec
            .Top = 0
            .Bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With

    End If

    Call Engine_BltFast(ConvertMapX(MapItem(itemnum).x * PIC_X), ConvertMapY(MapItem(itemnum).y * PIC_Y), DDS_Item(PicNum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub BltMapResource(ByVal Resource_num As Long)
    Dim Resource_master As Long
    Dim Resource_state As Long
    Dim Resource_sprite As Long
    Dim rec As DxVBLib.RECT
    Dim x As Long, y As Long
    ' Get the Resource type
    Resource_master = map.Tile(MapResource(Resource_num).x, MapResource(Resource_num).y).data1
    
    If Resource_master = 0 Then Exit Sub

    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    ' Get the Resource state
    Resource_state = MapResource(Resource_num).ResourceState

    If Resource_state = 0 Then ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If

    ' Load early
    If DDS_Resource(Resource_sprite) Is Nothing Then
        Call InitDDSurf("Resources\" & Resource_sprite, DDSD_Resource(Resource_sprite), DDS_Resource(Resource_sprite))
    End If

    ' src rect
    With rec
        .Top = 0
        .Bottom = DDSD_Resource(Resource_sprite).lHeight
        .Left = 0
        .Right = DDSD_Resource(Resource_sprite).lWidth
    End With

    ' Set base x + y, then the offset due to size
    x = (MapResource(Resource_num).x * PIC_X) - (DDSD_Resource(Resource_sprite).lWidth / 2) + 16
    y = (MapResource(Resource_num).y * PIC_Y) - DDSD_Resource(Resource_sprite).lHeight + 32
    Call BltResource(Resource_sprite, x, y, rec)
End Sub

Private Sub BltResource(ByVal Resource As Long, ByVal dx As Long, dy As Long, rec As DxVBLib.RECT)

    If Resource < 1 Or Resource > NumResources Then Exit Sub
    Dim x As Long
    Dim y As Long
    Dim Width As Long
    Dim Height As Long
    Dim destRECT As DxVBLib.RECT
    ResourceTimer(Resource) = GetTickCount + SurfaceTimerMax

    If DDS_Resource(Resource) Is Nothing Then
        Call InitDDSurf("Resources\" & Resource, DDSD_Resource(Resource), DDS_Resource(Resource))
    End If

    x = ConvertMapX(dx)
    y = ConvertMapY(dy)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    ' Clip to screen
    If y < 0 Then

        With rec
            .Top = .Top - y
        End With

        y = 0
    End If

    If x < 0 Then

        With rec
            .Left = .Left - x
        End With

        x = 0
    End If

    If y + Height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (y + Height - DDSD_BackBuffer.lHeight)
    End If

    If x + Width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (x + Width - DDSD_BackBuffer.lWidth)
    End If

    ' End clipping
    Call Engine_BltFast(x, y, DDS_Resource(Resource), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Private Sub BltBars()
Dim tmpY As Long
Dim tmpX As Long
Dim barWidth As Long

    ' check for casting time bar
    If SpellBuffer > 0 Then
        ' lock to player
        tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset
        tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).YOffset + 35
        ' calculate the width to fill
        barWidth = ((GetTickCount - SpellBufferTimer) / ((GetTickCount - SpellBufferTimer) + (Spell(PlayerSpells(SpellBuffer)).CastTime * 1000)) * 64)
        ' draw bars
        Call DDS_BackBuffer.SetFillColor(RGB(0, 0, 0))
        Call DDS_BackBuffer.DrawBox(ConvertMapX(tmpX), ConvertMapY(tmpY), ConvertMapX(tmpX + 32), ConvertMapY(tmpY + 4))
        Call DDS_BackBuffer.SetFillColor(QBColor(BrightCyan))
        Call DDS_BackBuffer.DrawBox(ConvertMapX(tmpX), ConvertMapY(tmpY), ConvertMapX(tmpX + barWidth), ConvertMapY(tmpY + 4))
    End If
End Sub

Public Sub BltPlayer(ByVal Index As Long, ByVal mainpoke As Long)
    Dim anim As Byte
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim pokex As Long
    Dim pokey As Long
    Dim Sprite As Long, spriteleft As Long, spritevert As Long, spritevertFollow As Long, spritehori As Long
    Dim rec As DxVBLib.RECT
    Dim recFollow As DxVBLib.RECT
    Dim pokerec As DxVBLib.RECT
    Dim attackspeed As Long
    Sprite = GetPlayerSprite(Index)
    If Player(Index).notVisible And Index <> MyIndex Then Exit Sub
     If FlashLight = True Then
If Player(Index).x > Player(MyIndex).x + 3 Or Player(Index).x < Player(MyIndex).x - 3 Or Player(Index).y > Player(MyIndex).y + 3 Or Player(Index).y < Player(MyIndex).y - 3 Then Exit Sub
End If

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    ' speed from weapon
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        attackspeed = Item(GetPlayerEquipment(Index, Weapon)).speed
    Else
        attackspeed = 1000
    End If

    ' Reset frame
    anim = 0

    ' Check for attacking animation
    If Player(Index).AttackTimer + (attackspeed / 2) > GetTickCount Then
        If Player(Index).Attacking = 1 Then
            anim = 2
        End If
    Else
        ' If not attacking, walk normally
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).YOffset > 8) Then
                    anim = Player(Index).Step + 1
                End If
            Case DIR_DOWN
                If (Player(Index).YOffset < -8) Then anim = Player(Index).Step + 1
            Case DIR_LEFT
                If (Player(Index).XOffset > 8) Then anim = Player(Index).Step + 1
            Case DIR_RIGHT
                If (Player(Index).XOffset < -8) Then anim = Player(Index).Step + 1
            End Select
    End If

    ' Check to see if we want to stop making him attack
    With Player(Index)

        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If

    End With

    ' Set the left
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            spriteleft = 66
            spritevert = (DDSD_Character(Sprite).lHeight / 4) * 3 '192
            spritevertFollow = (256 / 4) * 3
        Case DIR_RIGHT
            spriteleft = 77
            spritevert = (DDSD_Character(Sprite).lHeight / 4) * 2 '128
        Case DIR_DOWN
            spriteleft = 88
            spritevert = 0 '0
        Case DIR_LEFT
            spriteleft = 99
            spritevert = DDSD_Character(Sprite).lHeight / 4 '64
    End Select
    
     
    With rec
        .Top = spritevert
        .Bottom = spritevert + DDSD_Character(Sprite).lHeight / 4
        .Left = anim * (DDSD_Character(Sprite).lWidth / 4)
        .Right = .Left + (DDSD_Character(Sprite).lWidth / 4)
    End With
    
    With recFollow
        .Top = spritevert
        .Bottom = spritevert + 256 / 4
        .Left = anim * (256 / 4)
        .Right = .Left + (256 / 4)
    End With
    
    ' Calculate the X
    x = GetPlayerX(Index) * PIC_X + Player(Index).XOffset - ((DDSD_Character(Sprite).lWidth / 4 - (DDSD_Character(Sprite).lWidth / 4)) / 2) - 15
    
    
    ' Is the player's height more than 32..?
    If (DDSD_Character(Sprite).lHeight) > 256 Then
        ' Create a 32 pixel offset for larger sprites
        y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - ((DDSD_Character(Sprite).lHeight) - 32)
    Else
        ' Proceed as normal
        y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - 32
    End If
    
       
     Select Case Player(Index).dir
        Case DIR_UP
            Call BltSprite(Sprite, x, y, rec)
            Call Blt2Sprite(Player(Index).Pokes(1), x, y + 40, recFollow)
        Case DIR_DOWN
            Call Blt2Sprite(Player(Index).Pokes(1), x, y - 40, recFollow)
             'Call DrawOverworld(Player(Index).dir, Player(Index).Step, x + 40, y + 20, Player(Index).Pokes(1))
            Call BltSprite(Sprite, x, y, rec)
        Case DIR_LEFT
            Call Blt2Sprite(Player(Index).Pokes(1), x + 40, y, recFollow)
             'Call DrawOverworld(Player(Index).dir, Player(Index).Step, x + 70, y + 48, Player(Index).Pokes(1))
            Call BltSprite(Sprite, x, y, rec)
        Case DIR_RIGHT
            Call Blt2Sprite(Player(Index).Pokes(1), x - 40, y, recFollow)
             'Call DrawOverworld(Player(Index).dir, Player(Index).Step, x + 8, y + 48, Player(Index).Pokes(1))
            Call BltSprite(Sprite, x, y, rec)
     End Select
     
     
     
     
    ' render the actual sprite
    ' Call BltSprite(Sprite, x, y, rec)
    'Call DrawOverworld(Player(Index).dir, Player(Index).Step, x + 8, y + 48, Player(Index).Pokes(1))
    'Render the following pokemon
    
    
    
    If Player(Index).HasBike <> YES Then
    ' check for paperdolling
        For i = 1 To Equipment.Equipment_Count - 1
            If GetPlayerEquipment(Index, i) <> 0 Then
                 If Item(GetPlayerEquipment(Index, i)).Paperdoll > 0 Then
                    Call BltPaperdoll(x, y, Item(GetPlayerEquipment(Index, i)).Paperdoll, anim, spriteleft)
                 End If
            End If
        Next
    End If
End Sub

Public Sub BltNpc(ByVal MapNpcNum As Long)
    Dim anim As Byte
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim Sprite As Long, spriteleft As Long, spritevert As Long
    Dim rec As DxVBLib.RECT
    Dim attackspeed As Long
    If FlashLight = True Then
If MapNpc(MapNpcNum).x > Player(MyIndex).x + 3 Or MapNpc(MapNpcNum).x < Player(MyIndex).x - 3 Or MapNpc(MapNpcNum).y > Player(MyIndex).y + 3 Or MapNpc(MapNpcNum).y < Player(MyIndex).y - 3 Then Exit Sub
End If
    
    If MapNpc(MapNpcNum).num = 0 Then Exit Sub ' no npc set
    
    Sprite = NPC(MapNpc(MapNpcNum).num).Sprite

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    attackspeed = 1000

    ' Reset frame
    anim = 0

    ' Check for attacking animation
    If MapNpc(MapNpcNum).AttackTimer + (attackspeed / 2) > GetTickCount Then
        If MapNpc(MapNpcNum).Attacking = 1 Then
            anim = 2
        End If
    Else
        ' If not attacking, walk normally
        Select Case MapNpc(MapNpcNum).dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).YOffset > 8) Then anim = MapNpc(MapNpcNum).Step + 1
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).YOffset < -8) Then anim = MapNpc(MapNpcNum).Step + 1
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).XOffset > 8) Then anim = MapNpc(MapNpcNum).Step + 1
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).XOffset < -8) Then anim = MapNpc(MapNpcNum).Step + 1
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    ' Set the left
    Select Case MapNpc(MapNpcNum).dir
        Case DIR_UP
            spriteleft = (DDSD_Character(Sprite).lHeight / 4) * 3
            spritevert = (DDSD_Character(Sprite).lHeight / 4) * 3
        Case DIR_RIGHT
            spriteleft = (DDSD_Character(Sprite).lHeight / 4) * 2
            spritevert = (DDSD_Character(Sprite).lHeight / 4) * 2
        Case DIR_DOWN
            spriteleft = 0
            spritevert = 0
        Case DIR_LEFT
            spriteleft = (DDSD_Character(Sprite).lHeight / 4)
            spritevert = DDSD_Character(Sprite).lHeight / 4
    End Select

    With rec
        .Top = spritevert
        .Bottom = spritevert + DDSD_Character(Sprite).lHeight / 4
        .Left = anim * (DDSD_Character(Sprite).lWidth / 4)
        .Right = .Left + (DDSD_Character(Sprite).lWidth / 4)
    End With

    ' Calculate the X
      x = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).XOffset - ((DDSD_Character(Sprite).lWidth / 4 - (DDSD_Character(Sprite).lWidth / 4)) / 2) - 15

    ' Is the player's height more than 32..?
    If (DDSD_Character(Sprite).lHeight) > 256 Then
        ' Create a 32 pixel offset for larger sprites
        y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - ((DDSD_Character(Sprite).lHeight) - 32)
    Else
        ' Proceed as normal
        y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset
    End If

    Call BltSprite(Sprite, x, y, rec)
    
    'Check for Paperdolls
    
    
    If NPC(MapNpc(MapNpcNum).num).Paperdoll1 > 0 Then
      Call BltPaperdoll(x, y, NPC(MapNpc(MapNpcNum).num).Paperdoll1, anim, spriteleft)
    End If
    If NPC(MapNpc(MapNpcNum).num).Paperdoll2 > 0 Then
      Call BltPaperdoll(x, y, NPC(MapNpc(MapNpcNum).num).Paperdoll2, anim, spriteleft)
    End If
    If NPC(MapNpc(MapNpcNum).num).Paperdoll3 > 0 Then
      Call BltPaperdoll(x, y, NPC(MapNpc(MapNpcNum).num).Paperdoll3, anim, spriteleft)
    End If
End Sub


Sub BltNpcScripts()

    
End Sub


Public Sub BltPaperdoll(ByVal x2 As Long, ByVal y2 As Long, ByVal Sprite As Long, ByVal anim As Long, ByVal spriteleft As Long)
    Dim rec As DxVBLib.RECT
    Dim x As Long, y As Long
    Dim Width As Long, Height As Long

    If Sprite < 1 Or Sprite > NumPaperdolls Then Exit Sub
    
    If DDS_Paperdoll(Sprite) Is Nothing Then
        Call InitDDSurf("Paperdolls\" & Sprite, DDSD_Paperdoll(Sprite), DDS_Paperdoll(Sprite))
    End If
    
    Select Case spriteleft
        Case 66
             With rec
                .Top = 192
                .Bottom = 192 + DDSD_Paperdoll(Sprite).lHeight / 4
                .Left = anim * (DDSD_Paperdoll(Sprite).lWidth / 4)
                .Right = .Left + (DDSD_Paperdoll(Sprite).lWidth / 4)
            End With
        Case 77
             With rec
                .Top = 128
                .Bottom = 128 + DDSD_Paperdoll(Sprite).lHeight / 4
                .Left = anim * (DDSD_Paperdoll(Sprite).lWidth / 4)
                .Right = .Left + (DDSD_Paperdoll(Sprite).lWidth / 4)
            End With
        Case 88
             With rec
                .Top = 0
                .Bottom = 0 + DDSD_Paperdoll(Sprite).lHeight / 4
                .Left = anim * (DDSD_Paperdoll(Sprite).lWidth / 4)
                .Right = .Left + (DDSD_Paperdoll(Sprite).lWidth / 4)
            End With
        Case 99
             With rec
                .Top = 64
                .Bottom = 64 + DDSD_Paperdoll(Sprite).lHeight / 4
                .Left = anim * (DDSD_Paperdoll(Sprite).lWidth / 4)
                .Right = .Left + (DDSD_Paperdoll(Sprite).lWidth / 4)
            End With
        Case Else
             With rec
                .Top = 0
                .Bottom = 0 + DDSD_Paperdoll(Sprite).lHeight / 4
                .Left = anim * (DDSD_Paperdoll(Sprite).lWidth / 4)
                .Right = .Left + (DDSD_Paperdoll(Sprite).lWidth / 4)
            End With
    End Select
    
   
    
    ' clipping
    x = ConvertMapX(x2)
    y = ConvertMapY(y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    ' Clip to screen
    If y < 0 Then
        With rec
            .Top = .Top - y
        End With
        y = 0
    End If

    If x < 0 Then
        With rec
            .Left = .Left - x
        End With
        x = 0
    End If

    If y + Height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (y + Height - DDSD_BackBuffer.lHeight)
    End If

    If x + Width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (x + Width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping
    
    Call Engine_BltFast(x, y, DDS_Paperdoll(Sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Private Sub BltSprite(ByVal Sprite As Long, ByVal x2 As Long, y2 As Long, rec As DxVBLib.RECT)
    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    Dim x As Long
    Dim y As Long
    Dim Width As Long
    Dim Height As Long
    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    x = ConvertMapX(x2)
    y = ConvertMapY(y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    ' clipping
    If y < 0 Then
        With rec
            .Top = .Top - y
        End With
        y = 0
    End If

    If x < 0 Then
        With rec
            .Left = .Left - x
        End With
        x = 0
    End If

    If y + Height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (y + Height - DDSD_BackBuffer.lHeight)
    End If

    If x + Width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (x + Width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping
    
    Call Engine_BltFast(x, y, DDS_Character(Sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Private Sub Blt2Sprite(ByVal Sprite As Long, ByVal x2 As Long, y2 As Long, rec As DxVBLib.RECT)
    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    Dim x As Long
    Dim y As Long
    Dim Width As Long
    Dim Height As Long
    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        If PokemonInstance(1).isShiny = True Then
            Call InitDDSurf("characters\" & Sprite & "-1", DDSD_Character(Sprite), DDS_Character(Sprite))
        Else
            Call InitDDSurf("characters\" & Sprite & "", DDSD_Character(Sprite), DDS_Character(Sprite))
        End If
        
    End If

    'x = ConvertMapX(x2)
    'y = ConvertMapY(y2)
    x = x2 - (TileView.Left * PIC_X)
    y = y2 - (TileView.Top * PIC_Y)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    ' clipping
    If y < 0 Then
        With rec
            .Top = .Top - y
        End With
        y = 0
    End If

    If x < 0 Then
        With rec
            .Left = .Left - x
        End With
        x = 0
    End If

    If y + Height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (y + Height - DDSD_BackBuffer.lHeight)
    End If

    If x + Width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (x + Width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping
    
    Call Engine_BltFast(x, y, DDS_Character(Sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltAnimatedInvItems()

End Sub

Sub BltEquipment()
   
End Sub

Sub BltInventory()
   
End Sub

Sub BltPlayerSpells()
  
End Sub

Sub BltShop()
   
End Sub

Public Sub BltInventoryItem(ByVal x As Long, ByVal y As Long)


End Sub

' ******************
' ** Game Editors **
' ******************
Public Sub EditorMap_BltTileset()
    Dim Height As Long
    Dim Width As Long
    Dim tileset As Byte
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    
    ' find tileset number
    tileset = frmEditor_Map.scrlTileSet.Value
    
    ' exit out if doesn't exist
    If tileset < 0 Or tileset > NumTileSets Then Exit Sub
    
    ' make sure it's loaded
    If DDS_Tileset(tileset) Is Nothing Then
        Call InitDDSurf("tilesets\" & tileset, DDSD_Tileset(tileset), DDS_Tileset(tileset))
    End If
    
    Height = DDSD_Tileset(tileset).lHeight
    Width = DDSD_Tileset(tileset).lWidth
    
    dRECT.Top = 0
    dRECT.Bottom = Height
    dRECT.Left = 0
    dRECT.Right = Width
    
    frmEditor_Map.picBackSelect.Height = Height
    frmEditor_Map.picBackSelect.Width = Width
    
    Call Engine_BltToDC(DDS_Tileset(tileset), sRECT, dRECT, frmEditor_Map.picBackSelect)
End Sub

Public Sub BltTileOutline()
    Dim rec As DxVBLib.RECT

    With rec
        .Top = 0
        .Bottom = .Top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With

    Call Engine_BltFast(ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), DDS_Misc, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub NewCharacterBltSprite(ByVal Sprite As Long)
     Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    Dim Width As Long, Height As Long
    
    If frmMenu.cmbClass.ListIndex = -1 Then Exit Sub
    
    If newCharClass = 0 Then
        Sprite = 657
    Else
        Sprite = 658
    End If
    
    If Sprite < 1 Or Sprite > NumCharacters Then
        frmMenu.picSprite.Cls
        Exit Sub
    End If
    
    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("Characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If
    
    Width = DDSD_Character(Sprite).lWidth / 4
    Height = DDSD_Character(Sprite).lHeight / 4
    
    frmMenu.picSprite.Width = Width
    frmMenu.picSprite.Height = Height
    
    sRECT.Top = 0
    sRECT.Bottom = Height
    sRECT.Left = Width * 7 'looking down
    sRECT.Right = sRECT.Left + Width
    
    dRECT.Top = 0
    dRECT.Bottom = Height
    dRECT.Left = 0
    dRECT.Right = Width
    
    Call Engine_BltToDC(DDS_Character(Sprite), sRECT, dRECT, frmMenu.picSprite)
End Sub

Public Sub EditorMap_BltMapItem()
    Dim itemnum As Integer
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    itemnum = Item(frmEditor_Map.scrlMapItem.Value).pic

    If itemnum < 1 Or itemnum > NumItems Then
        frmEditor_Map.picMapItem.Cls
        Exit Sub
    End If

    ItemTimer(itemnum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(itemnum) Is Nothing Then
        Call InitDDSurf("Items\" & itemnum, DDSD_Item(itemnum), DDS_Item(itemnum))
    End If

    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    Call Engine_BltToDC(DDS_Item(itemnum), sRECT, dRECT, frmEditor_Map.picMapItem)
End Sub

Public Sub EditorMap_BltKey()
    Dim itemnum As Integer
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    itemnum = Item(frmEditor_Map.scrlMapKey.Value).pic

    If itemnum < 1 Or itemnum > NumItems Then
        frmEditor_Map.picMapKey.Cls
        Exit Sub
    End If

    ItemTimer(itemnum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(itemnum) Is Nothing Then
        Call InitDDSurf("Items\" & itemnum, DDSD_Item(itemnum), DDS_Item(itemnum))
    End If

    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    Call Engine_BltToDC(DDS_Item(itemnum), sRECT, dRECT, frmEditor_Map.picMapKey)
End Sub

Public Sub EditorItem_BltItem()
    Dim itemnum As Integer
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    itemnum = frmEditor_Item.scrlPic.Value

    If itemnum < 1 Or itemnum > NumItems Then
        frmEditor_Item.picItem.Cls
        Exit Sub
    End If

    ItemTimer(itemnum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(itemnum) Is Nothing Then
        Call InitDDSurf("Items\" & itemnum, DDSD_Item(itemnum), DDS_Item(itemnum))
    End If

    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    ' same for destination as source
    dRECT = sRECT
    Call Engine_BltToDC(DDS_Item(itemnum), sRECT, dRECT, frmEditor_Item.picItem)
End Sub


Public Sub EditorItem_BltPaperdoll()
    Dim Sprite As Long
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    
    frmEditor_Item.picPaperdoll.Cls
    
    Sprite = frmEditor_Item.scrlPaperdoll.Value

    If Sprite < 1 Or Sprite > NumPaperdolls Then
        frmEditor_Item.picItem.Cls
        Exit Sub
    End If

    PaperdollTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Paperdoll(Sprite) Is Nothing Then
        Call InitDDSurf("paperdolls\" & Sprite, DDSD_Paperdoll(Sprite), DDS_Paperdoll(Sprite))
    End If

    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = DDSD_Paperdoll(Sprite).lHeight
    sRECT.Left = 0
    sRECT.Right = DDSD_Paperdoll(Sprite).lWidth
    ' same for destination as source
    dRECT = sRECT
    
    Call Engine_BltToDC(DDS_Paperdoll(Sprite), sRECT, dRECT, frmEditor_Item.picPaperdoll)
End Sub




Public Sub EditorSpell_BltIcon()
    Dim iconnum As Integer
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    iconnum = frmEditor_Spell.scrlIcon.Value
    
    If iconnum < 1 Or iconnum > NumSpellIcons Then
        frmEditor_Spell.picSprite.Cls
        Exit Sub
    End If
    
    SpellIconTimer(iconnum) = GetTickCount + SurfaceTimerMax
    
    If DDS_SpellIcon(iconnum) Is Nothing Then
        Call InitDDSurf("SpellIcons\" & iconnum, DDSD_SpellIcon(iconnum), DDS_SpellIcon(iconnum))
    End If
    
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    
    Call Engine_BltToDC(DDS_SpellIcon(iconnum), sRECT, dRECT, frmEditor_Spell.picSprite)
End Sub

Public Sub EditorAnim_BltAnim()
    Dim Animationnum As Integer
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    Dim i As Long
    Dim Width As Long, Height As Long
    Dim looptime As Long
    Dim FrameCount As Long
    Dim ShouldRender As Boolean
    
    For i = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(i).Value
        
        If Animationnum < 1 Or Animationnum > NumAnimations Then
            frmEditor_Animation.picSprite(i).Cls
        Else
            looptime = frmEditor_Animation.scrlLoopTime(i)
            FrameCount = frmEditor_Animation.scrlFrameCount(i)
            
            ShouldRender = False
            
            ' check if we need to render new frame
            If AnimEditorTimer(i) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimEditorFrame(i) >= FrameCount Then
                    AnimEditorFrame(i) = 1
                Else
                    AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                End If
                AnimEditorTimer(i) = GetTickCount
                ShouldRender = True
            End If
        
            If ShouldRender Then
                frmEditor_Animation.picSprite(i).Cls
            
                AnimationTimer(Animationnum) = GetTickCount + SurfaceTimerMax
                
                If DDS_Animation(Animationnum) Is Nothing Then
                    Call InitDDSurf("animations\" & Animationnum, DDSD_Animation(Animationnum), DDS_Animation(Animationnum))
                End If
                
                If frmEditor_Animation.scrlFrameCount(i).Value > 0 Then
                    ' total width divided by frame count
                    Width = DDSD_Animation(Animationnum).lWidth / frmEditor_Animation.scrlFrameCount(i).Value
                    Height = DDSD_Animation(Animationnum).lHeight
                    
                    sRECT.Top = 0
                    sRECT.Bottom = Height
                    sRECT.Left = (AnimEditorFrame(i) - 1) * Width
                    sRECT.Right = sRECT.Left + Width
                    
                    dRECT.Top = 0
                    dRECT.Bottom = Height
                    dRECT.Left = 0
                    dRECT.Right = Width
                    
                    Call Engine_BltToDC(DDS_Animation(Animationnum), sRECT, dRECT, frmEditor_Animation.picSprite(i))
                End If
            End If
        End If
    Next
End Sub

Public Sub EditorNpc_BltSprite()
    Dim Sprite As Long
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    Sprite = frmEditor_NPC.scrlSprite.Value

    If Sprite < 1 Or Sprite > NumCharacters Then
        frmEditor_NPC.picSprite.Cls
        Exit Sub
    End If

    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    sRECT.Top = 0
    sRECT.Bottom = SIZE_Y
    sRECT.Left = PIC_X * 3 ' facing down
    sRECT.Right = sRECT.Left + SIZE_X
    dRECT.Top = 0
    dRECT.Bottom = SIZE_Y
    dRECT.Left = 0
    dRECT.Right = SIZE_X
    Call Engine_BltToDC(DDS_Character(Sprite), sRECT, dRECT, frmEditor_NPC.picSprite)
End Sub

Public Sub EditorResource_BltSprite()
    Dim Sprite As Long
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    
    ' normal sprite
    Sprite = frmEditor_Resource.scrlNormalPic.Value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picNormalPic.Cls
    Else
        ResourceTimer(Sprite) = GetTickCount + SurfaceTimerMax
        If DDS_Resource(Sprite) Is Nothing Then
            Call InitDDSurf("Resources\" & Sprite, DDSD_Resource(Sprite), DDS_Resource(Sprite))
        End If
        sRECT.Top = 0
        sRECT.Bottom = DDSD_Resource(Sprite).lHeight
        sRECT.Left = 0
        sRECT.Right = DDSD_Resource(Sprite).lWidth
        dRECT.Top = 0
        dRECT.Bottom = DDSD_Resource(Sprite).lHeight
        dRECT.Left = 0
        dRECT.Right = DDSD_Resource(Sprite).lWidth
        Call Engine_BltToDC(DDS_Resource(Sprite), sRECT, dRECT, frmEditor_Resource.picNormalPic)
    End If

    ' exhausted sprite
    Sprite = frmEditor_Resource.scrlExhaustedPic.Value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picExhaustedPic.Cls
    Else
        ResourceTimer(Sprite) = GetTickCount + SurfaceTimerMax
        If DDS_Resource(Sprite) Is Nothing Then
            Call InitDDSurf("Resources\" & Sprite, DDSD_Resource(Sprite), DDS_Resource(Sprite))
        End If
        sRECT.Top = 0
        sRECT.Bottom = DDSD_Resource(Sprite).lHeight
        sRECT.Left = 0
        sRECT.Right = DDSD_Resource(Sprite).lWidth
        dRECT.Top = 0
        dRECT.Bottom = DDSD_Resource(Sprite).lHeight
        dRECT.Left = 0
        dRECT.Right = DDSD_Resource(Sprite).lWidth
        Call Engine_BltToDC(DDS_Resource(Sprite), sRECT, dRECT, frmEditor_Resource.picExhaustedPic)
    End If
End Sub

Public Sub Render_Graphics()
    'On Error GoTo ErrorHandle
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim rec As DxVBLib.RECT
    Dim rec_pos As DxVBLib.RECT
    
    ' don't render
    If Not CheckSurfaces Then Exit Sub
    ' don't render
    If frmMainGame.WindowState = vbMinimized Then Exit Sub
    If GettingMap Then Exit Sub
    
    UpdateCamera
    
    ' update animation editor
    If Editor = EDITOR_ANIMATION Then
        EditorAnim_BltAnim
    End If
    
    DDS_BackBuffer.BltColorFill rec_pos, 0

    ' blit lower tiles
    If NumTileSets > 0 Then
        For x = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(x, y) Then
                    Call BltMapTile(x, y)
                    Else
                    If Options.NearbyMaps = YES Then
                    If map.Up > 0 Then 'If it has upper map
                    If IsUpMapPoint(x, y) Then 'If the point in view is in upper map blt it
                      Call BltMapUpTile(x, UpMap.MaxY + y + 1, x, y)
                    End If
                    End If
                    
                    If map.Down > 0 Then
                    If IsDownMapPoint(x, y) Then
                    Call BltMapDownTile(x, y - map.MaxY - 1, x, y)
                    End If
                    End If
                    
                    If map.Right > 0 Then
                    If IsRightMapPoint(x, y) Then
                    Call BltMapRightTile(x - map.MaxX - 1, y, x, y)
                    End If
                    End If
                    
                    If map.Left > 0 Then
                    If IsLeftMapPoint(x, y) Then
                    Call BltMapLeftTile(LeftMap.MaxX + x + 1, y, x, y)
                    End If
                    End If
                    End If
                   End If
            Next
        Next
    End If
    
    
    
    'Check to draw GymBlock
    If NumTileSets > 0 Then
    For x = TileView.Left To TileView.Right
    For y = TileView.Top To TileView.Bottom
    If IsValidMapPoint(x, y) Then
    
       If map.Tile(x, y).Type = TILE_TYPE_GYMBLOCK Then
       
       Dim dirr As Long
       dirr = map.Tile(x, y).data2
       Dim gymneeded As Long
       gymneeded = map.Tile(x, y).data1
       If Not Player(MyIndex).Badge(gymneeded) = GYM_DEFEATED Then

       End If
       'Call BltCustomMapTile(X, Y, 2, dirr, 138)
       
    
       
       End If
       End If
    Next
    Next
    End If
   
    
    For i = 1 To MAX_PLAYERS
    If IsPlaying(i) Then
    If Player(i).Access >= 1 Then
    If Player(i).Moving = 0 Then
    'Call BltMouseTile(i, 2)
    End If
    End If
    End If
    Next
    
    'Blt Mouse
    Call BltMouseTile(MyIndex, 1)

For i = 1 To MAX_PLAYERS
If IsPlaying(i) Then
If Player(i).inBattle = YES And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
 Call BltMouseTile(i, 2)
End If
End If
Next

    
    For i = 1 To MAX_BYTE
        Call BltBlood(i)
    Next

    ' Blit out the items
    If NumItems > 0 Then
        For i = 1 To MAX_MAP_ITEMS
    
            If MapItem(i).num > 0 Then
                Call BltItem(i)
            End If
    
        Next
    End If
    
    'Draw sum d00rs.
    For x = TileView.Left To TileView.Right
        For y = TileView.Top To TileView.Bottom

            If IsValidMapPoint(x, y) Then
                If map.Tile(x, y).Type = TILE_TYPE_DOOR Then
                    BltDoor x, y
                End If
            End If

        Next
    Next
    
    ' draw animations
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(0) Then
                BltAnimation i, 0
            End If
        Next
    End If

    ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
    For y = 0 To map.MaxY

        If NumCharacters > 0 Then
            ' Players
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If Player(i).y = y Then
                        If Player(i).Access > 0 And GetPlayerX(i) = CurX And GetPlayerY(i) = CurY Then
                        If Player(i).notVisible = True And i <> MyIndex Then
                        Else
                        Call BltMouseTile(i, 3)
                        End If
                        Else
                        
                        End If
                        Call BltPlayer(i, PokemonInstance(1).PokemonNumber)
                    End If
                End If
            
            Next
        
            ' Npcs
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).y = y Then
                    Call BltNpc(i)
                
                End If
            Next
        End If
        
        BltNpcScripts
        
        
        ' Resources
        If NumResources > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For i = 1 To Resource_Index
                        If MapResource(i).y = y Then
                            Call BltMapResource(i)
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    ' animations
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(1) Then
                BltAnimation i, 1
            End If
        Next
    End If

    ' blit out upper tiles
    ' blit out upper tiles
    If NumTileSets > 0 Then
        For x = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(x, y) Then
                    Call BltMapFringeTile(x, y)
                    Else
                    If Options.NearbyMaps = YES Then
                    If map.Up > 0 Then 'If it has upper map
                    If IsUpMapPoint(x, y) Then 'If the point in view is in upper map blt it
                      Call BltMapUpFringeTile(x, UpMap.MaxY + y + 1, x, y)
                    End If
                    End If
                    
                    If map.Down > 0 Then
                    If IsDownMapPoint(x, y) Then
                    Call BltMapDownFringeTile(x, y - map.MaxY - 1, x, y)
                    End If
                    End If
                    
                    If map.Right > 0 Then
                    If IsRightMapPoint(x, y) Then
                    Call BltMapRightFringeTile(x - map.MaxX - 1, y, x, y)
                    End If
                    End If
                    
                    If map.Left > 0 Then
                    If IsLeftMapPoint(x, y) Then
                    Call BltMapLeftFringeTile(LeftMap.MaxX + x + 1, y, x, y)
                    End If
                    End If
                    End If
                End If
            Next
        Next
        
    End If
    If FlashLight = True Then
    DrawMenu
    End If
    ' blit out a square at mouse cursor
    If InMapEditor Then
        Call BltTileOutline
    End If
    
   DrawGDI
    
   
    
    'Call DrawFrames
    BltBars
    


    ' Lock the backbuffer so we can draw text and names
    TexthDC = DDS_BackBuffer.GetDC

    ' draw FPS
    If BFPS Then
        Call DrawText(TexthDC, Camera.Right - (Len("FPS: " & GameFPS) * 8), Camera.Top + 1, Trim$("FPS: " & GameFPS), QBColor(Yellow))
    End If
    Call DrawBattleText
   

    ' draw cursor, player X and Y locations
    If BLoc Then
        Call DrawText(TexthDC, Camera.Left, Camera.Top + 1, Trim$("cur x: " & CurX & " y: " & CurY), QBColor(Yellow))
        Call DrawText(TexthDC, Camera.Left, Camera.Top + 15, Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), QBColor(Yellow))
        Call DrawText(TexthDC, Camera.Left, Camera.Top + 27, Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), QBColor(Yellow))
    End If

    
    
    'draw Weather
    If Weather(1).Pics > 0 Then
    BltMouseTile MyIndex, 2
    End If

    ' draw player names
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
        If i <> MyIndex Then
            Call DrawPlayerName(i)
            End If
        End If
    Next
    
    'Draw npc names
    
    For i = 1 To MAX_MAP_NPCS
     If MapNpc(i).num > 0 Then
     DrawNpcName (i)
     End If
    Next
    
    For i = 1 To MAX_BYTE
        Call BltActionMsg(i)
    Next i

    ' Blit out map attributes
    If InMapEditor Then
        Call BltMapAttributes
    End If

    ' Draw map name
    Call DrawText(TexthDC, DrawMapNameX, DrawMapNameY, map.Name, DrawMapNameColor)
    
   
Continue:
    ' Release DC
    Call DDS_BackBuffer.ReleaseDC(TexthDC)
    ' Get the rect to blit to
    Call DX7.GetWindowRect(frmMainGame.picScreen.hwnd, rec_pos)
    ' Blit the backbuffer
    Call DDS_Primary.Blt(rec_pos, DDS_BackBuffer, Camera, DDBLT_WAIT)
    Exit Sub
ErrorHandle:

    If Err.number = 91 Then
        Sleep 100
        Call ReInitDD
        Err.Clear
        Exit Sub
    End If

    On Error Resume Next

    If Not CheckSurfaces Then Exit Sub ' surfaces can get lost, check again
    TexthDC = DDS_BackBuffer.GetDC ' Lock the backbuffer so we can draw text and names
    Call DrawText(TexthDC, ConvertMapX(10), ConvertMapY(15), "Error Rendering Graphics - Unhandled Error", QBColor(BrightRed))
    Call DrawText(TexthDC, ConvertMapX(10), ConvertMapY(28), "Error Number : " & Err.number & " - " & Err.Description, QBColor(BrightCyan))
    GoTo Continue
End Sub

Public Sub UpdateCamera()
    Dim OffsetX As Long
    Dim OffsetY As Long
    Dim StartX As Long
    Dim StartY As Long
    Dim EndX As Long
    Dim EndY As Long
    OffsetX = Player(MyIndex).XOffset + PIC_X
    OffsetY = Player(MyIndex).YOffset + PIC_Y
    StartX = GetPlayerX(MyIndex) - ((MAX_MAPX + 1) \ 2) - 1
    StartY = GetPlayerY(MyIndex) - ((MAX_MAPY + 1) \ 2) - 1
    
    
    If Options.CameraFollowPlayer = NO Then
    If StartX < 0 Then
        OffsetX = 0

        If StartX = -1 Then
            If Player(MyIndex).XOffset > 0 Then
                OffsetX = Player(MyIndex).XOffset
            End If
        End If

        StartX = 0
    End If

    If StartY < 0 Then
        OffsetY = 0

        If StartY = -1 Then
            If Player(MyIndex).YOffset > 0 Then
                OffsetY = Player(MyIndex).YOffset
            End If
        End If

        StartY = 0
    End If
    End If
    
    
    EndX = StartX + (MAX_MAPX + 1) + 1
    EndY = StartY + (MAX_MAPY + 1) + 1


    If Options.CameraFollowPlayer = NO Then
    If EndX > map.MaxX Then
        OffsetX = 32

        If EndX = map.MaxX + 1 Then
            If Player(MyIndex).XOffset < 0 Then
                OffsetX = Player(MyIndex).XOffset + PIC_X
            End If
        End If

        EndX = map.MaxX
        StartX = EndX - MAX_MAPX - 1
    End If
'
    If EndY > map.MaxY Then
        OffsetY = 32

        If EndY = map.MaxY + 1 Then
            If Player(MyIndex).YOffset < 0 Then
                OffsetY = Player(MyIndex).YOffset + PIC_Y
            End If
        End If

        EndY = map.MaxY
        StartY = EndY - MAX_MAPY - 1
    End If
    End If

    With TileView
        .Top = StartY
        .Bottom = EndY
        .Left = StartX
        .Right = EndX
    End With
    
    With Camera
        .Top = OffsetY
        .Bottom = .Top + ScreenY
        .Left = OffsetX
        .Right = .Left + ScreenX
    End With
    
    UpdateDrawMapName

End Sub

Public Function ConvertMapX(ByVal x As Long) As Long
    ConvertMapX = x - (TileView.Left * PIC_X)
End Function

Public Function ConvertMapY(ByVal y As Long) As Long
    ConvertMapY = y - (TileView.Top * PIC_Y)
End Function

Public Function InViewPort(ByVal x As Long, ByVal y As Long) As Boolean
    InViewPort = False

    If x < TileView.Left Then Exit Function
    If y < TileView.Top Then Exit Function
    If x > TileView.Right Then Exit Function
    If y > TileView.Bottom Then Exit Function
    InViewPort = True
End Function

Public Function IsValidMapPoint(ByVal x As Long, ByVal y As Long) As Boolean
    IsValidMapPoint = False

    If x < 0 Then Exit Function
    If y < 0 Then Exit Function
    If x > map.MaxX Then Exit Function
    If y > map.MaxY Then Exit Function
    IsValidMapPoint = True
End Function


Public Function IsUpMapPoint(ByVal x As Long, ByVal y As Long)
IsUpMapPoint = False
If x < 0 Then Exit Function
If x > map.MaxX Then Exit Function
If y < 0 Then
IsUpMapPoint = True
End If
End Function

Public Function IsDownMapPoint(ByVal x As Long, ByVal y As Long)
IsDownMapPoint = False
If x < 0 Then Exit Function
If x > map.MaxX Then Exit Function
If y > map.MaxY Then
IsDownMapPoint = True
End If
End Function

Public Function IsRightMapPoint(ByVal x As Long, ByVal y As Long)
IsRightMapPoint = False
If y > map.MaxY Then Exit Function
If y < 0 Then Exit Function
If x > map.MaxX Then
IsRightMapPoint = True
End If
End Function

Public Function IsLeftMapPoint(ByVal x As Long, ByVal y As Long)
IsLeftMapPoint = False
If y > map.MaxY Then Exit Function
If y < 0 Then Exit Function
If x < map.MaxX Then
IsLeftMapPoint = True
End If
End Function








Public Sub LoadTilesets()
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim tilesetInUse() As Boolean
    
    ReDim tilesetInUse(0 To NumTileSets)
    
    For x = 0 To map.MaxX
        For y = 0 To map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                tilesetInUse(map.Tile(x, y).Layer(i).tileset) = True
            Next
        Next
    Next
    
    'neighbour maps
    If map.Up > 0 Then
    For x = 0 To UpMap.MaxX
        For y = 0 To UpMap.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                tilesetInUse(UpMap.Tile(x, y).Layer(i).tileset) = True
            Next
        Next
    Next
    End If
    
    If map.Down > 0 Then
    For x = 0 To DownMap.MaxX
        For y = 0 To DownMap.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                tilesetInUse(DownMap.Tile(x, y).Layer(i).tileset) = True
            Next
        Next
    Next
    End If
    
    If map.Left > 0 Then
    For x = 0 To LeftMap.MaxX
        For y = 0 To LeftMap.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                tilesetInUse(LeftMap.Tile(x, y).Layer(i).tileset) = True
            Next
        Next
    Next
    End If
    
    If map.Right > 0 Then
    For x = 0 To RightMap.MaxX
        For y = 0 To RightMap.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                tilesetInUse(RightMap.Tile(x, y).Layer(i).tileset) = True
            Next
        Next
    Next
    End If
    
    tilesetInUse(8) = True
    
    
    For i = 1 To NumTileSets
        If tilesetInUse(i) Then
            ' load tileset
            If DDS_Tileset(i) Is Nothing Then
                Call InitDDSurf("tilesets\" & i, DDSD_Tileset(i), DDS_Tileset(i))
            End If
        Else
            ' unload tileset
            Call ZeroMemory(ByVal VarPtr(DDSD_Tileset(i)), LenB(DDSD_Tileset(i)))
            Set DDS_Tileset(i) = Nothing
        End If
    Next
    
    
    
End Sub

Sub DrawOverworld(ByVal dir As Long, ByVal frame As Long, ByVal x As Long, ByVal y As Long, ByVal pokenum As Long)
Dim i As Long
Dim rec As DxVBLib.RECT
'
    Select Case dir
    Case DIR_DOWN
    If DDS_DownFrame(pokenum) Is Nothing Then
    Call InitDDSurf("Overworld\NewDown\" & pokenum, DDSD_DownFrame(pokenum), DDS_DownFrame(pokenum))
    Call InitDDSurf("Overworld\NewDown\Frame2\" & pokenum, DDSD_DownFrame2(pokenum), DDS_DownFrame2(pokenum))
    End If
    Case DIR_UP
    If DDS_UpFrame(pokenum) Is Nothing Then
    Call InitDDSurf("Overworld\NewUp\" & pokenum, DDSD_UpFrame(pokenum), DDS_UpFrame(pokenum))
    Call InitDDSurf("Overworld\NewUp\Frame2\" & pokenum, DDSD_UpFrame2(pokenum), DDS_UpFrame2(pokenum))
    End If
    Case DIR_LEFT
    If DDS_LeftFrame(pokenum) Is Nothing Then
    '
    Call InitDDSurf("Overworld\NewLeft\" & pokenum, DDSD_LeftFrame(pokenum), DDS_LeftFrame(pokenum))
    Call InitDDSurf("Overworld\NewLeft\Frame2\" & pokenum, DDSD_LeftFrame2(pokenum), DDS_LeftFrame2(pokenum))
    End If
    Case DIR_RIGHT
    If DDS_RightFrame(pokenum) Is Nothing Then
  '
    Call InitDDSurf("Overworld\NewRight\" & pokenum, DDSD_RightFrame(pokenum), DDS_RightFrame(pokenum))
    Call InitDDSurf("Overworld\NewRight\Frame2\" & pokenum, DDSD_RightFrame2(pokenum), DDS_RightFrame2(pokenum))
    End If
    End Select
    
    
'
With rec
.Top = 0
.Bottom = DDSD_DownFrame(pokenum).lHeight
.Left = 0
.Right = DDSD_DownFrame(pokenum).lWidth
End With

'
Select Case dir
''''
Case DIR_DOWN
If frame = 2 Then
Call Engine_BltFast(x, y, DDS_DownFrame2(pokenum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
Else
Call Engine_BltFast(x, y, DDS_DownFrame(pokenum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End If
''''
Case DIR_UP
If frame = 2 Then
Call Engine_BltFast(x, y, DDS_UpFrame2(pokenum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
Else
Call Engine_BltFast(x, y, DDS_UpFrame(pokenum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End If
'''''
Case DIR_LEFT
If frame = 2 Then
Call Engine_BltFast(x, y, DDS_LeftFrame2(pokenum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
Else
Call Engine_BltFast(x, y, DDS_LeftFrame(pokenum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End If
'''''
Case DIR_RIGHT
If frame = 2 Then
Call Engine_BltFast(x, y, DDS_RightFrame2(pokenum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
Else
Call Engine_BltFast(x, y, DDS_RightFrame(pokenum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End If
'''''
End Select
End Sub
Sub DrawBattleText()
If inBattle = False Then Exit Sub
If BattlePokemon > 0 Then
If PokemonInstance(BattlePokemon).PokemonNumber > 0 Then
DrawText TexthDC, Camera.Left + 20 + 45, Camera.Top + 91 + 110, "Lvl." & Trim$(PokemonInstance(BattlePokemon).Level) & "      " & Trim$(Pokemon(PokemonInstance(BattlePokemon).PokemonNumber).Name), QBColor(White)
DrawText TexthDC, Camera.Left + 20 + 45, Camera.Top + 91 + 125, "HP " & PokemonInstance(BattlePokemon).HP & "/" & PokemonInstance(BattlePokemon).MaxHp, QBColor(White)
End If
End If
If enemyPokemon.PokemonNumber > 0 Then
DrawText TexthDC, Camera.Left + 20 + 300, Camera.Top + 91 + 10, "Lvl." & Trim$(enemyPokemon.Level) & "      " & Trim$(Pokemon(enemyPokemon.PokemonNumber).Name), QBColor(White)
DrawText TexthDC, Camera.Left + 20 + 300, Camera.Top + 91 + 25, "HP " & enemyPokemon.HP & "/" & enemyPokemon.MaxHp, QBColor(White)
End If
End Sub
