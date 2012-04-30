Attribute VB_Name = "modDirectX"
Option Explicit

'Le code utilise pour l'alpha blending est modifie a partir
'du code de Matt Hafermann(http://www.rookscape.com/vbgaming/tutBT.php)

Public Const TilesInSheets As Byte = 14

Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public ExtraSheets As Long

Public dX As New DirectX7
Public DD As DirectDraw7

Public D3D As Direct3D7
Public Dev As Direct3DDevice7

Public DD_Clip As DirectDrawClipper

Public DD_PrimarySurf As DirectDrawSurface7
Public DDSD_Primary As DDSURFACEDESC2

Public DD_SpriteSurf() As DirectDrawSurface7
Public DDSD_Character() As DDSURFACEDESC2
Public SpriteTimer() As Long
Public SpriteUsed() As Boolean

Public DD_ItemSurf As DirectDrawSurface7
Public DDSD_Item As DDSURFACEDESC2

Public DD_EmoticonSurf As DirectDrawSurface7
Public DDSD_Emoticon As DDSURFACEDESC2

Public DD_BackBuffer As DirectDrawSurface7
Public DDSD_BackBuffer As DDSURFACEDESC2

Public DD_SpellAnim() As DirectDrawSurface7
Public DDSD_SpellAnim() As DDSURFACEDESC2
Public SpellTimer() As Long
Public SpellUsed() As Boolean

Public DD_BigSpellAnim() As DirectDrawSurface7
Public DDSD_BigSpellAnim() As DDSURFACEDESC2
Public BigSpellTimer() As Long
Public BigSpellUsed() As Boolean

Public DD_TileSurf() As DirectDrawSurface7
Public DDSD_Tile() As DDSURFACEDESC2
Public TileFile() As Boolean

Public DDSD_ArrowAnim As DDSURFACEDESC2
Public DD_ArrowAnim As DirectDrawSurface7

Public DDSD_Outil As DDSURFACEDESC2
Public DD_OutilSurf As DirectDrawSurface7

'PAPERDOLL

Public DD_PaperDollSurf() As DirectDrawSurface7
Public DDSD_PaperDoll() As DDSURFACEDESC2
Public PaperDollTimer() As Long
Public PaperDollUsed() As Boolean

'FIN PAPERDOLL

Public DD_PetsSurf() As DirectDrawSurface7
Public DDSD_Pets() As DDSURFACEDESC2
Public PetTimer() As Long
Public PetUsed() As Boolean

Public DDSD_Blood As DDSURFACEDESC2
Public DD_Blood As DirectDrawSurface7

Public DDSD_PanoInf As DDSURFACEDESC2
Public DD_PanoInfSurf As DirectDrawSurface7

Public DDSD_PanoSup As DDSURFACEDESC2
Public DD_PanoSupSurf As DirectDrawSurface7

Public DDSD_Night As DDSURFACEDESC2
Public DD_NightSurf As DirectDrawSurface7
Public NightVerts(3) As D3DTLVERTEX

Public DDSD_Fog As DDSURFACEDESC2
Public DD_FogSurf As DirectDrawSurface7
Public FogVerts(3) As D3DTLVERTEX

Public DDSD_Tmp As DDSURFACEDESC2
Public DD_TmpSurf As DirectDrawSurface7

Public rec As RECT
Public rec_pos As RECT

Public AlphaBlendDXIsInit As Boolean
Public ABDXWidth As Integer
Public ABDXHeight As Integer
Public ABDXAlpha As Single
Public Const SurfaceTimerMax As Long = 30000

Sub InitDirectX()
    
    Set DD = dX.DirectDrawCreate(vbNullString)
    
    AlphaBlendDXIsInit = False
    
    ' Indicate windows mode application
    Call DD.SetCooperativeLevel(frmMirage.hwnd, DDSCL_NORMAL)
    
    ' Init type and get the primary surface
    DDSD_Primary.lFlags = DDSD_CAPS
    DDSD_Primary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set DD_PrimarySurf = DD.CreateSurface(DDSD_Primary)
    
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    
    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmMirage.picScreen.hwnd
        
    ' Have the blits to the screen clipped to the picture box
    DD_PrimarySurf.SetClipper DD_Clip
    
    'Initisalisation de D3D
    Set D3D = DD.GetDirect3D
    
End Sub
Function LoadMaxSprite() As Long
Dim i As Long
    
i = 0
Do While FileExiste("\GFX\Sprites\Sprites" & i & ".png")
LoadMaxSprite = i
i = i + 1
Loop

End Function

Function LoadMaxPaperdolls() As Long
Dim i As Long
    
i = 0
Do While FileExiste("\GFX\Paperdolls\Paperdolls" & i & ".png")
LoadMaxPaperdolls = i
i = i + 1
Loop

End Function

Function LoadMaxSpells() As Long
Dim i As Long
    
i = 0
Do While FileExiste("\GFX\Spells\Spells" & i & ".png")
LoadMaxSpells = i
i = i + 1
Loop

End Function

Function LoadMaxBigSpells() As Long
Dim i As Long

i = 0
Do While FileExiste("\GFX\BigSpells\BigSpells" & i & ".png")
LoadMaxBigSpells = i
i = i + 1
Loop

End Function

Function LoadMaxPet() As Long
Dim i As Long

i = 0
Do While FileExiste("\GFX\Pets\Pet" & i & ".png")
LoadMaxPet = i
i = i + 1
Loop
    
    
End Function

Sub InitBackBuffer()
Dim Key As DDCOLORKEY
Dim i As Long

    ' Check for files existing
    'If Not FileExiste("\GFX\sprites.png") Or Not
   
    
    ' Set the key for masks
    Key.low = 0
    Key.high = 0
    
    ' Initialize back buffer
    DDSD_BackBuffer.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_BackBuffer.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    DDSD_BackBuffer.lWidth = (MAX_MAPX + 1) * PIC_X
    DDSD_BackBuffer.lHeight = (MAX_MAPY + 1) * PIC_Y
    Set DD_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
            
   
    
    'Initisialisation de la surface temporaire
    DDSD_Tmp.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Tmp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE
    'DDSCAPS_3DDEVICE est utilise pour pouvoir utiliser D3D sur cette surface
    DDSD_Tmp.lWidth = (MAX_MAPX + 1) * PIC_X
    DDSD_Tmp.lHeight = (MAX_MAPY + 1) * PIC_Y
    Set DD_TmpSurf = DD.CreateSurface(DDSD_Tmp)
    Call DD_TmpSurf.SetForeColor(0)
    Call DD_TmpSurf.setDrawStyle(5)

End Sub
Public Sub UnloadTextures()
    Dim i As Long
    For i = 1 To MaxSprite
            If SpriteTimer(i) > GetTickCount + SurfaceTimerMax Then
                Set DD_SpriteSurf(i) = Nothing
                Call ZeroMemory(DDSD_Character(i), Len(DDSD_Character(i)))
                SpriteTimer(i) = 0
                SpriteUsed(i) = False
            End If
    Next
    
    For i = 1 To MaxSpell
            If SpellTimer(i) > GetTickCount + SurfaceTimerMax Then
                Set DD_SpellAnim(i) = Nothing
                Call ZeroMemory(DDSD_SpellAnim(i), Len(DDSD_SpellAnim(i)))
                SpellTimer(i) = 0
                SpellUsed(i) = False
            End If
    Next
    
    For i = 1 To MaxBigSpell
            If BigSpellTimer(i) > GetTickCount + SurfaceTimerMax Then
                Set DD_BigSpellAnim(i) = Nothing
                Call ZeroMemory(DDSD_BigSpellAnim(i), Len(DDSD_BigSpellAnim(i)))
                BigSpellTimer(i) = 0
                BigSpellUsed(i) = False
            End If
    Next
    For i = 1 To MaxPaperdoll
            If PaperDollTimer(i) > GetTickCount + SurfaceTimerMax Then
                Set DD_PaperDollSurf(i) = Nothing
                Call ZeroMemory(DDSD_PaperDoll(i), Len(DDSD_PaperDoll(i)))
                PaperDollTimer(i) = 0
                PaperDollUsed(i) = False
            End If
    Next
    
    For i = 1 To MaxPet
            If PetTimer(i) > GetTickCount + SurfaceTimerMax Then
                Set DD_PetsSurf(i) = Nothing
                Call ZeroMemory(DDSD_Pets(i), Len(DDSD_Pets(i)))
                PetTimer(i) = 0
                PetUsed(i) = False
            End If
    Next
    
End Sub
Sub PrepareSprite(i As Long)
If SpriteUsed(i) = False Then
If FileExiste("\GFX\Sprites\Sprites" & i & ".png") Then
DDSD_Character(i).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
DDSD_Character(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
Set DD_SpriteSurf(i) = LoadImage(App.Path & "\GFX\Sprites\Sprites" & i & ".png", DD, DDSD_Character(i))
SetMaskColorFromPixel DD_SpriteSurf(i), 0, 0
SpriteUsed(i) = True
End If
End If
SpriteTimer(i) = GetTickCount
End Sub
Sub PrepareSpell(i As Long)
If SpellUsed(i) = False Then
If FileExiste("\GFX\Spells\Spells" & i & ".png") Then
DDSD_SpellAnim(i).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
DDSD_SpellAnim(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
Set DD_SpellAnim(i) = LoadImage(App.Path & "\GFX\Spells\Spells" & i & ".png", DD, DDSD_SpellAnim(i))
SetMaskColorFromPixel DD_SpellAnim(i), 0, 0
SpellUsed(i) = True
End If
End If
SpellTimer(i) = GetTickCount
End Sub
Sub PrepareBigSpell(i As Long)
If BigSpellUsed(i) = False Then
If FileExiste("\GFX\BigSpells\BigSpells" & i & ".png") Then
DDSD_BigSpellAnim(i).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
DDSD_BigSpellAnim(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
Set DD_BigSpellAnim(i) = LoadImage(App.Path & "\GFX\BigSpells\BigSpells" & i & ".png", DD, DDSD_BigSpellAnim(i))
SetMaskColorFromPixel DD_BigSpellAnim(i), 0, 0
BigSpellUsed(i) = True
End If
End If
BigSpellTimer(i) = GetTickCount
End Sub
Sub PreparePaperDoll(i As Long)
If PaperDollUsed(i) = False Then
If FileExiste("\GFX\Paperdolls\Paperdolls" & i & ".png") Then
DDSD_PaperDoll(i).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
DDSD_PaperDoll(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Set DD_PaperDollSurf(i) = LoadImage(App.Path & "\GFX\Paperdolls\Paperdolls" & i & ".png", DD, DDSD_PaperDoll(i))
SetMaskColorFromPixel DD_PaperDollSurf(i), 0, 0
PaperDollUsed(i) = True
End If
End If
PaperDollTimer(i) = GetTickCount
End Sub
Sub PreparePet(i As Long)
If PetUsed(i) = False Then
If FileExiste("\GFX\Pets\Pet" & i & ".png") Then
DDSD_Pets(i).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
DDSD_Pets(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
Set DD_PetsSurf(i) = LoadImage(App.Path & "\GFX\Pets\Pet" & i & ".png", DD, DDSD_Pets(i))
SetMaskColorFromPixel DD_PetsSurf(i), 0, 0
PetUsed(i) = True
End If
End If
PetTimer(i) = GetTickCount
End Sub
Sub InitSurfaces()
Dim i As Integer
 If Not FileExiste("\GFX\items.png") Or Not FileExiste("\GFX\emoticons.png") Or Not FileExiste("\GFX\Outils.png") Or Not FileExiste("\GFX\arrows.png") Then Call MsgBox("Plusieur fichier manquants", vbOKOnly, GAME_NAME): Call GameDestroy
    DDSD_Outil.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Outil.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_OutilSurf = LoadImage(App.Path & "\GFX\Outils.png", DD, DDSD_Outil)
    SetMaskColorFromPixel DD_OutilSurf, 0, 0
    
    ' Init tiles ddsd type and load the bitmap
    For i = 0 To ExtraSheets
        If FileExiste("\GFX\tiles" & i & ".png") Then
            DDSD_Tile(i).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
            DDSD_Tile(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
            Set DD_TileSurf(i) = LoadImage(App.Path & "\GFX\tiles" & i & ".png", DD, DDSD_Tile(i))
            SetMaskColorFromPixel DD_TileSurf(i), 0, 0
            TileFile(i) = True
        Else
            TileFile(i) = False
        End If
        DoEvents
    Next
    
    ' Init items ddsd type and load the bitmap
    DDSD_Item.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Item.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_ItemSurf = LoadImage(App.Path & "\GFX\items.png", DD, DDSD_Item)
    SetMaskColorFromPixel DD_ItemSurf, 0, 0
        
    ' Init emoticons ddsd type and load the bitmap
    DDSD_Emoticon.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Emoticon.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_EmoticonSurf = LoadImage(App.Path & "\GFX\emoticons.png", DD, DDSD_Emoticon)
    SetMaskColorFromPixel DD_EmoticonSurf, 0, 0
        
    ' Init arrows ddsd type and load the bitmap
    DDSD_ArrowAnim.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_ArrowAnim.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_ArrowAnim = LoadImage(App.Path & "\GFX\arrows.png", DD, DDSD_ArrowAnim)
    SetMaskColorFromPixel DD_ArrowAnim, 0, 0
    
    'Chargement de la planche de sang
    DDSD_Blood.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Blood.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_Blood = LoadImage(App.Path & "\GFX\blood.png", DD, DDSD_Blood)
    SetMaskColorFromPixel DD_Blood, 0, 0

End Sub
Sub DestroyDirectX()
Dim i As Long

    Set dX = Nothing
    Set DD = Nothing
    Set DD_PrimarySurf = Nothing
    For i = 0 To LoadMaxSprite()
        Set DD_SpriteSurf(i) = Nothing
    Next i
    For i = 0 To ExtraSheets
        If TileFile(i) Then Set DD_TileSurf(i) = Nothing
    Next i
    Set DD_ItemSurf = Nothing
    Set DD_EmoticonSurf = Nothing
    For i = 0 To LoadMaxSpells()
        Set DD_SpellAnim(i) = Nothing
    Next i
    For i = 0 To LoadMaxBigSpells()
        Set DD_BigSpellAnim(i) = Nothing
    Next i
    Set DD_ArrowAnim = Nothing
    Set DD_PanoInfSurf = Nothing
    Set DD_PanoSupSurf = Nothing
    Set DD_FogSurf = Nothing
    Set DD_NightSurf = Nothing
    Call DestroyAlphaBlendDX
End Sub

Function NeedToRestoreSurfaces() As Boolean
    Dim TestCoopRes As Long
    
    TestCoopRes = DD.TestCooperativeLevel
    
    If (TestCoopRes = DD_OK) Then NeedToRestoreSurfaces = False Else NeedToRestoreSurfaces = True
End Function

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
.low = TheSurface.GetLockedPixel(x, y)
.high = .low
End With

TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey

TheSurface.Unlock TmpR
End Sub

Sub DisplayFx(intX As Long, intY As Long, intWidth As Long, intHeight As Long, lngROP As Long, blnFxCap As Boolean, Tile As Long, ByRef lngDestDC As Long, ByRef lngSrcDC As Long)
    BitBlt lngDestDC, intX, intY, intWidth, intHeight, lngSrcDC, (Tile - (Tile \ TilesInSheets) * TilesInSheets) * PIC_X, (Tile \ TilesInSheets) * PIC_Y, lngROP
End Sub

Sub Night(ByVal MinX As Long, ByVal MaxX As Long, ByVal MinY As Long, ByVal MaxY As Long)
Dim x As Long, y As Long
    If AccOpt.LowEffect Then
        Dim lngSrcDC As Long
        
        lngSrcDC = DD_OutilSurf.GetDC
        For y = MinY To MaxY
            For x = MinX To MaxX
                If Map(GetPlayerMap(MyIndex)).Tile(x, y).Light <= 0 Then
                    DisplayFx (x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 31, TexthDC, lngSrcDC
                Else
                    DisplayFx (x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sy - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, Map(GetPlayerMap(MyIndex)).Tile(x, y).Light, TexthDC, lngSrcDC
                End If
            Next x
        Next y
        DD_OutilSurf.ReleaseDC lngSrcDC
    Else
        'Initialisation du RECT source
        'Initialisation du RECT source
        With rec_pos
            .Top = 0
            .Bottom = (MaxY - MinY + 1) * PIC_Y
            .Left = 0
            .Right = .Left + (MaxX - MinX + 1) * PIC_X
        End With
        
        'Initialisation du RECT destination
        With rec
            .Top = -PIC_Y + (NewPlayerY * 32) + NewYOffset
            .Bottom = .Top + rec_pos.Bottom
            .Left = -PIC_X + (NewPlayerX * 32) + NewXOffset
            .Right = .Left + (MaxX - MinX + 1) * PIC_X
        End With
        
        'Dessin de la nuit
        Call AlphaBlendDX(rec_pos, rec, NightVerts)
    End If
End Sub

Sub BltWeather()
Dim i As Long

    Call DD_BackBuffer.SetForeColor(RGB(0, 0, 200))
    
    If GameWeather = WEATHER_RAINING Or GameWeather = WEATHER_THUNDER Then
        For i = 1 To MAX_RAINDROPS
            If DropRain(i).Randomized = False Then
                If frmMirage.tmrRainDrop.Enabled = False Then
                    BLT_RAIN_DROPS = 1
                    frmMirage.tmrRainDrop.Enabled = True
                    If frmMirage.tmrRainDrop.Tag = vbNullString Then frmMirage.tmrRainDrop.Interval = 200: frmMirage.tmrRainDrop.Tag = "123"
                End If
            End If
        Next i
    ElseIf GameWeather = WEATHER_SNOWING Then
        For i = 1 To MAX_RAINDROPS
            If DropSnow(i).Randomized = False Then
                If frmMirage.tmrSnowDrop.Enabled = False Then
                    BLT_SNOW_DROPS = 1
                    frmMirage.tmrSnowDrop.Enabled = True
                    If frmMirage.tmrSnowDrop.Tag = vbNullString Then frmMirage.tmrSnowDrop.Interval = 200: frmMirage.tmrSnowDrop.Tag = "123"
                End If
            End If
        Next i
    Else
        If BLT_RAIN_DROPS > 0 And BLT_RAIN_DROPS <= RainIntensity Then Call ClearRainDrop(BLT_RAIN_DROPS)
        frmMirage.tmrRainDrop.Tag = vbNullString
    End If
    
    If GameWeather = WEATHER_NONE Then Exit Sub
    
    For i = 1 To MAX_RAINDROPS
        If Not ((DropRain(i).x = 0) Or (DropRain(i).y = 0)) Then
            DropRain(i).x = DropRain(i).x + DropRain(i).speed
            DropRain(i).y = DropRain(i).y + DropRain(i).speed
            Call DD_BackBuffer.DrawLine(DropRain(i).x, DropRain(i).y, DropRain(i).x + DropRain(i).speed, DropRain(i).y + DropRain(i).speed)
            If (DropRain(i).x > (MAX_MAPX + 1) * PIC_X) Or (DropRain(i).y > (MAX_MAPY + 1) * PIC_Y) Then DropRain(i).Randomized = False
        End If
    Next i
    If TileFile(ExtraSheets) Then
        rec.Top = (14 \ TilesInSheets) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (14 - (14 \ TilesInSheets) * TilesInSheets) * PIC_X
        rec.Right = rec.Left + PIC_X
            
        For i = 1 To MAX_RAINDROPS
            If Not ((DropSnow(i).x = 0) Or (DropSnow(i).y = 0)) Then
                DropSnow(i).x = DropSnow(i).x + DropSnow(i).speed
                DropSnow(i).y = DropSnow(i).y + DropSnow(i).speed
                Call DD_BackBuffer.BltFast(DropSnow(i).x + DropSnow(i).speed, DropSnow(i).y + DropSnow(i).speed, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                If (DropSnow(i).x > (MAX_MAPX + 1) * PIC_X) Or (DropSnow(i).y > (MAX_MAPY + 1) * PIC_Y) Then DropSnow(i).Randomized = False
            End If
        Next i
    End If
        
    ' If it's thunder, make the screen randomly flash white
    If GameWeather = WEATHER_THUNDER Then If Int((100 - 1 + 1) * Rnd) + 1 = 8 Then DD_BackBuffer.SetFillColor RGB(255, 255, 255): Call PlaySound("Thunder.wav"): Call DD_BackBuffer.DrawBox(0, 0, (MAX_MAPX + 1) * PIC_X, (MAX_MAPY + 1) * PIC_Y)
End Sub

Sub RNDRainDrop(ByVal RDNumber As Long)
Start:
    DropRain(RDNumber).x = Int((((MAX_MAPX + 1) * PIC_X) * Rnd) + 1)
    DropRain(RDNumber).y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd) + 1)
    If (DropRain(RDNumber).y > (MAX_MAPY + 1) * PIC_Y / 4) And (DropRain(RDNumber).x > (MAX_MAPX + 1) * PIC_X / 4) Then GoTo Start
    DropRain(RDNumber).speed = Int((10 * Rnd) + 6)
    DropRain(RDNumber).Randomized = True
End Sub

Sub ClearRainDrop(ByVal RDNumber As Long)
On Error Resume Next
    DropRain(RDNumber).x = 0
    DropRain(RDNumber).y = 0
    DropRain(RDNumber).speed = 0
    DropRain(RDNumber).Randomized = False
End Sub

Sub RNDSnowDrop(ByVal RDNumber As Long)
Start:
    With DropSnow(RDNumber)
        .x = Int((((MAX_MAPX + 1) * PIC_X) * Rnd) + 1)
        .y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd) + 1)
        If (.y > (MAX_MAPY + 1) * PIC_Y / 4) And (.x > (MAX_MAPX + 1) * PIC_X / 4) Then GoTo Start
        .speed = Int((10 * Rnd) + 6)
        .Randomized = True
    End With
End Sub

Sub ClearSnowDrop(ByVal RDNumber As Long)
On Error Resume Next
    With DropSnow(RDNumber)
        .x = 0
        .y = 0
        .speed = 0
        .Randomized = False
    End With
End Sub

Sub BltPlayerAnim(ByVal Index As Long)
Dim x As Long, y As Long
    If PlayerAnim(Index, 0) = 0 Then Exit Sub
    Call PrepareSpell(PlayerAnim(Index, 0) - 1)
    x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset
    y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset
    
    If GetTickCount > PlayerAnim(Index, 1) + 500 Then
        PlayerAnim(Index, 2) = PlayerAnim(Index, 2) + 1
        If PlayerAnim(Index, 2) > 11 Then PlayerAnim(Index, 2) = 0
        PlayerAnim(Index, 1) = GetTickCount
    End If
    
    rec.Top = 0 * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = PlayerAnim(Index, 2) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_SpellAnim(PlayerAnim(Index, 0) - 1), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    If PlayerAnim(Index, 3) <> 0 Then
         If PlayerAnim(Index, 1) > PlayerAnim(Index, 3) Then
            PlayerAnim(Index, 0) = 0
            PlayerAnim(Index, 1) = GetTickCount
            PlayerAnim(Index, 2) = 0
            PlayerAnim(Index, 3) = 0
            If PlayerAnim(Index, 4) > 0 Then
                Call SendData("exscript" & SEP_CHAR & PlayerAnim(Index, 4) - 1 & END_CHAR)
            End If
            PlayerAnim(Index, 4) = 0
         End If
    End If
End Sub


Sub BltPlayerInt(ByVal Index As Long)
rec.Top = 6 * PIC_Y + (PIC_Y / 2)
rec.Bottom = rec.Top + (PIC_Y / 2)
rec.Left = 0 * PIC_X
rec.Right = rec.Left + (PIC_X * 4)
Call DD_BackBuffer.BltFast(35, 40, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

If Player(Index).HP <> 0 Then
    rec.Top = 6 * PIC_Y
    rec.Bottom = rec.Top + (PIC_Y / 2)
    rec.Left = 0 * PIC_X
    rec.Right = rec.Left + (((PIC_X * 4) * GetPlayerHP(Index)) / GetPlayerMaxHP(Index))
    Call DD_BackBuffer.BltFast(35, 40, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    TexthDC = DD_BackBuffer.GetDC
    
    DrawTextInter TexthDC, 38, 39, Trim$(GetPlayerHP(Index) & " / " & GetPlayerMaxHP(Index))
    Call DD_BackBuffer.ReleaseDC(TexthDC)
End If

rec.Top = 7 * PIC_Y + (PIC_Y / 2)
rec.Bottom = rec.Top + (PIC_Y / 2)
rec.Left = 0 * PIC_X
rec.Right = rec.Left + (PIC_X * 4)
Call DD_BackBuffer.BltFast(35, 60, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

If Player(Index).MP <> 0 Then
    rec.Top = 7 * PIC_Y
    rec.Bottom = rec.Top + (PIC_Y / 2)
    rec.Left = 0 * PIC_X
    rec.Right = rec.Left + (((PIC_X * 4) * GetPlayerMP(Index)) / GetPlayerMaxMP(Index))
    Call DD_BackBuffer.BltFast(35, 60, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    TexthDC = DD_BackBuffer.GetDC
    DrawTextInter TexthDC, 38, 59, Trim$(GetPlayerMP(Index) & " / " & GetPlayerMaxMP(Index))
    Call DD_BackBuffer.ReleaseDC(TexthDC)
End If

rec.Top = 8 * PIC_Y + (PIC_Y / 2)
rec.Bottom = rec.Top + (PIC_Y / 2)
rec.Left = 0 * PIC_X
rec.Right = rec.Left + (PIC_X * 4)
Call DD_BackBuffer.BltFast(35, 80, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

If GetPlayerExp(Index) <> 0 Then
    rec.Top = 8 * PIC_Y
    rec.Bottom = rec.Top + (PIC_Y / 2)
    rec.Left = 0 * PIC_X
    rec.Right = rec.Left + (((PIC_X * 4) * GetPlayerExp(Index)) / nelvl)
    Call DD_BackBuffer.BltFast(35, 80, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    TexthDC = DD_BackBuffer.GetDC
    DrawTextInter TexthDC, 38, 79, Trim$(GetPlayerExp(Index) & " / " & nelvl)
    Call DD_BackBuffer.ReleaseDC(TexthDC)
End If
End Sub


Sub BltSpell(ByVal Index As Long)
Dim x As Long, y As Long, i As Long

If Player(Index).SpellNum <= 0 Or Player(Index).SpellNum > MAX_SPELLS Then Exit Sub
If Spell(Player(Index).SpellNum).SpellAnim < 0 Then Exit Sub
Call PrepareBigSpell(Spell(Player(Index).SpellNum).SpellAnim)
Call PrepareSpell(Spell(Player(Index).SpellNum).SpellAnim)
For i = 1 To MAX_SPELL_ANIM
    With Player(Index).SpellAnim(i)
    If .CastedSpell = YES Then
        If .SpellDone < Spell(Player(Index).SpellNum).SpellDone Then
            If .SpellVar > 10 Then .SpellDone = .SpellDone + 1: .SpellVar = 0
            If GetTickCount > .SpellTime + Spell(Player(Index).SpellNum).SpellTime Then .SpellTime = GetTickCount: .SpellVar = .SpellVar + 1
            
            If Spell(Player(Index).SpellNum).Big > 0 Then
                rec.Top = 0 * (PIC_Y * 2)
                rec.Bottom = rec.Top + (PIC_Y * 2)
                rec.Left = .SpellVar * (PIC_X * 2)
                rec.Right = rec.Left + (PIC_X * 2)

            Else
                rec.Top = 0 * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .SpellVar * PIC_X
                rec.Right = rec.Left + PIC_X
            End If
            
            If .TargetType = 0 Then
                If .Target > 0 Then
                    If .Target = MyIndex Then
                        x = NewX + sx
                        y = NewY + sy
                        If Spell(Player(Index).SpellNum).Big > 0 Then
                            Call DD_BackBuffer.BltFast(x - 16, y - 32, DD_BigSpellAnim(Spell(Player(Index).SpellNum).SpellAnim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                        Else
                            Call DD_BackBuffer.BltFast(x, y, DD_SpellAnim(Spell(Player(Index).SpellNum).SpellAnim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                        End If
                    Else
                        x = GetPlayerX(.Target) * PIC_X + sx + Player(.Target).XOffset
                        y = GetPlayerY(.Target) * PIC_Y + sy + Player(.Target).YOffset
                        
                        If Spell(Player(Index).SpellNum).Big > 0 Then
                            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX - 16, y - NewPlayerPOffsetY - 32, DD_BigSpellAnim(Spell(Player(Index).SpellNum).SpellAnim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                        Else
                            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_SpellAnim(Spell(Player(Index).SpellNum).SpellAnim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                        End If
                    End If
                End If
            Else
                x = MapNpc(.Target).x * PIC_X + sx + MapNpc(.Target).XOffset
                y = MapNpc(.Target).y * PIC_Y + sy + MapNpc(.Target).YOffset
                If Spell(Player(Index).SpellNum).Big > 0 Then
                    Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX - 16, y - NewPlayerPOffsetY - 16, DD_BigSpellAnim(Spell(Player(Index).SpellNum).SpellAnim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_SpellAnim(Spell(Player(Index).SpellNum).SpellAnim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If
        Else
            .CastedSpell = NO
        End If
    End If
    End With
Next i
End Sub
' Effet de sang
Sub BltBlood(ByVal Index As Long, Optional TailleX As Byte, Optional TailleY As Byte, Optional ImgTime As Byte)
Dim x As Integer, y As Integer, i As Integer, BloodNum As Byte


If IsMissing(TailleX) Or TailleX <= 0 Then TailleX = PIC_X
If IsMissing(TailleY) Or TailleY <= 0 Then TailleY = PIC_Y
If IsMissing(ImgTime) Or ImgTime <= 0 Then ImgTime = 40

BloodNum = CByte(Player(Index).BloodAnim.SpellDone)

If Player(Index).BloodAnim.CastedSpell = YES Then
    With Player(Index).BloodAnim
           
            
            rec.Top = BloodNum * TailleY
            rec.Bottom = rec.Top + TailleY
            rec.Left = .SpellVar * TailleX
            rec.Right = rec.Left + TailleX
            
            If .TargetType = 0 Then
                If .Target > 0 Then
                    If .Target = MyIndex Then
                        x = NewX + sx
                        y = NewY + sy
                        Call DD_BackBuffer.BltFast(x, y, DD_Blood, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    Else
                        x = GetPlayerX(.Target) * PIC_X + sx + Player(.Target).XOffset
                        y = GetPlayerY(.Target) * PIC_Y + sy + Player(.Target).YOffset
                        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset - 16, DD_Blood, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If
            Else
                If .Target > 0 And MapNpc(.Target).num > 0 And MapNpc(.Target).HP > 0 Then
                    If Npc(MapNpc(.Target).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(.Target).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(MapNpc(.Target).num).Behavior <> NPC_BEHAVIOR_QUETEUR Then
                        x = MapNpc(.Target).x * PIC_X + sx + MapNpc(.Target).XOffset
                        y = MapNpc(.Target).y * PIC_Y + sy + MapNpc(.Target).YOffset
                        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset - 16, DD_Blood, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If
            End If
            If GetTickCount > .SpellTime + ImgTime Then .SpellTime = GetTickCount: .SpellVar = .SpellVar + 1
            If .SpellVar = 11 Then .CastedSpell = NO: .SpellVar = 0
    End With
End If
End Sub

Sub BltEmoticons(ByVal Index As Long)
Dim x2 As Long, y2 As Long
    If Player(Index).EmoticonNum < 0 Then Exit Sub

    With Player(Index)
        If .EmoticonTime + 1300 > GetTickCount Then
            If GetTickCount >= .EmoticonTime + (108 * (.EmoticonVar + 1)) Then .EmoticonVar = .EmoticonVar + 1
                
            rec.Top = .EmoticonNum * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = .EmoticonVar * PIC_X
            rec.Right = rec.Left + PIC_X
            
            If Index = MyIndex Then
                x2 = NewX + sx + 16
                y2 = NewY + sx - 32
                
                If y2 < 0 Then Exit Sub
                
                Call DD_BackBuffer.BltFast(x2, y2, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Else
                x2 = GetPlayerX(Index) * PIC_X + sx + .XOffset + 16
                y2 = GetPlayerY(Index) * PIC_Y + sx + .YOffset - (PIC_Y * 2)
                
                If y2 < 0 Then Exit Sub
                
                Call DD_BackBuffer.BltFast(x2 - NewPlayerPOffsetX, y2 - NewPlayerPOffsetY, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
    End With
End Sub

Sub BltArrow(ByVal Index As Long)
Dim x As Long, y As Long, i As Long, z As Long
Dim BX As Long, by As Long

For z = 1 To MAX_PLAYER_ARROWS
    With Player(Index).Arrow(z)
    If .Arrow > 0 Then
    
        rec.Top = .ArrowAnim * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = .ArrowPosition * PIC_X
        rec.Right = rec.Left + PIC_X
        
        If GetTickCount > .ArrowTime + 30 Then .ArrowTime = GetTickCount: .ArrowVarX = .ArrowVarX + 10: .ArrowVarY = .ArrowVarY + 10
        
        If .ArrowPosition = 0 Then
            x = .ArrowX
            y = .ArrowY + (.ArrowVarY \ 32)
            If y > .ArrowY + Arrows(.ArrowNum).Range - 2 Then .Arrow = 0
            
            If y <= MAX_MAPY Then Call DD_BackBuffer.BltFast((.ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (.ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset + .ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
        If .ArrowPosition = 1 Then
            x = .ArrowX
            y = .ArrowY - (.ArrowVarY \ 32)
            If y < .ArrowY - Arrows(.ArrowNum).Range + 2 Then .Arrow = 0
                        
            If y >= 0 Then Call DD_BackBuffer.BltFast((.ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (.ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset - .ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
        If .ArrowPosition = 2 Then
            x = .ArrowX + (.ArrowVarX \ 32)
            y = .ArrowY
            If x > .ArrowX + Arrows(.ArrowNum).Range - 2 Then .Arrow = 0
            
            If x <= MAX_MAPX Then Call DD_BackBuffer.BltFast((.ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset + .ArrowVarX, (.ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
        If .ArrowPosition = 3 Then
            x = .ArrowX - (.ArrowVarX \ 32)
            y = .ArrowY
            If x < .ArrowX - Arrows(.ArrowNum).Range + 2 Then .Arrow = 0
                        
            If x >= 0 Then Call DD_BackBuffer.BltFast((.ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset - .ArrowVarX, (.ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
        If x >= 0 And x <= MAX_MAPX Then
            If y >= 0 And y <= MAX_MAPY Then
                If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_BLOCK_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_BLOCK_DIR Then .Arrow = 0
            End If
        End If
        
        For i = 1 To MAX_PLAYERS
           If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                If GetPlayerX(i) = x And GetPlayerY(i) = y Then
                    If Index = MyIndex Then Call SendData("arrowhit" & SEP_CHAR & 0 & SEP_CHAR & i & SEP_CHAR & x & SEP_CHAR & y & END_CHAR)
                    If Index <> i Then .Arrow = 0
                    Exit Sub
                End If
            End If
        Next i
        
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(i).num > 0 Then
                If MapNpc(i).x = x And MapNpc(i).y = y Then
                    If Index = MyIndex Then Call SendData("arrowhit" & SEP_CHAR & 1 & SEP_CHAR & i & SEP_CHAR & x & SEP_CHAR & y & END_CHAR)
                    .Arrow = 0
                    Exit Sub
                End If
            End If
        Next i
    End If
    End With
Next z
End Sub

Sub ChrgSpriteSurf()
Dim i As Long
Set DD = dX.DirectDrawCreate("")
Call DD.SetCooperativeLevel(frmMirage.hwnd, DDSCL_NORMAL)
' Init sprite ddsd type and load the bitmap
For i = 0 To MaxSprite
    DDSD_Character(i).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Character(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    SpriteUsed(i) = False
    'Set DD_SpriteSurf(i) = LoadImage(App.Path & "\GFX\Sprites\Sprites" & i & ".png", DD, DDSD_Character(i))
    'SetMaskColorFromPixel DD_SpriteSurf(i), 0, 0
Next i

End Sub

Public Sub InitPano(ByVal MapNum As Long)
    If MapNum <= 0 Or MapNum > MAX_MAPS Or DD Is Nothing Then Exit Sub
    If Trim$(Map(MapNum).PanoInf) <> vbNullString And InStr(1, Trim$(Map(MapNum).PanoInf), ".png") > 0 Then
        If Not FileExiste("GFX\" & Trim$(Map(MapNum).PanoInf)) Then
            Map(MapNum).PanoInf = vbNullString
        Else
            'Initiialisation de la surface PanoInfSurf
            DDSD_PanoInf.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
            DDSD_PanoInf.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
            Set DD_PanoInfSurf = LoadImage(App.Path & "\GFX\" & Map(MapNum).PanoInf, DD, DDSD_PanoInf)
            If Map(MapNum).TranInf = 1 Then SetMaskColorFromPixel DD_PanoInfSurf, 0, 0
        End If
    End If
    
    If Trim$(Map(MapNum).PanoSup) <> vbNullString And InStr(1, Trim$(Map(MapNum).PanoSup), ".png") > 0 Then
        If Not FileExiste("GFX\" & Trim$(Map(MapNum).PanoSup)) Then
            Map(MapNum).PanoSup = vbNullString
        Else
            'Initiialisation de la surface PanoSupSurf
            DDSD_PanoSup.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
            DDSD_PanoSup.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
            Set DD_PanoSupSurf = LoadImage(App.Path & "\GFX\" & Map(MapNum).PanoSup, DD, DDSD_PanoSup)
            If Map(MapNum).TranSup = 1 Then SetMaskColorFromPixel DD_PanoSupSurf, 0, 0
        End If
    End If
End Sub

Public Sub InitNightAndFog(ByVal MapNum As Long)
    If MapNum <= 0 Or MapNum > MAX_MAPS Or DD Is Nothing Then Exit Sub
    
    'Initialisation du brouillard
    If Map(MapNum).Fog <> 0 Then
        'Initialisation de la texture pour le brouillard si il y en a un
        DDSD_Fog.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_CKSRCBLT
        DDSD_Fog.ddsCaps.lCaps = DDSCAPS_TEXTURE
        DDSD_Fog.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
        Set DD_FogSurf = LoadImage(App.Path & "\GFX\fog" & Map(MapNum).Fog & ".png", DD, DDSD_Fog)
        
        'Initialisation du RECT de la texture
        With rec
            .Top = 0
            .Bottom = DDSD_Fog.lHeight
            .Left = 0
            .Right = DDSD_Fog.lWidth
        End With
        
        'Initialisation de l'Alpha Blending
        Call InitAlphaBlendDX(DD_TmpSurf, DD_FogSurf, rec, FogVerts(), Map(MapNum).FogAlpha / 100)
    Else
        Set DD_FogSurf = Nothing
    End If
    
    'Initialisation de la nuit
    If GameTime = TIME_NIGHT Then
        Dim x As Long
        Dim y As Long
        Dim Tile As Long
                
        'Initialisation de la texture pour la nuit
        DDSD_Night.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_CKSRCBLT
        DDSD_Night.ddsCaps.lCaps = DDSCAPS_TEXTURE
        DDSD_Night.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
        DDSD_Night.lWidth = (MAX_MAPX + 1) * PIC_X
        DDSD_Night.lHeight = (MAX_MAPY + 1) * PIC_Y
        Set DD_NightSurf = DD.CreateSurface(DDSD_Night)
        
        'Dessin de l'effet de la nuit en "low effect"
        Tile = 31
        With rec
            .Top = (Tile \ TilesInSheets) * PIC_Y
            .Bottom = .Top + PIC_X
            .Left = (Tile - (Tile \ TilesInSheets) * TilesInSheets) * PIC_X
            .Right = .Left + PIC_Y
        End With
        DD_NightSurf.BltFast 0, 0, DD_OutilSurf, rec, DDBLTFAST_WAIT
        
        'Récupération de la couleur de la nuit en "low effect" en changeant le colokey du Tiles ou sont les effets
        SetMaskColorFromPixel DD_OutilSurf, (Tile - (Tile \ TilesInSheets) * TilesInSheets) * PIC_X, (Tile \ TilesInSheets) * PIC_Y
        
        'Remplissage de la surface en noir
        With rec
            .Top = 0
            .Bottom = (MAX_MAPX + 1) * PIC_X
            .Left = 0
            .Right = (MAX_MAPY + 1) * PIC_Y
        End With
        Call DD_NightSurf.BltColorFill(rec, 0)
        
        'Dessin des lumiéres
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                If Map(GetPlayerMap(MyIndex)).Tile(x, y).Light > 0 Then
                    Tile = Map(GetPlayerMap(MyIndex)).Tile(x, y).Light
                    With rec
                        .Top = (Tile \ TilesInSheets) * PIC_Y
                        .Bottom = .Top + PIC_X
                        .Left = (Tile - (Tile \ TilesInSheets) * TilesInSheets) * PIC_X
                        .Right = .Left + PIC_Y
                    End With
                    DD_NightSurf.BltFast x * PIC_X, y * PIC_Y, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            Next x
        Next y
        
        'Remettre la colorkey de base
        SetMaskColorFromPixel DD_OutilSurf, 0, 0
                
        'Initialisation du RECT de la texture
        With rec
            .Top = 0
            .Bottom = DDSD_Night.lHeight
            .Left = 0
            .Right = DDSD_Night.lWidth
        End With
        
        'Initialisation de l'Alpha Blending
        Call InitAlphaBlendDX(DD_TmpSurf, DD_NightSurf, rec, NightVerts(), 0.5)
    Else
        Set DD_NightSurf = Nothing
    End If
End Sub

Public Sub DestroyAlphaBlendDX()
    'Desactivation de l'Alpha Blending
    Dev.SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, False
    
    'Destruction de D3D et de sont Device
    Set Dev = Nothing
    Set D3D = Nothing
    
    AlphaBlendDXIsInit = False
End Sub

Public Sub InitAlphaBlendDX(DSurf As DirectDrawSurface7, TSurf As DirectDrawSurface7, Vrec As RECT, DVerts() As D3DTLVERTEX, ByVal Alpha As Single)
    'Verification de l'initialisation de DirectX
    If DD Is Nothing Then Exit Sub
    
    'Verification de la surface
    If DSurf Is Nothing Then Exit Sub
    
    'Verification de l'alpha
    If Alpha < 0 Or Alpha > 1 Then Alpha = 0.5
    
    'Initisalisation de D3D si il ne l'est pas
    If D3D Is Nothing Then Set D3D = DD.GetDirect3D
    
    'Initialisation du Device avec la surface de destination
    Set Dev = D3D.CreateDevice("IID_IDirect3DHALDevice", DSurf)
    
    'Activation de l'Alpha Blending
    Dev.SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, True
    
    'Initialisation des parametres du Device
    Dev.SetRenderState D3DRENDERSTATE_SRCBLEND, D3DBLEND_SRCALPHA
    Dev.SetRenderState D3DRENDERSTATE_DESTBLEND, D3DBLEND_INVSRCALPHA
    Dev.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTA_TFACTOR
    
    'Definition de la texture utiliser par le Device
    Dev.SetTexture 0, TSurf
    Dev.SetTextureStageState 0, D3DTSS_MIPFILTER, 3
    
    'Initialisation des variable pour la taille de la texture et l'alpha
    ABDXHeight = Vrec.Bottom
    ABDXWidth = Vrec.Right
    ABDXAlpha = Alpha
    
    'Initialisation du vecteur qui permet de dessiner la texture ou il faut sur la surface
    Call SetUpVerts(DVerts(), Vrec, Vrec, Alpha)
    
    AlphaBlendDXIsInit = True
End Sub

Public Sub AlphaBlendDX(Srec As RECT, Drec As RECT, DVerts() As D3DTLVERTEX)
    'Verification de l'initialisation de l'Alpha Blending
    If Not AlphaBlendDXIsInit Then Exit Sub

    'Initialisation du vecteur qui permet de dessiner la texture ou il faut sur la surface
    Call SetUpVerts(DVerts(), Drec, Srec, ABDXAlpha)
    
    'Debute la scene D3D (obligatoire avant tous dessin)
    Dev.BeginScene

    'Récupération du contenu du BackBuffer
    Call DD_TmpSurf.BltFast(Srec.Left, Srec.Top, DD_BackBuffer, Srec, DDBLTFAST_WAIT)
    
    'Dessin de la texture
    Dev.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, DVerts(0), 4, D3DDP_DEFAULT
    
    'Fin de la scene D3D (obligatoire apres tous dessin)
    Dev.EndScene
    
    'Dessin du résultat sur le BackBuffer
    Call DD_BackBuffer.BltFast(Srec.Left, Srec.Top, DD_TmpSurf, Srec, DDBLTFAST_WAIT)
End Sub

Public Sub SetUpVerts(Verts() As D3DTLVERTEX, Src As RECT, Dest As RECT, ByVal Alpha As Single)
On Error Resume Next
    'Ce sub permet d'initialiser 4 points qui forme 3 vecteurs
    'pour une texture rectangulaire
    'Les 3 vecteurs forme 2 triangles(polygones) qui forme le rectangle
    'pour la surface, ces points sont définis par plusieurs parametres,
    'leur coordones,taille et leur couleur au format RGBA si R,G ou B sont a 1
    'la surface serat dessine avec un filtre de la/les couleurs correspondants
    'A permet de regler la valeur de l'opaciter de 0 a 1(0=invisible,1=opaque)
    'par exemple si A est a 0.5 la transparence est de 50%
    
    Dim SurfW As Single
    Dim SurfH As Single
    Dim XCor As Single
    Dim YCor As Single
    Dim AlphaColor As Long
    
    'Récupération de la couleur avec l'alpha
    AlphaColor = GetColorPlusAlpha(Alpha)
        
    'Largeur de la surface
    SurfW = ABDXWidth 'Src.Right - Src.Left
    'Hauteur de la surface
    SurfH = ABDXHeight 'Src.Bottom - Src.Top
    
    'Dest.Left = Dest.Left
    'Dest.Top = Dest.Top
    'Dest.Bottom = Dest.Bottom
    'Dest.Right = Dest.Right
   'Coordonees du 1er point sur la surface de destination
    XCor = Dest.Left
    YCor = Dest.Bottom
    
    '1er point - en bas a gauche
    dX.CreateD3DTLVertex _
    XCor, _
    YCor, _
    0, _
    1, _
    AlphaColor, _
    0, _
    Src.Left / SurfW, _
    (Src.Bottom) / SurfH, _
    Verts(0)
    
    'Coordonees du 2eme point sur la surface de destination
    XCor = Dest.Left
    YCor = Dest.Top
    
    '2eme point - en haut a gauche
    dX.CreateD3DTLVertex _
    XCor, _
    YCor, _
    0, _
    1, _
    AlphaColor, _
    0, _
    Src.Left / SurfW, _
    Src.Top / SurfH, _
    Verts(1)
    
    'Coordonees du 3eme point sur la surface de destination
    XCor = Dest.Right
    YCor = Dest.Bottom
    
    '3eme point - en bas a droite
    dX.CreateD3DTLVertex _
    XCor, _
    YCor, _
    0, _
    1, _
    AlphaColor, _
    0, _
    (Src.Right) / SurfW, _
    (Src.Bottom) / SurfH, _
    Verts(2)
    
    'Coordonees du 4eme point sur la surface de destination
    XCor = Dest.Right
    YCor = Dest.Top
    
    '4eme point - en haut a droite
    dX.CreateD3DTLVertex _
    XCor, _
    YCor, _
    0, _
    1, _
    AlphaColor, _
    0, _
    (Src.Right) / SurfW, _
    Src.Top / SurfH, _
    Verts(3)
End Sub

Sub DD_SetD3DSprite(lX As Long, lY As Long, lW As Long, lH As Long, coloring As Long, _
                    tx As Single, ty As Single, ByVal tW As Single, ByVal tH As Single, D3DSprite() As D3DTLVERTEX)
dX.CreateD3DTLVertex lX, lY + lH, 0, 1, coloring, 0, tx, ty + tH, D3DSprite(0)
dX.CreateD3DTLVertex lX, lY, 0, 1, coloring, 0, tx, ty, D3DSprite(1)
dX.CreateD3DTLVertex lX + lW, lY + lH, 0, 1, coloring, 0, tx + tW, ty + tH, D3DSprite(2)
dX.CreateD3DTLVertex lX + lW, lY, 0, 1, coloring, 0, tx + tW, ty, D3DSprite(3)
End Sub

Public Function GetColorPlusAlpha(ByVal Alpha As Single) As Long
    'Simulation de dX.CreateColorRGBA pour R=G=B=1 et pour une valeur d'alpha
    'La fonction doit etre simuler pour la compatibiliter avec vista
    
    'Valeur par defaut
    GetColorPlusAlpha = -1
    
    'Verification du parametre
    If Alpha < 0 Or Alpha > 1 Then Exit Function
    
    If Alpha = 0 Then
        'Valeur definit pour 0
        GetColorPlusAlpha = 16777215
    ElseIf Alpha = 1 Then
        'Valeur definit pour 1
        GetColorPlusAlpha = -1
    ElseIf Alpha <= 0.5 Then
        Dim b As Currency
        Dim Tmp As Currency
        Dim n As Byte
        'On decompose le calcul pour eviter les "depassement de capaciter"
        If (Alpha * 100 - (Fix(Alpha * 10) * 10)) Mod 2 = 0 Then
            Tmp = 16777215
            
            n = (Alpha * 100) - ((Alpha * 100) \ 2) - (Fix(Alpha * 10) \ 2)
            Tmp = Tmp + (n * 33554432)
            
            n = (Alpha * 100) - ((Alpha * 100) \ 2) + (Fix(Alpha * 10) \ 2)
            b = n * 25165824
            Tmp = Tmp + (b * 2)
        Else
            Tmp = 16777215
            
            n = (Alpha * 100) - ((Alpha * 100) \ 2) - ((Fix(Alpha * 10) + 1) \ 2)
            Tmp = Tmp + (n * 33554432)
            
            n = (Alpha * 100) - ((Alpha * 100) \ 2) + ((Fix(Alpha * 10) - 1) \ 2)
            If Fix(Alpha * 10) = 0 Then n = n - 1
            b = n * 25165824
            Tmp = Tmp + (b * 2)
        End If
        GetColorPlusAlpha = CLng(Tmp)
    Else
        GetColorPlusAlpha = (GetColorPlusAlpha(0.5 - (Alpha - 0.5)) + 2) * -1
    End If
End Function
