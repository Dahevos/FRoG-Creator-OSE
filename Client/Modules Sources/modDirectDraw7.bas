Attribute VB_Name = "modDirectDraw7"
'Numéro pour Sprites
Public NumCharacters As Long

Public Sub DestroyDirectDraw()
    Dim i As Long
    
    ' Unload DirectDraw
    Set DDS_Misc = Nothing
    

    For i = 1 To NumCharacters
        Set DDS_Character(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Character(i)), LenB(DDSD_Character(i))
    Next
    

    Set DDS_BackBuffer = Nothing
    Set DDS_Primary = Nothing
    Set DD_Clip = Nothing
    Set DD = Nothing
End Sub

Private Sub BltSprite(ByVal Sprite As Long, ByVal x2 As Long, y2 As Long, rec As DxVBLib.RECT)
    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    Dim x As Long
    Dim y As Long
    Dim width As Long
    Dim height As Long
    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("GFX\" & Sprite, DDSD_Character, DDS_Character) 'recherche un fichier de sprite (exemple: Sprite1.bmp)
    End If

    x = ConvertMapX(x2)
    y = ConvertMapY(y2)
    width = (rec.Right - rec.Left)
    height = (rec.Bottom - rec.Top)

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

    If y + height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (y + height - DDSD_BackBuffer.lHeight)
    End If

    If x + width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (x + width - DDSD_BackBuffer.lWidth)
    End If
    ' FIN clipping
    
    Call Engine_BltFast(x, y, DDS_Character(Sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub NewCharacterBltSprite()
    Dim Sprite As Long
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    Dim width As Long, height As Long
    
    If frmMenu.cmbClass.ListIndex = -1 Then Exit Sub
    
    If newCharClass = 0 Then
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).MaleSprite(newCharSprite)
    Else
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).FemaleSprite(newCharSprite)
    End If
    
    If Sprite < 1 Or Sprite > NumCharacters Then
        frmMenu.picSprite.Cls
        Exit Sub
    End If
    
    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("GFX\" & Sprite, DDSD_Character, DDS_Character) ' Va chercher dans le dossier GFX le numéro du sprite (exemple: Sprite1.bmp)
    End If
    
    width = DDSD_Character.lWidth / 12
    height = DDSD_Character.lHeight
    
    frmMenu.picSprite.width = width
    frmMenu.picSprite.height = height
    
    sRECT.Top = 0
    sRECT.Bottom = height
    sRECT.Left = width * 7 'looking down
    sRECT.Right = sRECT.Left + width
    
    dRECT.Top = 0
    dRECT.Bottom = height
    dRECT.Left = 0
    dRECT.Right = width
    
    Call Engine_BltToDC(DDS_Character(Sprite), sRECT, dRECT, frmMenu.picSprite)
End Sub
