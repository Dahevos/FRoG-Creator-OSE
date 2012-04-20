Attribute VB_Name = "modGameLogic"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const SRCAND As Long = &H8800C6
Public Const SRCCOPY As Long = &HCC0020
Public Const SRCPAINT As Long = &HEE0086

Public Const VK_UP As Long = &H26
Public Const VK_DOWN As Long = &H28
Public Const VK_LEFT As Long = &H25
Public Const VK_RIGHT As Long = &H27
Public Const VK_SHIFT As Long = &H10
Public Const VK_RETURN As Long = &HD
Public Const VK_CONTROL As Long = &H11

' Menu states
Public Const MENU_STATE_LOGIN As Byte = 2
Public Const MENU_STATE_GETCHARS As Byte = 3
Public Const MENU_STATE_USECHAR As Byte = 7
Public Const MENU_STATE_INIT As Byte = 8

' Speed moving vars
Public Const WALK_SPEED As Byte = 4
Public Const RUN_SPEED As Byte = 8
Public Const GM_WALK_SPEED As Byte = 4
Public Const GM_RUN_SPEED As Byte = 8
'Set the variable to your desire,
'32 is a safe and recommended setting

' Game direction vars
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean

' Multi-Serveur
Public CHECK_WAIT As Boolean

' Game text buffer
Public MyText As String

' Index of actual player
Public MyIndex As Long

' Map animation #, used to keep track of what map animation is currently on
Public MapAnim As Boolean
Public MapAnimTimer As Long

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' Used to check if in editor or not and variables for use in editor
Public InEditor As Boolean
Public InProprieter As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long
Public EditorSet As Byte

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long

' pour les porte a code
Public CodePorte As String

' pour les coffres
Public CodeCoffre As String
Public CleCoffreNum As Long
Public CleCoffreSupr As Long
Public ObjCoffreNum As Long

' pour les block niv
Public NivMin As Long

' pour les block monture
Public MontureRequi As Long

' pour les block guilde
Public NomGuilde As String

' Used for map key opene ditor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long
Public KeyOpenEditorMsg As String

' Map for local use
Public SaveMapItem() As MapItemRec
Public SaveMapNpc(1 To MAX_MAP_NPCS) As MapNpcRec

' Used for index based editors
Public InItemsEditor As Boolean
Public InNpcEditor As Boolean
Public InShopEditor As Boolean
Public InSpellEditor As Boolean
Public InEmoticonEditor As Boolean
Public InArrowEditor As Boolean
Public InMouvEditor As Boolean
Public InToit As Boolean
Public InQuetesEditor As Boolean
Public InCineEditor As Boolean
Public EditorIndex As Long
Public EditorMouvIndex As Long
Public EditorQueteIndex As Long
Public InDefTel As Boolean
Public InDefKey As Boolean
Public InPetsEditor As Boolean
Public InMetierEditor As Boolean
Public InRecetteEditor As Boolean
' Game fps
Public GameFPS As Long

'Loc of pointer
Public CurX As Single '/case
Public CurY As Single '/case
Public PotX As Single 'réel
Public PotY As Single 'réel

' Used for atmosphere
Public GameWeather As Long
Public GameTime As Long
Public RainIntensity As Long

' Scrolling Variables
Public NewPlayerX As Long
Public NewPlayerY As Long
Public NewXOffset As Long
Public NewYOffset As Long
Public NewX As Long
Public NewY As Long
Public NewPlayerPicX As Long
Public NewPlayerPicY As Long
Public NewPlayerPOffsetX As Long
Public NewPlayerPOffsetY As Long

' Damage Variables
Public DmgDamage As Long
Public DmgTime As Long
Public NPCDmgDamage As Long
Public NPCDmgTime As Long
Public NPCWho As Long

Public EditorItemX As Long
Public EditorItemY As Long

Public EditorShopNum As Long

Public EditorItemNum1 As Byte
Public EditorItemNum2 As Byte
Public EditorItemNum3 As Byte

Public AccptDir1 As Long
Public AccptDir2 As Long
Public AccptDir3 As Long

Public Arena1 As Byte
Public Arena2 As Byte
Public Arena3 As Byte

Public ii As Long, iii As Long
Public sx As Long
Public sy As Long

Public MouseDownX As Long
Public MouseDownY As Long

Public SpritePic As Long
Public SpriteItem As Long
Public SpritePrice As Long

Public SoundFileName As String

Public ScreenMode As Boolean

Public SignLine1 As String

Public ClassChange As Long
Public ClassChangeReq As Long

Public NoticeTitle As String
Public NoticeText As String
Public NoticeSound As String

Public ScriptNum As Long
Public Connucted As Boolean

'use for bank
Public bankmsg As String

'pour le sub tester
Public test As Byte

'pour les classes
Public classe As Long

'pour les sauvegarde
Public save As Byte

' pour les mouvement map npc
Public cordo As Long

'pour les quetes
Public Accepter As Boolean

'pour les controlles
Public ConOff As Boolean
Public OldMap As Long
Public NumShop As Long

'pour le zoom
Public VZoom As Byte
Public OldVZoom As Byte
Public ScreenDC As Boolean

'pour le mouvement des fenetre
Public drx As Long
Public dry As Long
Public dr As Boolean

'pour les couleurs personalisables
Public AccModo As Long
Public AccMapeur As Long
Public AccDevelopeur As Long
Public AccAdmin As Long

'Mémoire
Public CoordX As Long
Public CoordY As Long
Public CoordM As Long
Public DonID As Long
Public DonTP As Byte
Public TempNum As Byte

'Mouvement des PNJs
Public PNJAnim(1 To MAX_MAP_NPCS) As Byte

'Sauvegarde automatique
Public SauvAuto As Long

'Carte pas ftp
Public CarteFTP As Boolean

'Variables de FrmMirage
Public PicScWidth As Single
Public PicScHeight As Single

Sub Main()
Dim i As Long
Dim Ending As String
On Error GoTo er:

    If FileExiste("r.exe") Then Kill App.Path & "\r.exe"
    Call InitXpStyle
    Call EcrireEtat(vbNullString)
    Call EcrireEtat("Démarrage du logiciel")
    
    save = 0
    VZoom = 3
    ScreenDC = False
    InProprieter = False
    InDefTel = False
    InDefKey = False
    ScreenMode = False
    DonID = 0
    AccptDir1 = -1
    AccptDir2 = -1
    AccptDir3 = -1
    frmsplash.Show
    Call SetStatus("Vérification des dossiers...")
    DoEvents
    ExtraSheets = 0
    
    Dim PathSource As String, Part() As String
    Part = Split(App.Path, "\")
    If UBound(Part) > 0 Then PathSource = Mid$(App.Path, 1, Len(App.Path) - Len(Part(UBound(Part)))) Else PathSource = App.Path & "\"
    
    If FileExiste("GFX\Tiles0.png") Then
        For i = 0 To 256
            If Not FileExiste("GFX\Tiles" & i & ".png") And i <> 0 Then ExtraSheets = i - 1: Exit For
        Next i
    Else
        For i = 0 To 255
            If Not FileExistes(PathSource & "Client\GFX\Tiles" & i & ".png") And i <> 0 Then ExtraSheets = i - 1: Exit For
        Next i
    End If

    ReDim DD_TileSurf(0 To ExtraSheets) As DirectDrawSurface7
    ReDim DDSD_Tile(0 To ExtraSheets) As DDSURFACEDESC2
    ReDim TileFile(0 To ExtraSheets) As Boolean
    
    Call LoadMaxSprite
    Call LoadMaxPaperdolls
    Call LoadMaxSpells
    Call LoadMaxBigSpells
    Call LoadMaxPet
    
    ReDim DD_SpriteSurf(0 To MAX_DX_SPRITE) As DirectDrawSurface7
    ReDim DDSD_Character(0 To MAX_DX_SPRITE) As DDSURFACEDESC2
        
    ReDim DD_PaperDollSurf(0 To MAX_DX_PAPERDOLL) As DirectDrawSurface7
    ReDim DDSD_PaperDoll(0 To MAX_DX_PAPERDOLL) As DDSURFACEDESC2
    
    ReDim DD_SpellAnim(0 To MAX_DX_SPELLS) As DirectDrawSurface7
    ReDim DDSD_SpellAnim(0 To MAX_DX_SPELLS) As DDSURFACEDESC2
    
    ReDim DD_BigSpellAnim(0 To MAX_DX_BIGSPELLS) As DirectDrawSurface7
    ReDim DDSD_BigSpellAnim(0 To MAX_DX_BIGSPELLS) As DDSURFACEDESC2
    
    ReDim DD_PetsSurf(0 To MAX_DX_PETS) As DirectDrawSurface7
    ReDim DDSD_Pets(0 To MAX_DX_PETS) As DDSURFACEDESC2
    
    ' Check if the maps directory is there, if its not make it
    If LCase$(Dir$(App.Path & "\Classes", vbDirectory)) <> "classes" Then Call MkDir$(App.Path & "\Classes")
    If LCase$(Dir$(App.Path & "\Maps", vbDirectory)) <> "maps" Then Call MkDir$(App.Path & "\Maps")
    If UCase$(Dir$(App.Path & "\GFX", vbDirectory)) <> "GFX" Then Call MkDir$(App.Path & "\GFX")
    If UCase$(Dir$(App.Path & "\Music", vbDirectory)) <> "MUSIC" Then Call MkDir$(App.Path & "\Music")
    If UCase$(Dir$(App.Path & "\SFX", vbDirectory)) <> "SFX" Then Call MkDir$(App.Path & "\SFX")
    If UCase$(Dir$(App.Path & "\Flashs", vbDirectory)) <> "FLASHS" Then Call MkDir$(App.Path & "\Flashs")
    If UCase$(Dir$(App.Path & "\items", vbDirectory)) <> "ITEMS" Then Call MkDir$(App.Path & "\items")
    If UCase$(Dir$(App.Path & "\maps", vbDirectory)) <> "MAPS" Then Call MkDir$(App.Path & "\maps")
    If UCase$(Dir$(App.Path & "\shops", vbDirectory)) <> "SHOPS" Then Call MkDir$(App.Path & "\shops")
    If UCase$(Dir$(App.Path & "\pnjs", vbDirectory)) <> "PNJS" Then Call MkDir$(App.Path & "\pnjs")
    If UCase$(Dir$(App.Path & "\spells", vbDirectory)) <> "SPELLS" Then Call MkDir$(App.Path & "\spells")
    If UCase$(Dir$(App.Path & "\quetes", vbDirectory)) <> "QUETES" Then Call MkDir$(App.Path & "\quetes")
    If UCase$(Dir$(App.Path & "\Config", vbDirectory)) <> "CONFIG" Then Call MkDir$(App.Path & "\Config")
    frmsplash.chrg.value = 10
        
    Call SetStatus("Transfère des données...")
    If Not FileExiste("Config\account.ini") Or UCase$(Dir$(App.Path & "\Music", vbDirectory)) <> "MUSIC" Then
        'Call FileCopy(PathSource & "Client\GFX\Sprites.png", App.Path & "\GFX\Sprites.png")
        Call FileCopy(PathSource & "Client\GFX\Arrows.png", App.Path & "\GFX\Arrows.png")
        'Call FileCopy(PathSource & "Client\GFX\BigSprites.png", App.Path & "\GFX\BigSprites.png")
        Call FileCopy(PathSource & "Client\GFX\Emoticons.png", App.Path & "\GFX\Emoticons.png")
        Call FileCopy(PathSource & "Client\GFX\items.png", App.Path & "\GFX\items.png")
        'Call FileCopy(PathSource & "Client\GFX\Spells.png", App.Path & "\GFX\Spells.png")
        'For i = 0 To ExtraSheets
            'Call FileCopy(PathSource & "Client\GFX\Tiles" & i & ".png", App.Path & "\GFX\Tiles" & i & ".png")
        'Next i
        Dim rep As String
        'SFX deplacement
        'obtient le premier fichier ou répertoire qui est dans "c:\"
        On Error Resume Next
        rep = Dir$(PathSource & "Client\SFX\*.*", vbDirectory)
        'boucle tant que le répertoire n'a pas été entièrement parcouru
        Do While (rep > vbNullString)
            'teste si c'est un fichier ou un répertoire
            If GetAttr(PathSource & "Client\SFX\" & rep) <> vbDirectory Then Call FileCopy(PathSource & "Client\SFX\" & rep, App.Path & "\SFX\" & rep)
            'passe à l'élément suivant
            rep = Dir
            Sleep 1
        Loop
        
        'config deplacement
        On Error Resume Next
        rep = Dir$(PathSource & "Client\Config\*.*", vbDirectory)
        Do While (rep > vbNullString)
            If GetAttr(PathSource & "Client\Config\" & rep) <> vbDirectory Then Call FileCopy(PathSource & "Client\Config\" & rep, App.Path & "\Config\" & rep)
            rep = Dir
            Sleep 1
        Loop
    
       'music deplacement
        On Error Resume Next
        rep = Dir$(PathSource & "Client\Music\*.*", vbDirectory)
        Do While (rep > vbNullString)
            If GetAttr(PathSource & "Client\Music\" & rep) <> vbDirectory Then Call FileCopy(PathSource & "Client\Music\" & rep, App.Path & "\Music\" & rep)
            rep = Dir
            Sleep 1
        Loop
    End If
    frmsplash.chrg.value = 20
    
    Dim FileName As String
    FileName = App.Path & "\Config\Account.ini"
    
    If FileExiste("Config\Account.ini") Then
        With frmoptions
            .chkbubblebar.value = ReadINI("CONFIG", "SpeechBubbles", FileName)
            .chknpcbar.value = ReadINI("CONFIG", "NpcBar", FileName)
            .chknpcname.value = ReadINI("CONFIG", "NPCName", FileName)
            .chkplayerbar.value = ReadINI("CONFIG", "PlayerBar", FileName)
            .chkplayername.value = ReadINI("CONFIG", "PlayerName", FileName)
            .chkplayerdamage.value = ReadINI("CONFIG", "NPCDamage", FileName)
            .chknpcdamage.value = ReadINI("CONFIG", "PlayerDamage", FileName)
            .chkmusic.value = ReadINI("CONFIG", "Music", FileName)
            .chksound.value = ReadINI("CONFIG", "Sound", FileName)
            .chkAutoScroll.value = ReadINI("CONFIG", "AutoScroll", FileName)
            .chknobj.value = Val(ReadINI("CONFIG", "NomObjet", FileName))
            .chkLowEffect.value = Val(ReadINI("CONFIG", "LowEffect", FileName))
        End With
        If Val(ReadINI("CONFIG", "MapGrid", FileName)) = 0 Then frmMirage.grile.Checked = False Else frmMirage.grile.Checked = True
        If Val(ReadINI("CONFIG", "PreVisu", FileName)) = 0 Then frmMirage.previsu.Checked = False Else frmMirage.previsu.Checked = True
    Else
        WriteINI "INFO", "Account", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "INFO", "Password", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CONFIG", "WebSite", "http://www.frogcreator.fr", App.Path & "\Config\Client.ini"
        WriteINI "CONFIG", "Version", "0.5", App.Path & "\Config\Client.ini"
        WriteINI "CONFIG", "auto-maj", "1", App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine1", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine2", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine3", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine4", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine5", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine6", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine7", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine8", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine9", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine10", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine11", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CONFIG", "SpeechBubbles", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "NpcBar", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "NPCName", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "NPCDamage", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "PlayerBar", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "PlayerName", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "PlayerDamage", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "MapGrid", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "PreVisu", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "Music", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "Sound", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "AutoScroll", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "LowEffect", 1, App.Path & "\Config\Account.ini"
    End If
    
    If Not FileExiste("config.ini") Then
        WriteINI "INFO", "PIC_PL", 64, App.Path & "\Config.ini"
        WriteINI "INFO", "PIC_NPC1", 2, App.Path & "\Config.ini"
        WriteINI "INFO", "PIC_NPC2", 32, App.Path & "\Config.ini"
    End If
    frmsplash.chrg.value = 30
    Call InitAccountOpt
    Call InitMirageVars
    
    Call SetStatus("Vérification du status...")
    DoEvents
    
    ' Make sure we set that we aren't in the game
    InGame = False
    GettingMap = True
    InEditor = False
    InItemsEditor = False
    InNpcEditor = False
    InShopEditor = False
    InEmoticonEditor = False
    InArrowEditor = False
    InMouvEditor = False
    InQuetesEditor = False
    frmsplash.chrg.value = 40
    
    If Not FileExiste("Config\Serveur.ini") Then
        WriteINI "SERVER0", "Name", "Server", App.Path & "\Config\Serveur.ini"
        WriteINI "SERVER0", "IP", "127.0.0.1", App.Path & "\Config\Serveur.ini"
        WriteINI "SERVER0", "Port", "4000", App.Path & "\Config\Serveur.ini"
    End If
    frmsplash.chrg.value = 60
    
    Call SetStatus("Initialisation du protocole TCP...")
    DoEvents
    
    Call TcpInit
    frmsplash.Show
    DoEvents
    Call Sleep(1)
    frmsplash.chrg.value = 80
    If Val(ReadINI("CONFIG", "jeu", App.Path & "\Config\Client.ini")) = 0 Then
        Shell (Mid$(App.Path, 1, Len(App.Path) - Len(Dir$(App.Path, vbDirectory))) & "Assistant.exe")
        Call GameDestroy
    End If
    If Val(ReadINI("CONFIG", "auto-maj", App.Path & "\Config\Client.ini")) = 1 Then Call Updater
    frmMainMenu.Show
    ConOff = False
    Call SendData("PICVALUE" & END_CHAR)
    frmsplash.Hide
    frmMirage.Timer2.Enabled = False
    frmMainMenu.Timer2.Enabled = False
    frmsplash.chrg.value = 100
    If Val(ReadINI("CONFIG", "ERR", App.Path & "\Config.ini")) <> 0 Then Call CheckErr
Exit Sub
er:
MsgBox "Erreur dans le code d'initialisation(" & Err.Number & " : " & Err.description & ")" & vbCrLf & "Merci de la rapporter sur le forum de FRoG Creator si elle persiste."
Call GameDestroy
End Sub

Sub Main2()
Dim i As Long
Dim Ending As String
    Call EcrireEtat(vbNullString)
    save = 0
    InProprieter = False
    InDefTel = False
    ScreenMode = False
    frmsplash.Show
    Call SetStatus("Vérification des dossiers...")
    DoEvents
    
    ' Check if the maps directory is there, if its not make it
    If LCase$(Dir$(App.Path & "\Classes", vbDirectory)) <> "classes" Then Call MkDir$(App.Path & "\Classes")
    If LCase$(Dir$(App.Path & "\Maps", vbDirectory)) <> "maps" Then Call MkDir$(App.Path & "\Maps")
    If UCase$(Dir$(App.Path & "\GFX", vbDirectory)) <> "GFX" Then Call MkDir$(App.Path & "\GFX")
    If UCase$(Dir$(App.Path & "\Music", vbDirectory)) <> "MUSIC" Then Call MkDir$(App.Path & "\Music")
    If UCase$(Dir$(App.Path & "\SFX", vbDirectory)) <> "SFX" Then Call MkDir$(App.Path & "\SFX")
    If UCase$(Dir$(App.Path & "\Flashs", vbDirectory)) <> "FLASHS" Then Call MkDir$(App.Path & "\Flashs")
    If UCase$(Dir$(App.Path & "\items", vbDirectory)) <> "ITEMS" Then Call MkDir$(App.Path & "\items")
    If UCase$(Dir$(App.Path & "\maps", vbDirectory)) <> "MAPS" Then Call MkDir$(App.Path & "\maps")
    If UCase$(Dir$(App.Path & "\shops", vbDirectory)) <> "SHOPS" Then Call MkDir$(App.Path & "\shops")
    If UCase$(Dir$(App.Path & "\pnjs", vbDirectory)) <> "PNJS" Then Call MkDir$(App.Path & "\pnjs")
    If UCase$(Dir$(App.Path & "\spells", vbDirectory)) <> "SPELLS" Then Call MkDir$(App.Path & "\spells")
    If UCase$(Dir$(App.Path & "\quetes", vbDirectory)) <> "QUETES" Then Call MkDir$(App.Path & "\quetes")
    If UCase$(Dir$(App.Path & "\Config", vbDirectory)) <> "CONFIG" Then Call MkDir$(App.Path & "\Config")
    frmsplash.chrg.value = 10
        
    Call SetStatus("Transfère des données...")
    If Not FileExiste("Config\account.ini") Or UCase$(Dir$(App.Path & "\Music", vbDirectory)) <> "MUSIC" Then
        Dim PathSource As String
        PathSource = Mid$(App.Path, 1, Len(App.Path) - 7)
        'Call FileCopy(PathSource & "Client\GFX\Sprites.png", App.Path & "\GFX\Sprites.png")
        Call FileCopy(PathSource & "Client\GFX\Arrows.png", App.Path & "\GFX\Arrows.png")
        'Call FileCopy(PathSource & "Client\GFX\BigSprites.png", App.Path & "\GFX\BigSprites.png")
        Call FileCopy(PathSource & "Client\GFX\Emoticons.png", App.Path & "\GFX\Emoticons.png")
        Call FileCopy(PathSource & "Client\GFX\items.png", App.Path & "\GFX\items.png")
        'Call FileCopy(PathSource & "Client\GFX\Spells.png", App.Path & "\GFX\Spells.png")
        Call FileCopy(PathSource & "Client\GFX\Tiles0.png", App.Path & "\GFX\Tiles0.png")
        Call FileCopy(PathSource & "Client\GFX\Tiles1.png", App.Path & "\GFX\Tiles1.png")
        Call FileCopy(PathSource & "Client\GFX\Tiles2.png", App.Path & "\GFX\Tiles2.png")
        Call FileCopy(PathSource & "Client\GFX\Tiles3.png", App.Path & "\GFX\Tiles3.png")
        Call FileCopy(PathSource & "Client\GFX\Tiles4.png", App.Path & "\GFX\Tiles4.png")
        Call FileCopy(PathSource & "Client\GFX\Tiles5.png", App.Path & "\GFX\Tiles5.png")
        Call FileCopy(PathSource & "Client\GFX\Tiles6.png", App.Path & "\GFX\Tiles6.png")
        Dim rep As String
        'SFX deplacement
        'obtient le premier fichier ou répertoire qui est dans "c:\"
        On Error Resume Next
        rep = Dir$(PathSource & "Client\SFX\*.*", vbDirectory)
        'boucle tant que le répertoire n'a pas été entièrement parcouru
    Do While (rep > vbNullString)
        'teste si c'est un fichier ou un répertoire
        If GetAttr(PathSource & "Client\SFX\" & rep) = vbDirectory Then
        Else
        'MsgBox "Fichier " & rep
        Call FileCopy(PathSource & "Client\SFX\" & rep, App.Path & "\SFX\" & rep)
        End If
        'passe à l'élément suivant
        rep = Dir
        Sleep 1
    Loop
      'config deplacement
        On Error Resume Next
        rep = Dir$(PathSource & "Client\Config\*.*", vbDirectory)
    Do While (rep > vbNullString)
        If GetAttr(PathSource & "Client\Config\" & rep) = vbDirectory Then
        Else
        Call FileCopy(PathSource & "Client\Config\" & rep, App.Path & "\Config\" & rep)
        End If
        rep = Dir
        Sleep 1
    Loop
       'music deplacement
        On Error Resume Next
        rep = Dir$(PathSource & "Client\Music\*.*", vbDirectory)
    Do While (rep > vbNullString)
        If GetAttr(PathSource & "Client\Music\" & rep) = vbDirectory Then
        Else
        Call FileCopy(PathSource & "Client\Music\" & rep, App.Path & "\Music\" & rep)
        End If
        rep = Dir
        Sleep 1
    Loop
    End If
    frmsplash.chrg.value = 20
    
    Dim FileName As String
    FileName = App.Path & "\Config\Account.ini"
    If FileExiste("Config\Account.ini") Then
        frmoptions.chkbubblebar.value = ReadINI("CONFIG", "SpeechBubbles", FileName)
        frmoptions.chknpcbar.value = ReadINI("CONFIG", "NpcBar", FileName)
        frmoptions.chknpcname.value = ReadINI("CONFIG", "NPCName", FileName)
        frmoptions.chkplayerbar.value = ReadINI("CONFIG", "PlayerBar", FileName)
        frmoptions.chkplayername.value = ReadINI("CONFIG", "PlayerName", FileName)
        frmoptions.chkplayerdamage.value = ReadINI("CONFIG", "NPCDamage", FileName)
        frmoptions.chknpcdamage.value = ReadINI("CONFIG", "PlayerDamage", FileName)
        frmoptions.chkmusic.value = ReadINI("CONFIG", "Music", FileName)
        frmoptions.chksound.value = ReadINI("CONFIG", "Sound", FileName)
        frmoptions.chkAutoScroll.value = ReadINI("CONFIG", "AutoScroll", FileName)
        frmoptions.chknobj.value = Val(ReadINI("CONFIG", "NomObjet", FileName))
        frmoptions.chkLowEffect.value = Val(ReadINI("CONFIG", "LowEffect", FileName))

        If ReadINI("CONFIG", "MapGrid", FileName) = 0 Then
            frmMirage.grile.Checked = False
        Else
            frmMirage.grile.Checked = True
        End If
    Else
        WriteINI "INFO", "Account", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "INFO", "Password", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CONFIG", "WebSite", "http://www.frogcreator.fr", App.Path & "\Config\Client.ini"
        WriteINI "CONFIG", "Version", "0.4", App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine1", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine2", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine3", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine4", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine5", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine6", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine7", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine8", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine9", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine10", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine11", vbNullString, App.Path & "\Config\Client.ini"
        WriteINI "CONFIG", "SpeechBubbles", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "NpcBar", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "NPCName", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "NPCDamage", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "PlayerBar", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "PlayerName", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "PlayerDamage", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "MapGrid", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "Music", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "Sound", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "AutoScroll", 1, App.Path & "\Config\Account.ini"
        WriteINI "CONFIG", "LowEffect", 0, App.Path & "\Config\Account.ini"
    End If
    If Not FileExiste("config.ini") Then
        WriteINI "INFO", "PIC_PL", 64, App.Path & "\Config.ini"
        WriteINI "INFO", "PIC_NPC1", 2, App.Path & "\Config.ini"
        WriteINI "INFO", "PIC_NPC2", 32, App.Path & "\Config.ini"
    End If
    frmsplash.chrg.value = 30
    
    Call SetStatus("Vérification du status...")
    DoEvents
    
    ' Make sure we set that we aren't in the game
    InGame = False
    GettingMap = True
    InEditor = False
    InItemsEditor = False
    InNpcEditor = False
    InShopEditor = False
    InEmoticonEditor = False
    InArrowEditor = False
    InMouvEditor = False
    InQuetesEditor = False
    
    'frmMirage.picItems.Picture = LoadPNG(App.Path & "\GFX\items.png")
    'frmSpriteChange.picSprites.Picture = LoadPNG(App.Path & "\GFX\sprites.png") 'a faire
    'frmclasseseditor.picSprites.Picture = LoadPNG(App.Path & "\GFX\sprites.png")
    frmsplash.chrg.value = 40
    
    If FileExiste("Config\Serveur.ini") = False Then
        WriteINI "SERVER0", "Name", "Server 0", App.Path & "\Config\Serveur.ini"
        WriteINI "SERVER0", "IP", "127.0.0.1", App.Path & "\Config\Serveur.ini"
        WriteINI "SERVER0", "Port", "4000", App.Path & "\Config\Serveur.ini"
    End If
    
    Call SetStatus("Initialisation des mises à jours...")
    If FileExiste("Config\Updater.ini") = False Then
    WriteINI "UPDATER", "WebSite", "http://roonline.free.fr/patch/", App.Path & "\Config\Updater.ini"
    WriteINI "UPDATER", "WebNews", "http://roonline.free.fr/patch/patch.html", App.Path & "\Config\Updater.ini"
    WriteINI "VERSION", "Version", "0.1", App.Path & "\Config\info.ini"
    End If
    
    frmsplash.chrg.value = 60
    
    Call SetStatus("Initialisation du protocole TCP...")
    DoEvents
    
    
    
    Call TcpInit
    frmsplash.Show
    Call Sleep(1)
    frmsplash.chrg.value = 80
    
    If Val(ReadINI("CONFIG", "jeu", App.Path & "\Config\Client.ini")) = 0 Then
        Shell (Mid$(App.Path, 1, Len(App.Path) - Len(Dir$(App.Path, vbDirectory))) & "Assistant.exe")
        Call GameDestroy
    Else
        Update = False
        Call Updater
        If Update = False Then frmMainMenu.Show
    End If
    
    ConOff = False
    Call SendData("PICVALUE" & END_CHAR)
    frmsplash.chrg.value = 100
    frmsplash.Visible = False
    frmMirage.Timer2.Enabled = False
    frmMainMenu.Timer2.Enabled = False
    
End Sub

Sub SetStatus(ByVal Caption As String)
    frmsplash.lblStatus.Caption = Caption
    Call EcrireEtat(Caption)
End Sub

Sub MenuState(ByVal State As Long)
    If frmMainMenu.Check2.value = Unchecked Then
        Connucted = True
        frmsplash.Visible = True
        DoEvents
        frmsplash.chrg.value = 50
        Call SetStatus("Connection au Serveur...")
    End If
    
    Select Case State
        Case MENU_STATE_LOGIN
            frmMainMenu.Visible = False
            If frmMainMenu.Check2.value = Checked Then HORS_LIGNE = 1: Call Horsligne: Exit Sub
            If ConnectToServer Then
                HORS_LIGNE = 0
                Call SetStatus("Connecté, Envoie de la connexion au compte..")
                frmsplash.chrg.value = 80
                Call SendLogin(frmMainMenu.txtName.Text, frmMainMenu.txtPassword.Text)
            End If
            
        Case MENU_STATE_USECHAR
            frmChars.Visible = False
            If ConnectToServer Then
                Call StopMidi
                frmsplash.chrg.value = 80
                Call SetStatus("Patience...")
                Call SendUseChar(frmChars.lstChars.ListIndex + 1)
            End If
    End Select

    If Not IsConnected And Connucted Then
        frmMainMenu.Visible = True
        frmsplash.Visible = False
        Call MsgBox("Désolé, le serveur semble être indisponible, réessayer dans quelque minute ou visiter" & WEBSITE, vbOKOnly, GAME_NAME)
    End If
End Sub
Sub GameInit()
Dim i As Long
    Call StopMidi
    frmMirage.Visible = True
    Call SendData("mapreport" & END_CHAR)
    frmsplash.Visible = False
    Call InitDirectX
    Call SendRequestEditMap
    Call ChargerObjets(MyIndex)
    Call ChargerFleche
    Call ChargerEmots
    Call ChargerPnjs
    Call ChargerMagasins
    Call ChargerSorts
    Call ChargerQuetes
    Call ChargerRecette
    If ExtraSheets < frmMirage.Tiles.Count - 1 Then
        For i = ExtraSheets To 5
            Unload frmMirage.Tiles(i)
            Call frmMirage.tilescmb.RemoveItem(i)
        Next i
    Else
        For i = 0 To ExtraSheets
            If i > frmMirage.Tiles.Count - 1 Then Load frmMirage.Tiles(i): frmMirage.Tiles(i).Caption = "Tiles" & i: frmMirage.Tiles(i).Checked = False
            If i > frmMirage.tilescmb.ListCount - 1 Then Call frmMirage.tilescmb.AddItem("Tiles" & i, i)
        Next i
    End If
    Accepter = False
End Sub

Sub GameLoop()
Dim Tick As Long
Dim TickFPS As Byte
Dim FPS As Long
Dim TickMove As Long
Dim x As Long
Dim y As Long
Dim i As Long
Dim rec_back As RECT
Dim Coulor As Long
Dim screen_xg(1 To 9) As Integer
Dim screen_xd(1 To 9) As Integer
Dim screen_yh(1 To 9) As Integer
Dim screen_yb(1 To 9) As Integer
Dim MaxDrawMapX As Long 'Calcul du maximum a dessiner en X
Dim MinDrawMapX As Long 'Calcul du minimum a dessiner en X
Dim MaxDrawMapY As Long 'Calcul du maximum a dessiner en Y
Dim MinDrawMapY As Long 'Calcul du minimum a dessiner en Y
Dim XT As Long
Dim YT As Long
On Error GoTo er:
    If Not InGame Then Exit Sub
    
    ' Set the focus
    frmMirage.picScreen.SetFocus
    
    ' Set font
    Call SetFont("Fixedsys", 18)
                
    ' Used for calculating fps
    TickFPS = 0
    TickMove = 0
            
    'Initialisation du RECT pour le backbuffer
    rec_back.Top = 0
    rec_back.Bottom = (MAX_MAPY + 1) * PIC_Y
    rec_back.Left = 0
    rec_back.Right = (MAX_MAPX + 1) * PIC_X
    
    'Initialisation des variables pour les limites de la "vue" du joueur
    For i = 3 To 9 Step 3
        screen_xg(i) = ((frmMirage.picScreen.Width * i / 3) / 64) - 1
        screen_xd(i) = ((frmMirage.picScreen.Width * i / 3) / 32) - screen_xg(i) - 1
        screen_yh(i) = ((frmMirage.picScreen.Height * i / 3) / 64) - 1
        screen_yb(i) = ((frmMirage.picScreen.Height * i / 3) / 32) - screen_yh(i) - 1
    Next i
    
    Do While InGame
        Tick = GetTickCount
        
        ' Check to make sure they aren't trying to auto do anything
        If GetAsyncKeyState(VK_UP) >= 0 And DirUp = True Then DirUp = False
        If GetAsyncKeyState(VK_DOWN) >= 0 And DirDown = True Then DirDown = False
        If GetAsyncKeyState(VK_LEFT) >= 0 And DirLeft = True Then DirLeft = False
        If GetAsyncKeyState(VK_RIGHT) >= 0 And DirRight = True Then DirRight = False
        If GetAsyncKeyState(VK_CONTROL) >= 0 And ControlDown = True Then ControlDown = False
        If GetAsyncKeyState(VK_SHIFT) >= 0 And ShiftDown = True Then ShiftDown = False
        
        If frmcraft.Visible Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
            ControlDown = False
            ShiftDown = False
        End If
        ' Check to make sure we are still connected
        If Not IsConnected And HORS_LIGNE = 0 Then InGame = False
        
        'Effacer le BackBuffer avant de dessiner dessus
        Call DD_BackBuffer.BltColorFill(rec_back, 0)
        
        ' Check if we need to restore surfaces
        If NeedToRestoreSurfaces Then
rest:
            Do While NeedToRestoreSurfaces
                DoEvents
                Sleep 1
            Loop
            DD.RestoreAllSurfaces: Call InitSurfaces
        End If
                
        If Not GettingMap Then
            sx = 32
            sy = 32
        
            If MAX_MAPX < screen_xg(VZoom) + screen_xd(VZoom) + 1 Then
                NewX = Player(MyIndex).x * PIC_X + Player(MyIndex).XOffset
                NewXOffset = 0
                NewPlayerX = 0
                sx = 0
            ElseIf Player(MyIndex).x <= screen_xg(VZoom) Then
                NewPlayerX = 0
                If Player(MyIndex).x = screen_xg(VZoom) And Player(MyIndex).Dir = DIR_LEFT Then
                    NewX = screen_xg(VZoom) * PIC_X
                    NewXOffset = Player(MyIndex).XOffset
                Else
                    NewX = Player(MyIndex).x * PIC_X + Player(MyIndex).XOffset
                    NewXOffset = 0
                End If
            ElseIf MAX_MAPX - Player(MyIndex).x <= screen_xd(VZoom) Then
                NewPlayerX = MAX_MAPX - screen_xd(VZoom) - screen_xg(VZoom)
                If MAX_MAPX - Player(MyIndex).x = screen_xd(VZoom) And Player(MyIndex).Dir = DIR_RIGHT Then
                    NewX = screen_xg(VZoom) * PIC_X
                    NewXOffset = Player(MyIndex).XOffset
                Else
                    NewX = (Player(MyIndex).x - MAX_MAPX + screen_xd(VZoom) + screen_xg(VZoom)) * PIC_X + Player(MyIndex).XOffset
                    NewXOffset = 0
                End If
            Else
                NewPlayerX = Player(MyIndex).x - screen_xg(VZoom)
                NewX = screen_xg(VZoom) * PIC_X
                NewXOffset = Player(MyIndex).XOffset
            End If
            
            If MAX_MAPY < screen_yh(VZoom) + screen_yb(VZoom) + 1 Then
                NewY = Player(MyIndex).y * PIC_Y + Player(MyIndex).YOffset
                NewYOffset = 0
                NewPlayerY = 0
                sy = 0
            ElseIf Player(MyIndex).y <= screen_yh(VZoom) Then
                NewPlayerY = 0
                If Player(MyIndex).y = screen_yh(VZoom) And Player(MyIndex).Dir = DIR_UP Then
                    NewY = screen_yh(VZoom) * PIC_Y
                    NewYOffset = Player(MyIndex).YOffset
                Else
                    NewY = Player(MyIndex).y * PIC_Y + Player(MyIndex).YOffset
                    NewYOffset = 0
                End If
            ElseIf MAX_MAPY - Player(MyIndex).y <= screen_yb(VZoom) Then
                NewPlayerY = MAX_MAPY - screen_yb(VZoom) - screen_yh(VZoom)
                If MAX_MAPY - Player(MyIndex).y = screen_yb(VZoom) And Player(MyIndex).Dir = DIR_DOWN Then
                    NewY = screen_yh(VZoom) * PIC_Y
                    NewYOffset = Player(MyIndex).YOffset
                Else
                    NewY = (Player(MyIndex).y - MAX_MAPY + screen_yb(VZoom) + screen_yh(VZoom)) * PIC_Y + Player(MyIndex).YOffset
                    NewYOffset = 0
                End If
            Else
                NewPlayerY = Player(MyIndex).y - screen_yh(VZoom)
                NewY = screen_yh(VZoom) * PIC_Y
                NewYOffset = Player(MyIndex).YOffset
            End If
            
            'Calcul des variables de scrolling restante
            NewPlayerPicX = NewPlayerX * PIC_X
            NewPlayerPicY = NewPlayerY * PIC_Y
            NewPlayerPOffsetX = NewPlayerPicX + NewXOffset
            NewPlayerPOffsetY = NewPlayerPicY + NewYOffset
            
            MaxDrawMapX = NewPlayerX + screen_xg(VZoom) + screen_xd(VZoom) + 1
            MinDrawMapX = NewPlayerX - 2
            MaxDrawMapY = NewPlayerY + screen_yh(VZoom) + screen_yb(VZoom) + 1
            MinDrawMapY = NewPlayerY - 2
            If MaxDrawMapX > MAX_MAPX Then MaxDrawMapX = MAX_MAPX
            If MaxDrawMapY > MAX_MAPY Then MaxDrawMapY = MAX_MAPY
            If MinDrawMapX < 0 Then MinDrawMapX = 0
            If MinDrawMapY < 0 Then MinDrawMapY = 0

            ' Blit out tiles layers ground/anim1/anim2 'la pour map zoom
            For y = MinDrawMapY To MaxDrawMapY
                For x = MinDrawMapX To MaxDrawMapX
                    Call BltTile(x, y)
                Next x
            Next y
            
            If Not ScreenMode Then
                ' Blit out the items
                For i = 1 To MAX_MAP_ITEMS
                    If MapItem(i).num > 0 Then Call BltItem(i)
                Next i
                                    
'                If AccOpt.PlayBar Then
                    ' Blit players bar
                    'For i = 1 To MAX_PLAYERS
                        'If IsPlaying(i) And GetPlayerMap(i) = Player(MyIndex).Map Then
                            'Call BltPlayerBars(i)
                        'End If
                    'Next i
 '               End If

                ' Blit out the sprite change attribute
                For y = MinDrawMapY To MaxDrawMapY
                    For x = MinDrawMapX To MaxDrawMapX
                        If Map(Player(MyIndex).Map).tile(x, y).Type = TILE_TYPE_SPRITE_CHANGE Then
                            Call BltSpriteChange(x, y)
                            'Dessin du haut des atributs sprite change en 32*64
                            If PIC_PL > 1 Then Call BltSpriteChange2(x, y)
                        End If
                    Next x
                Next y
                
                ' Blit out the npcs
               For i = 1 To MAX_MAP_NPCS
                   If MapNpc(i).num > 0 And MapNpc(i).num < MAX_NPCS Then
                       If CLng(Npc(MapNpc(i).num).vol) = 0 Then
                           Call BltNpc(i)
                           If AccOpt.NpcBar Then Call BltNpcBars(i)
                       End If
                   End If
               Next i
               
                If Not InEditor Then
                    ' Blit out players ,arrows and spells
                    For i = 1 To MAX_PLAYERS
                        If IsPlaying(i) And Player(i).Map = Player(MyIndex).Map Then
                            If Map(Player(MyIndex).Map).guildSoloView = 1 Then
                                If Player(MyIndex).guild = Player(i).guild Then
                                    Call BltPlayerOmbre(i)
                                    Call BltPlayer(i)
                                    Call BltArrow(i)
                                    If Player(i).PetSlot <> 0 Then Call BltPlayerPet(i)
                                End If
                            Else
                                Call BltPlayerOmbre(i)
                                Call BltPlayer(i)
                                Call BltArrow(i)
                                If Player(i).PetSlot <> 0 Then Call BltPlayerPet(i)
                            End If
                        End If
                    Next i
                End If
                
                If AccOpt.PlayBar And Not InEditor Then Call BltPlayerBar  's(i)
                
                ' Dessiner le haut des npc apres le bas des joueurs
                For i = 1 To MAX_MAP_NPCS
                    If MapNpc(i).num > 0 And MapNpc(i).num < MAX_NPCS Then
                        If CLng(Npc(MapNpc(i).num).vol) = 0 Then
                            If PIC_PL > 1 Then Call BltNpcTop(i)
                        End If
                    End If
                Next i
                    
                If Not InEditor Then
                    For i = 1 To MAX_PLAYERS
                        If IsPlaying(i) And Player(i).Map = Player(MyIndex).Map Then
                            If PIC_PL > 1 Then
                                If Map(Player(MyIndex).Map).guildSoloView = 1 Then
                                    If Player(MyIndex).guild = Player(i).guild Then
                                        Call BltPlayerTop(i)
                                    End If
                                Else
                                    Call BltPlayerTop(i)
                                End If
                            End If

                            Call BltBlood(i, PIC_X, PIC_Y, 40)
                            ' Call BltBlood(i) ferais aussi l'affaire car les autres paramètres peuvent être modifier selon le blood.png.
                            ' Le premier et le second paramètre sont la taille X et Y ce qui permet d'avoir des animations de sang 96X96 exemple.
                            ' Il se peux que le code demande à être modifié dans cette condition.
                            ' Le dernier paramètre est le temps de chaque image en ms (1000 ms = 1 seconde).
                        
                            Call BltSpell(i)
                        
                            If Player(i).LevelUpT + 3000 > Tick Then Call BltPlayerLevelUp(i) Else Player(i).LevelUpT = 0
                            Call BltEmoticons(i)
                            
                            Call BltPlayerAnim(i)
                        End If
                    Next i
                    'Dessiner le joueur locale
                    'If IsPlaying(MyIndex) Then
                    '    If PIC_PL > 1 Then Call BltPlayerTop(MyIndex): Call BltEmoticons(MyIndex)                   '<- pour 32*64
                    '    Call BltPlayer(MyIndex)
                    '    Call BltSpell(MyIndex)
                    '    If Player(MyIndex).LevelUpT + 3000 > Tick Then Call BltPlayerLevelUp(MyIndex) Else Player(MyIndex).LevelUpT = 0
                    'End If
                    
                End If
            End If
        End If
        
        If Not GettingMap And (Not ScreenMode Or ScreenDC) And Not InEditor And AccOpt.PlayName And Not ScreenMode Then
            'Verouiller le backbuffer pour pouvoir ecrire le nom des joueurs et de leur guildes
            TexthDC = DD_BackBuffer.GetDC
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And Player(i).Map = Player(MyIndex).Map Then
                    If Map(Player(MyIndex).Map).guildSoloView = 1 Then
                        If Player(MyIndex).guild = Player(i).guild Then
                            Call BltPlayerGuildName(i)
                            Call BltPlayerName(i)
                        End If
                    Else
                        Call BltPlayerGuildName(i)
                        Call BltPlayerName(i)
                    End If
                End If
            Next i
            Call DD_BackBuffer.ReleaseDC(TexthDC)
        End If
                
        ' Blit out tile layer fringe
        For y = MinDrawMapY To MaxDrawMapY
            For x = MinDrawMapX To MaxDrawMapX
                Call BltFringeTile(x, y)
            Next x
        Next y
                    
        If Not GettingMap And Not ScreenMode Then
            'Dessiner les PNJs volant
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).num > 0 And MapNpc(i).num < MAX_NPCS Then
                    If CLng(Npc(MapNpc(i).num).vol) <> 0 Then
                        Call BltNpc(i)
                        If AccOpt.NpcBar Then Call BltNpcBars(i)
                        If PIC_PL > 1 Then Call BltNpcTop(i)
                    End If
                End If
            Next i
        End If
        
        If AccOpt.CPreVisu And InEditor And frmMirage.tp(1).Checked And frmMirage.MousePointer <> 99 And frmMirage.MousePointer <> 2 Then Call BltVisu
        
        If Not GettingMap Then
            If Map(Player(MyIndex).Map).Indoors = 0 Then Call BltWeather
        End If

        If InEditor And AccOpt.MapGrid And Not ScreenDC And TileFile(ExtraSheets) Then
            For y = MinDrawMapY To MaxDrawMapY
                For x = MinDrawMapX To MaxDrawMapX
                    Call BltTile2(x * 32, y * 32, 0)
                Next x
            Next y
        End If
            
        ' Lock the backbuffer so we can draw text and names
        TexthDC = DD_BackBuffer.GetDC
        If Not GettingMap Then
            If Not ScreenMode Or ScreenDC Then
                If AccOpt.NpcDamage And Not ScreenMode Then
                    If Not AccOpt.PlayName Then
                        If Tick < NPCDmgTime + 2000 Then Call DrawText(TexthDC, ((Len(NPCDmgDamage)) \ 2) * 3 + NewX + sx, NewY - 22 - ii + sy, NPCDmgDamage, QBColor(BrightRed))
                    Else
                        If GetPlayerGuild(MyIndex) <> vbNullString Then
                            If Tick < NPCDmgTime + 2000 Then Call DrawText(TexthDC, ((Len(NPCDmgDamage)) \ 2) * 3 + NewX + sx, NewY - 42 - ii + sy, NPCDmgDamage, QBColor(BrightRed))
                        Else
                            If Tick < NPCDmgTime + 2000 Then Call DrawText(TexthDC, ((Len(NPCDmgDamage)) \ 2) * 3 + NewX + sx, NewY - 22 - ii + sy, NPCDmgDamage, QBColor(BrightRed))
                        End If
                    End If
                    ii = ii + 1
                End If
                
                If AccOpt.PlayDamage And Not ScreenMode Then
                    If NPCWho > 0 Then
                        If MapNpc(NPCWho).num > 0 Then
                            If Not AccOpt.NpcName Then
                                    If Tick < DmgTime + 2000 Then Call DrawText(TexthDC, (MapNpc(NPCWho).x - NewPlayerX) * PIC_X + sx + ((Len(DmgDamage)) \ 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sy - 20 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(White))
                            Else
                                    If Tick < DmgTime + 2000 Then Call DrawText(TexthDC, (MapNpc(NPCWho).x - NewPlayerX) * PIC_X + sx + ((Len(DmgDamage)) \ 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sy - 30 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(White))
                            End If
                            iii = iii + 1
                        End If
                    End If
                End If
         
                ' speech bubble stuffs
                If AccOpt.SpeechBubbles And Not ScreenMode Then
                    For i = 1 To MAX_PLAYERS
                        If IsPlaying(i) And Player(i).Map = Player(MyIndex).Map Then
                            If Bubble(i).Text <> vbNullString And Not InEditor Then Call BltPlayerText(i)
                            If Tick > Bubble(i).Created + DISPLAY_BUBBLE_TIME Then Bubble(i).Text = vbNullString
                        End If
                    Next i
                End If
        
                'Draw NPC Names
                If AccOpt.NpcName And Not ScreenMode Then
                    For i = LBound(MapNpc) To UBound(MapNpc)
                        If MapNpc(i).num > 0 Then Call BltMapNPCName(i)
                    Next i
                End If
                
                ' Blit out attribs if in editor
                If InEditor And Not ScreenDC Then
                    For y = MinDrawMapY To MaxDrawMapY
                        For x = MinDrawMapX To MaxDrawMapX
                            With Map(Player(MyIndex).Map).tile(x, y)
                                XT = x * PIC_X + sx + 8 - NewPlayerPOffsetX
                                YT = y * PIC_Y + sy + 8 - NewPlayerPOffsetY
                                Select Case .Type
                                    Case TILE_TYPE_BLOCKED: Call DrawText(TexthDC, XT, YT, "B", QBColor(BrightRed))
                                    Case TILE_TYPE_WARP: Call DrawText(TexthDC, XT, YT, "T", QBColor(BrightBlue))
                                    Case TILE_TYPE_ITEM: Call DrawText(TexthDC, XT, YT, "O", QBColor(White))
                                    Case TILE_TYPE_NPCAVOID: Call DrawText(TexthDC, XT, YT, "BP", QBColor(White))
                                    Case TILE_TYPE_KEY: Call DrawText(TexthDC, XT, YT, "C", QBColor(White))
                                    Case TILE_TYPE_KEYOPEN: Call DrawText(TexthDC, XT, YT, "OC", QBColor(White))
                                    Case TILE_TYPE_HEAL: Call DrawText(TexthDC, XT, YT, "G", QBColor(BrightGreen))
                                    Case TILE_TYPE_KILL: Call DrawText(TexthDC, XT, YT, "M", QBColor(BrightRed))
                                    Case TILE_TYPE_SHOP: Call DrawText(TexthDC, XT, YT, "MA", QBColor(Yellow))
                                    Case TILE_TYPE_CBLOCK: Call DrawText(TexthDC, XT, YT, "CB", QBColor(Black))
                                    Case TILE_TYPE_ARENA: Call DrawText(TexthDC, XT, YT, "A", QBColor(BrightGreen))
                                    Case TILE_TYPE_SOUND: Call DrawText(TexthDC, XT, YT, "S", QBColor(Yellow))
                                    Case TILE_TYPE_SPRITE_CHANGE: Call DrawText(TexthDC, XT, YT, "CS", QBColor(Grey))
                                    Case TILE_TYPE_SIGN: Call DrawText(TexthDC, XT, YT, "PN", QBColor(Yellow))
                                    Case TILE_TYPE_DOOR: Call DrawText(TexthDC, XT, YT, "P", QBColor(Black))
                                    Case TILE_TYPE_NOTICE: Call DrawText(TexthDC, XT, YT, "A", QBColor(BrightGreen))
                                    'Case TILE_TYPE_CHEST : Call DrawText(TexthDC, xt, yt, "C", QBColor(Brown))
                                    Case TILE_TYPE_CLASS_CHANGE: Call DrawText(TexthDC, XT, YT, "CC", QBColor(White))
                                    Case TILE_TYPE_SCRIPTED: Call DrawText(TexthDC, XT, YT, "SC", QBColor(Yellow))
                                    Case TILE_TYPE_BANK: Call DrawText(TexthDC, XT, YT, "BA", QBColor(Yellow))
                                    Case TILE_TYPE_COFFRE: Call DrawText(TexthDC, XT, YT, "CO", QBColor(Yellow))
                                    Case TILE_TYPE_PORTE_CODE: Call DrawText(TexthDC, XT, YT, "PC", QBColor(Black))
                                    Case TILE_TYPE_BLOCK_MONTURE: Call DrawText(TexthDC, XT, YT, "BM", QBColor(Red))
                                    Case TILE_TYPE_BLOCK_NIVEAUX: Call DrawText(TexthDC, XT, YT, "BN", QBColor(Red))
                                    Case TILE_TYPE_TOIT: Call DrawText(TexthDC, XT, YT, "TO", QBColor(White))
                                    Case TILE_TYPE_BLOCK_GUILDE: Call DrawText(TexthDC, XT, YT, "BG", QBColor(Red))
                                    Case TILE_TYPE_BLOCK_TOIT: Call DrawText(TexthDC, XT, YT, "BT", QBColor(BrightRed))
                                    Case TILE_TYPE_BLOCK_DIR: Call DrawText(TexthDC, XT, YT, "BD", QBColor(BrightRed))
                                    Case TILE_TYPE_CRAFT: Call DrawText(TexthDC, XT, YT, "TC", QBColor(Yellow))
                                    Case TILE_TYPE_METIER: Call DrawText(TexthDC, XT, YT, "M", QBColor(Yellow))
                                End Select
                                If .Light > 0 Then Call DrawText(TexthDC, x * PIC_X + sx + 18 - NewPlayerPOffsetX, y * PIC_Y + sy + 14 - NewPlayerPOffsetY, "L", QBColor(Yellow))
                            End With
                        Next x
                    Next y
                End If
                    
                ' Draw map name
                If Not ScreenDC Then
                    If Map(Player(MyIndex).Map).Moral = MAP_MORAL_NONE Then
                    ' Int((5) * PIC_X / 2) - (Len(Trim$(Map(GetPlayerMap(MyIndex)).name))) + sx
                        Call DrawText(TexthDC, (frmMirage.picScreen.Width / 2) - (Len(Trim$(Map(Player(MyIndex).Map).name)) / 2), 5 + sx, Trim$(Map(Player(MyIndex).Map).name), QBColor(BrightRed))
                    ElseIf Map(Player(MyIndex).Map).Moral = MAP_MORAL_SAFE Then
                        Call DrawText(TexthDC, (frmMirage.picScreen.Width / 2) - (Len(Trim$(Map(Player(MyIndex).Map).name)) / 2), 5 + sx, Trim$(Map(Player(MyIndex).Map).name), QBColor(White))
                    ElseIf Map(Player(MyIndex).Map).Moral = MAP_MORAL_NO_PENALTY Then
                        Call DrawText(TexthDC, (frmMirage.picScreen.Width / 2) - (Len(Trim$(Map(Player(MyIndex).Map).name)) / 2), 5 + sx, Trim$(Map(Player(MyIndex).Map).name), QBColor(Black))
                    End If
                
                    For i = 1 To MAX_BLT_LINE
                        If BattlePMsg(i).Index > 0 Then
                            If BattlePMsg(i).Color > 15 Then Coulor = BattlePMsg(i).Color Else Coulor = QBColor(BattlePMsg(i).Color)
                            If BattlePMsg(i).Time + 60000 > Tick Then Call DrawText(TexthDC, 1 + sx, BattlePMsg(i).y + PicScHeight - 80 + sx, Trim$(BattlePMsg(i).Msg), Coulor) Else BattlePMsg(i).Done = 0
                        End If
                        
                        If BattleMMsg(i).Index > 0 Then
                            If BattleMMsg(i).Color > 15 Then Coulor = BattleMMsg(i).Color Else Coulor = QBColor(BattleMMsg(i).Color)
                            If BattleMMsg(i).Time + 60000 > Tick Then Call DrawText(TexthDC, (PicScWidth - (Len(BattleMMsg(i).Msg) * 8)) + sx, BattleMMsg(i).y + PicScHeight - 80 + sx, Trim$(BattleMMsg(i).Msg), Coulor) Else BattleMMsg(i).Done = 0
                        End If
                    Next i
                End If
            End If
            
            'Dessin de la nuit en "low effect"
            If (GameTime = TIME_NIGHT Or (frmMirage.nuitjour.Checked And InEditor)) And AccOpt.LowEffect And Map(Player(MyIndex).Map).Indoors = 0 Then Call Night(MinDrawMapX, MaxDrawMapX, MinDrawMapY, MaxDrawMapY)
        End If
    
        ' Check if we are getting a map, and if we are tell them so
        If GettingMap And Not frmmsg.Visible Then frmmsg.Show
        
        ' Release DC
        Call DD_BackBuffer.ReleaseDC(TexthDC)
        
        'Dessin du brouillard
'        Stop
        If Map(Player(MyIndex).Map).Fog <> 0 And Not AccOpt.LowEffect And (GameTime <> TIME_NIGHT And (Not frmMirage.nuitjour.Checked Or Not InEditor)) Then Call BltFog(MinDrawMapX, MaxDrawMapX, MinDrawMapY, MaxDrawMapY)
        
        'Dessin de la nuit en "hight"
        If (GameTime = TIME_NIGHT Or (frmMirage.nuitjour.Checked And InEditor)) And Not AccOpt.LowEffect And Map(Player(MyIndex).Map).Indoors = 0 Then Call Night(MinDrawMapX, MaxDrawMapX, MinDrawMapY, MaxDrawMapY)
        
        'Capture d'une carte
        If ScreenDC Then Call CarteCapture: ScreenDC = False
            
        ' Get the rect to blit to
        Call Dx.GetWindowRect(frmMirage.picScreen.hwnd, rec_pos)
        rec_pos.Bottom = rec_pos.Top - sy + ((MAX_MAPY + 1) * PIC_Y) / VZoom * 3
        rec_pos.Right = rec_pos.Left - sx + ((MAX_MAPX + 1) * PIC_X) / VZoom * 3
        rec_pos.Top = rec_pos.Bottom - ((MAX_MAPY + 1) * PIC_Y) / VZoom * 3
        rec_pos.Left = rec_pos.Right - ((MAX_MAPX + 1) * PIC_X) / VZoom * 3
        
        ' Blit the backbuffer
        Call DD_PrimarySurf.Blt(rec_pos, DD_BackBuffer, rec_back, DDBLT_WAIT)
            
        If TickMove < Tick And Not GettingMap Then
            ' Check if player is trying to move
            Call CheckMovement
            
            ' Check to see if player is trying to attack
            Call CheckAttack
            
            ' Process player movements (actually move them)
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then Call ProcessMovement(i)
            Next i
        
               ' Process npc movements (actually move them)
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).num > 0 Then Call ProcessNpcMovement(i)
            Next i
            
            ' Change map animation every 250 milliseconds
            If Tick > MapAnimTimer + 250 Then
                If Not MapAnim Then MapAnim = True Else MapAnim = False
                MapAnimTimer = Tick
            End If
                
            Call MakeMidiLoop
            
            'Verifier si il faut sauvegarder
            If SauvAuto = 0 Then SauvAuto = Tick
            If Tick >= SauvAuto + 60000 Then SauvAuto = 0: Call SauveAuto
            
            TickMove = Tick + 30
            
            'Calcul des FPS
            TickFPS = TickFPS + 1
            If TickFPS >= 33 Then TickFPS = 0: GameFPS = FPS: FPS = 0
        End If
        
        'Bloquer les FPS a 30 pour éviter de surcharger le processeur
        Do While GetTickCount < Tick + 30
            DoEvents
            Sleep 1
        Loop
        
        DoEvents
        Sleep 2
        FPS = FPS + 1
    Loop
               
    frmMirage.Visible = False
    frmsplash.Visible = True
    DoEvents
    frmsplash.chrg.value = 80
    Call SetStatus("Destroying game data...")
            
    ' Shutdown the game
    Call MsgBox("Erreur relancez SVP")
    Call GameDestroy
    
    ' Report disconnection if server disconnects
    If IsConnected = False And HORS_LIGNE = 0 Then Call MsgBox("Merci d'avoirs joué à " & GAME_NAME & "!", vbOKOnly, GAME_NAME)
Exit Sub
er:
If Val(Mid$(Err.Number, 1, 9)) = -200553208 Then GoTo rest:
MsgBox "Erreur dans le code de boucle(" & Err.Number & " : " & Err.description & ")" & vbCrLf & "Merci de la rapporter sur le forum de FRoG Creator si elle persiste."
Call EcrireEtat("Une erreur de boucle (Numéros de l'erreur : " & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & ").")
Call GameDestroy
End Sub
Sub GameDestroy()
    On Error Resume Next
    Call TcpDestroy
    Call DestroyDirectX
    Call StopMidi
    WriteINI "CONFIG", "ERR", 0, App.Path & "\Config.ini"
    If FileExiste("r.exe") Then Call Shell("r.exe")
    End
End Sub

Sub BltTile(ByVal x As Long, ByVal y As Long)
Dim Ground As Long
Dim Anim1 As Long
Dim Anim2 As Long
Dim Mask2 As Long
Dim M2Anim As Long
Dim Mask3 As Long '<--
Dim M3Anim As Long '<--
Dim GroundTileSet As Byte
Dim MaskTileSet As Byte
Dim AnimTileSet As Byte
Dim Mask2TileSet As Byte
Dim M2AnimTileSet As Byte
Dim Mask3TileSet As Byte '<--
Dim M3AnimTileSet As Byte '<--
Dim tx As Long
Dim ty As Long
    Ground = Map(Player(MyIndex).Map).tile(x, y).Ground
    Anim1 = Map(Player(MyIndex).Map).tile(x, y).Mask
    Anim2 = Map(Player(MyIndex).Map).tile(x, y).Anim
    Mask2 = Map(Player(MyIndex).Map).tile(x, y).Mask2
    M2Anim = Map(Player(MyIndex).Map).tile(x, y).M2Anim
    Mask3 = Map(Player(MyIndex).Map).tile(x, y).Mask3 '<--
    M3Anim = Map(Player(MyIndex).Map).tile(x, y).M3Anim '<--
    
    GroundTileSet = Map(Player(MyIndex).Map).tile(x, y).GroundSet
    MaskTileSet = Map(Player(MyIndex).Map).tile(x, y).MaskSet
    AnimTileSet = Map(Player(MyIndex).Map).tile(x, y).AnimSet
    Mask2TileSet = Map(Player(MyIndex).Map).tile(x, y).Mask2Set
    M2AnimTileSet = Map(Player(MyIndex).Map).tile(x, y).M2AnimSet
    Mask3TileSet = Map(Player(MyIndex).Map).tile(x, y).Mask3Set '<--
    M3AnimTileSet = Map(Player(MyIndex).Map).tile(x, y).M3AnimSet '<--
   
    If GroundTileSet > ExtraSheets Then Exit Sub
    If Not TileFile(GroundTileSet) Then Exit Sub
    ty = (y - NewPlayerY) * PIC_Y + sy - NewYOffset
    tx = (x - NewPlayerX) * PIC_X + sx - NewXOffset
    
    rec.Top = (Ground \ TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Ground - (Ground \ TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
    Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(GroundTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    If ((Not MapAnim) Or (Anim2 <= 0)) Then
        ' Is there an animation tile to plot?
        If Anim1 > 0 And TempTile(x, y).DoorOpen = NO And MaskTileSet <= ExtraSheets Then
            If Not TileFile(MaskTileSet) Then Exit Sub
            rec.Top = (Anim1 \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Anim1 - (Anim1 \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(MaskTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        ' Is there a second animation tile to plot?
        If Anim2 > 0 And AnimTileSet <= ExtraSheets Then
            If Not TileFile(AnimTileSet) Then Exit Sub
            rec.Top = (Anim2 \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Anim2 - (Anim2 \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If (Not MapAnim) Or (M2Anim <= 0) Then
        ' Is there an animation tile to plot?
        If Mask2 > 0 And Mask2TileSet <= ExtraSheets Then
            If Not TileFile(Mask2TileSet) Then Exit Sub
            rec.Top = (Mask2 \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Mask2 - (Mask2 \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(Mask2TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        ' Is there a second animation tile to plot?
        If M2Anim > 0 And M2AnimTileSet <= ExtraSheets Then
            If Not TileFile(M2AnimTileSet) Then Exit Sub
            rec.Top = (M2Anim \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (M2Anim - (M2Anim \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(M2AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If (Not MapAnim) Or (M3Anim <= 0) Then   '<--
        ' Is there an animation tile to plot?
        If Mask3 > 0 And Mask3TileSet <= ExtraSheets Then
            If Not TileFile(Mask3TileSet) Then Exit Sub
            rec.Top = (Mask3 \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Mask3 - (Mask3 \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(Mask3TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        ' Is there a second animation tile to plot?
        If M3Anim > 0 And M3AnimTileSet <= ExtraSheets Then
            If Not TileFile(M3AnimTileSet) Then Exit Sub
            rec.Top = (M3Anim \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (M3Anim - (M3Anim \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(M3AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If  '<--
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = (y - NewPlayerY) * PIC_Y + sy - NewYOffset
        .Bottom = .Top + PIC_Y
        .Left = (x - NewPlayerX) * PIC_X + sx - NewXOffset
        .Right = .Left + PIC_X
    End With
    'Affichage du panorama inférieur si il y en à un
    If Trim$(Map(Player(MyIndex).Map).PanoInf) <> vbNullString Then
        rec.Top = y * PIC_Y
        If rec.Top + PIC_Y > DDSD_PanoInf.lHeight Then rec.Bottom = DDSD_PanoInf.lHeight: rec_pos.Bottom = rec_pos.Bottom - ((rec.Top + PIC_Y) - DDSD_PanoInf.lHeight) Else rec.Bottom = rec.Top + PIC_Y
        rec.Left = x * PIC_X
        If rec.Left + PIC_Y > DDSD_PanoInf.lWidth Then rec.Right = DDSD_PanoInf.lWidth: rec_pos.Right = rec_pos.Right - ((rec.Left + PIC_X) - DDSD_PanoInf.lWidth) Else rec.Right = rec.Left + PIC_X
        If Map(Player(MyIndex).Map).TranInf = 1 And TypeName(DD_PanoInfSurf) <> "Nothing" Then Call DD_BackBuffer.Blt(rec_pos, DD_PanoInfSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC) Else If TypeName(DD_PanoInfSurf) <> "Nothing" Then Call DD_BackBuffer.Blt(rec_pos, DD_PanoInfSurf, rec, DDBLT_WAIT)
    End If
End Sub

Sub BltItem(ByVal ItemNum As Long)
    ' Only used if ever want to switch to blt rather then bltfast
    'With rec_pos
        '.Top = MapItem(ItemNum).y * PIC_Y
        '.Bottom = .Top + PIC_Y
        '.Left = MapItem(ItemNum).x * PIC_X
        '.Right = .Left + PIC_X
    'End With
    
    rec.Top = (Item(MapItem(ItemNum).num).Pic \ 6) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Item(MapItem(ItemNum).num).Pic - (Item(MapItem(ItemNum).num).Pic \ 6) * 6) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_ItemSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast((MapItem(ItemNum).x - NewPlayerX) * PIC_X + sx - NewXOffset, (MapItem(ItemNum).y - NewPlayerY) * PIC_Y + sy - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltFog(ByVal MinX As Long, ByVal MaxX As Long, ByVal MinY As Long, ByVal MaxY As Long)
    'Initialisation du RECT source
    With rec_pos
        .Top = 0
        .Bottom = (MaxY - MinY + 1) * PIC_Y
        .Left = 0
        .Right = .Left + (MaxX - MinX + 1) * PIC_X
    End With
    
    'Initialisation du RECT destination
    With rec
        .Top = (NewPlayerY * 32) + NewYOffset
        .Bottom = .Top + rec_pos.Bottom
        .Left = (NewPlayerX * 32) + NewXOffset
        .Right = .Left + (MaxX - MinX + 1) * PIC_X
    End With
    
    'Dessin du brouillard
    Call AlphaBlendDX(rec_pos, rec, FogVerts)
End Sub

Sub BltFringeTile(ByVal x As Long, ByVal y As Long)
Dim Fringe As Long
Dim FAnim As Long
Dim Fringe2 As Long
Dim F2Anim As Long
Dim Fringe3 As Long '<--
Dim F3Anim As Long '<--
Dim FringeTileSet As Byte
Dim FAnimTileSet As Byte
Dim Fringe2TileSet As Byte
Dim F2AnimTileSet As Byte
Dim Fringe3TileSet As Byte '<--
Dim F3AnimTileSet As Byte '<--
Dim tx As Long
Dim ty As Long

    Fringe = Map(Player(MyIndex).Map).tile(x, y).Fringe
    FAnim = Map(Player(MyIndex).Map).tile(x, y).FAnim
    Fringe2 = Map(Player(MyIndex).Map).tile(x, y).Fringe2
    F2Anim = Map(Player(MyIndex).Map).tile(x, y).F2Anim
    Fringe3 = Map(Player(MyIndex).Map).tile(x, y).Fringe3 '<--
    F3Anim = Map(Player(MyIndex).Map).tile(x, y).F3Anim '<--
    
    FringeTileSet = Map(Player(MyIndex).Map).tile(x, y).FringeSet
    FAnimTileSet = Map(Player(MyIndex).Map).tile(x, y).FAnimSet
    Fringe2TileSet = Map(Player(MyIndex).Map).tile(x, y).Fringe2Set
    F2AnimTileSet = Map(Player(MyIndex).Map).tile(x, y).F2AnimSet
    Fringe3TileSet = Map(Player(MyIndex).Map).tile(x, y).Fringe3Set '<--
    F3AnimTileSet = Map(Player(MyIndex).Map).tile(x, y).F3AnimSet '<--
    
    tx = (x - NewPlayerX) * PIC_X + sx - NewXOffset
    ty = (y - NewPlayerY) * PIC_Y + sy - NewYOffset
    
    If (Not MapAnim) Or (FAnim <= 0) Then
        ' Is there an animation tile to plot?
        If Fringe > 0 And FringeTileSet <= ExtraSheets Then
            If Not TileFile(FringeTileSet) Then Exit Sub
            rec.Top = (Fringe \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Fringe - (Fringe \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(FringeTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If FAnim > 0 And FAnimTileSet <= ExtraSheets Then
            If Not TileFile(FAnimTileSet) Then Exit Sub
            rec.Top = (FAnim \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (FAnim - (FAnim \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(FAnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If

    If (Not MapAnim) Or (F2Anim <= 0) Then
        ' Is there an animation tile to plot?
        If Fringe2 > 0 And Fringe2TileSet <= ExtraSheets Then
            If Not TileFile(Fringe2TileSet) Then Exit Sub
            rec.Top = (Fringe2 \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Fringe2 - (Fringe2 \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(Fringe2TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If F2Anim > 0 And F2AnimTileSet <= ExtraSheets Then
            If Not TileFile(F2AnimTileSet) Then Exit Sub
            rec.Top = (F2Anim \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (F2Anim - (F2Anim \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(F2AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If (Not MapAnim) Or (F3Anim <= 0) Then  '<--
        ' Is there an animation tile to plot?
        If Fringe3 > 0 And Fringe3TileSet <= ExtraSheets Then
            If Not TileFile(Fringe3TileSet) Then Exit Sub
            rec.Top = (Fringe3 \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Fringe3 - (Fringe3 \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(Fringe3TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If F3Anim > 0 And F3AnimTileSet <= ExtraSheets Then
            If Not TileFile(F3AnimTileSet) Then Exit Sub
            rec.Top = (F3Anim \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (F3Anim - (F3Anim \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(F3AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If '<--
    'Affichage du panorama supérieur si il y en à un
    If Trim$(Map(Player(MyIndex).Map).PanoSup) <> vbNullString Then
        rec.Top = y * PIC_Y
        If rec.Top + PIC_Y > DDSD_PanoSup.lHeight Then rec.Bottom = DDSD_PanoSup.lHeight: rec_pos.Bottom = rec_pos.Bottom - ((rec.Top + PIC_Y) - DDSD_PanoSup.lHeight) Else rec.Bottom = rec.Top + PIC_Y
        rec.Left = x * PIC_X
        If rec.Left + PIC_Y > DDSD_PanoSup.lWidth Then rec.Right = DDSD_PanoSup.lWidth: rec_pos.Right = rec_pos.Right - ((rec.Left + PIC_X) - DDSD_PanoSup.lWidth) Else rec.Right = rec.Left + PIC_X
        If Map(Player(MyIndex).Map).TranSup = 1 And TypeName(DD_PanoSupSurf) <> "Nothing" Then Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_PanoSupSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY) Else If TypeName(DD_PanoSupSurf) <> "Nothing" Then Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_PanoSupSurf, rec, DDBLTFAST_WAIT)
    End If
End Sub

Sub BltVisu()
Dim Visu As Long
Dim VisuTileSet As Byte
If ScreenDC Then Exit Sub
    Visu = EditorTileY * TilesInSheets + EditorTileX
    VisuTileSet = EditorSet
 
    If Visu > 0 Then
        If Not TileFile(VisuTileSet) Then Exit Sub
        rec.Top = (Visu \ TilesInSheets) * PIC_Y
        rec.Bottom = rec.Top + frmMirage.shpSelected.Height
        rec.Left = (Visu - (Visu \ TilesInSheets) * TilesInSheets) * PIC_X
        rec.Right = rec.Left + frmMirage.shpSelected.Width
        'Set DD_Temp = DD_TileSurf(VisuTileSet)
        Call DD_BackBuffer.BltFast((CurX - NewPlayerX) * PIC_X + sx - NewXOffset, (CurY - NewPlayerY) * PIC_Y + sy - NewYOffset, DD_Temp, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Sub PreVisua()
Dim Visu As Long
Dim VisuTileSet As Byte
Dim i As Long
Dim t As Long
If ScreenDC Then Exit Sub
    Visu = EditorTileY * TilesInSheets + EditorTileX
    VisuTileSet = EditorSet
    
    Call DD_Temp.SetForeColor(RGB(0, 0, 0))
    For i = ((Visu \ TilesInSheets) * PIC_Y) To ((Visu \ TilesInSheets) * PIC_Y) + frmMirage.shpSelected.Height - 1 Step 2
        For t = (Visu - (Visu \ TilesInSheets) * TilesInSheets) * PIC_X To ((Visu - (Visu \ TilesInSheets) * TilesInSheets) * PIC_X) + frmMirage.shpSelected.Width - 1 Step 2
            Call DD_Temp.DrawLine(t, i, t + 1, i)
            Call DD_Temp.DrawLine(t + 1, i + 1, t + 2, i + 1)
        Next t
    Next i
End Sub

Sub BltPlayerOmbre(ByVal Index As Long)
Dim x As Long, y As Long

    If Index <= 0 And Index >= MAX_PLAYERS Then Exit Sub
    If Not IsPlaying(Index) Then Exit Sub

    x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset
    y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset
    
    rec.Top = 5 * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = 0 * PIC_X
    rec.Right = rec.Left + PIC_X
    
    Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPlayer(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim tx As Long, ty As Long
Dim AttackSpeed As Long
If Index <= 0 And Index >= MAX_PLAYERS Then Exit Sub
If Not IsPlaying(Index) Then Exit Sub
    
    If GetPlayerWeaponSlot(Index) > 0 Then
        If GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index)) > 0 Then
            AttackSpeed = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AttackSpeed
        Else: AttackSpeed = 1000: End If
    Else: AttackSpeed = 1000: End If
    
    ' Only used if ever want to switch to blt rather then bltfast
    'With rec_pos
        '.Top = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset
        '.Bottom = .Top + PIC_Y
        '.Left = GetPlayerX(Index) * PIC_X + Player(Index).XOffset
        '.Right = .Left + PIC_X
    'End With
   
    ' Check for animation
    Anim = 1
    If Player(Index).Attacking = 0 Or Player(Index).Moving > 0 Then
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).YOffset > PIC_Y / 2) Then Anim = Player(Index).Anim
            Case DIR_DOWN
                If (Player(Index).YOffset < PIC_Y / 2 * -1) Then Anim = Player(Index).Anim
            Case DIR_LEFT
                If (Player(Index).XOffset > PIC_Y / 2) Then Anim = Player(Index).Anim
            Case DIR_RIGHT
                If (Player(Index).XOffset < PIC_Y / 2 * -1) Then Anim = Player(Index).Anim
        End Select
    Else
        If Player(Index).AttackTimer + (AttackSpeed \ 2) > GetTickCount Then Anim = 2
    End If

    ' Check to see if we want to stop making him attack
    If Player(Index).AttackTimer + AttackSpeed < GetTickCount Then Player(Index).Attacking = 0: Player(Index).AttackTimer = 0
   
    ty = DDSD_Character(GetPlayerSprite(Index)).lHeight / 4
    tx = DDSD_Character(GetPlayerSprite(Index)).lWidth / 4
    
    rec.Top = GetPlayerDir(Index) * ty + (ty / 2)
    rec.Bottom = rec.Top + (ty / 2)
    rec.Left = Anim * tx + tx
    rec.Right = rec.Left + tx
    
    x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset - ((tx / 2) - 16)
    y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset
    
    
    If x < 0 Then rec.Left = rec.Left - x: rec.Right = rec.Left + (tx + x): x = 0
    If y < 0 Then rec.Top = rec.Top + (ty / 2): rec.Bottom = rec.Top: y = Player(Index).YOffset + sy
   
    Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_SpriteSurf(GetPlayerSprite(Index)), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    'PAPERDOLL
    If GetPlayerArmorSlot(Index) > 0 Then
        If Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).paperdoll = 1 Then
            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_PaperDollSurf(Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).paperdollPic), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If GetPlayerHelmetSlot(Index) > 0 Then
        If Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).paperdoll = 1 Then
            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_PaperDollSurf(Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).paperdollPic), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If GetPlayerWeaponSlot(Index) > 0 Then
        If Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).paperdoll = 1 Then
            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_PaperDollSurf(Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).paperdollPic), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If GetPlayerShieldSlot(Index) > 0 Then
        If Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).paperdoll = 1 Then
            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_PaperDollSurf(Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).paperdollPic), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    'FIN PAPERDOLL
End Sub

Sub BltPlayerPet(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long, tx As Long, ty As Long
Dim num As Long
If Index <= 0 And Index >= MAX_PLAYERS Then Exit Sub
If Player(Index).PetSlot <= 0 And Player(Index).PetSlot >= MAX_ITEMS Then Exit Sub

If Not IsPlaying(Index) Then Exit Sub
If Map(Player(MyIndex).Map).petView = 1 Then Exit Sub

    Anim = 1
    If Player(Index).Attacking = 0 Or Player(Index).Moving > 0 Then
        Select Case Player(Index).pet.Dir
            Case DIR_UP
                If (Player(Index).pet.YOffset > PIC_Y / 2) Then Anim = Player(Index).pet.Anim
            Case DIR_DOWN
                If (Player(Index).pet.YOffset < PIC_Y / 2 * -1) Then Anim = Player(Index).pet.Anim
            Case DIR_LEFT
                If (Player(Index).pet.XOffset > PIC_Y / 2) Then Anim = Player(Index).pet.Anim
            Case DIR_RIGHT
                If (Player(Index).pet.XOffset < PIC_Y / 2 * -1) Then Anim = Player(Index).pet.Anim
        End Select
    End If
       
    num = Pets(Item(GetPlayerInvItemNum(Index, GetPlayerPetSlot(Index))).Data1).sprite
    
    ty = DDSD_Pets(num).lHeight / 4
    tx = DDSD_Pets(num).lWidth / 4
    
    rec.Top = Player(Index).pet.Dir * ty
    rec.Bottom = rec.Top + ty
    rec.Left = Anim * tx + tx
    rec.Right = rec.Left + tx

        x = Player(Index).pet.x * PIC_X + sx + Player(Index).pet.XOffset - ((tx / 2) - 16)
        y = Player(Index).pet.y * PIC_Y + sx + Player(Index).pet.YOffset - (ty / 2)
        
    If x < 0 Then rec.Left = rec.Left - x: rec.Right = rec.Left + (tx + x): x = 0
    If y < 0 Then rec.Top = rec.Top + (ty / 2): rec.Bottom = rec.Top: y = Player(Index).YOffset + sy
   
    Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_PetsSurf(num), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPlayerTop(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim tx As Long, ty As Long
Dim AttackSpeed As Long
    
    If GetPlayerWeaponSlot(Index) > 0 Then
        If GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index)) > 0 Then
            AttackSpeed = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AttackSpeed
        Else: AttackSpeed = 1000: End If
    Else: AttackSpeed = 1000: End If

    ' Only used if ever want to switch to blt rather then bltfast
    'With rec_pos
        '.Top = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset
        '.Bottom = .Top + PIC_Y
        '.Left = GetPlayerX(Index) * PIC_X + Player(Index).XOffset
        '.Right = .Left + PIC_X
    'End With
   
    ' Check for animation
    Anim = 1
    If Player(Index).Attacking = 0 Then
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).YOffset > PIC_Y / 2) Then Anim = Player(Index).Anim
            Case DIR_DOWN
                If (Player(Index).YOffset < PIC_Y / 2 * -1) Then Anim = Player(Index).Anim
            Case DIR_LEFT
                If (Player(Index).XOffset > PIC_Y / 2) Then Anim = Player(Index).Anim
            Case DIR_RIGHT
                If (Player(Index).XOffset < PIC_Y / 2 * -1) Then Anim = Player(Index).Anim
        End Select
    Else
        If Player(Index).AttackTimer + (AttackSpeed \ 2) > GetTickCount Then Anim = 2
    End If
   
    ' Check to see if we want to stop making him attack
    If Player(Index).AttackTimer + AttackSpeed < GetTickCount Then Player(Index).Attacking = 0: Player(Index).AttackTimer = 0
              
    ty = DDSD_Character(GetPlayerSprite(Index)).lHeight / 4
    tx = DDSD_Character(GetPlayerSprite(Index)).lWidth / 4
    
    rec.Top = GetPlayerDir(Index) * ty
    rec.Bottom = rec.Top + (ty / 2)
    rec.Left = Anim * tx + tx
    rec.Right = rec.Left + tx
    
    x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset - ((tx / 2) - 16) '(tx / 4) - ((tx / 4) / 2)
    y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - (ty / 2)
    
    If x < 0 Then rec.Left = rec.Left - x: rec.Right = rec.Left + (tx + x): x = 0
    If y < 0 Then rec.Top = rec.Top + (ty / 2): rec.Bottom = rec.Top: y = Player(Index).YOffset + sy
    
     Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_SpriteSurf(GetPlayerSprite(Index)), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    'PAPERDOLL
    If GetPlayerArmorSlot(Index) > 0 Then
        If Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).paperdoll = 1 Then
            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_PaperDollSurf(Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).paperdollPic), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If GetPlayerHelmetSlot(Index) > 0 Then
        If Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).paperdoll = 1 Then
            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_PaperDollSurf(Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).paperdollPic), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If GetPlayerWeaponSlot(Index) > 0 Then
        If Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).paperdoll = 1 Then
            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_PaperDollSurf(Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).paperdollPic), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If GetPlayerShieldSlot(Index) > 0 Then
        If Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).paperdoll = 1 Then
            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_PaperDollSurf(Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).paperdollPic), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    'FIN PAPERDOLL
End Sub

Sub BltMapNPCName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim m As String

If Mid$(Trim$(Npc(MapNpc(Index).num).name), 1, 2) = "**" Then Exit Sub

    With Npc(MapNpc(Index).num)
    'Draw name
        TextX = MapNpc(Index).x * PIC_X + sx + MapNpc(Index).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.name)) / 2) * 8)
        If DDSD_Character(Npc(MapNpc(Index).num).sprite).lHeight = 128 And DDSD_Character(Npc(MapNpc(Index).num).sprite).lWidth = 128 Then
            TextY = MapNpc(Index).y * PIC_Y - 14 + MapNpc(Index).YOffset - CLng(PIC_Y / 2) + 48
        Else
            TextY = MapNpc(Index).y * PIC_Y - 14 + MapNpc(Index).YOffset - CLng(PIC_Y / 2) + 32
        End If
        If Npc(MapNpc(Index).num).Behavior = NPC_BEHAVIOR_QUETEUR Then
            DrawPlayerNameText TexthDC, TextX - NewPlayerPOffsetX, TextY - NewPlayerPOffsetY - (PIC_Y / 2), Trim$(.name), vbGreen
        Else
            DrawPlayerNameText TexthDC, TextX - NewPlayerPOffsetX, TextY - NewPlayerPOffsetY - (PIC_Y / 2), Trim$(.name), vbWhite
        End If
    End With

End Sub
Sub BltNpc(ByVal MapNpcNum As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim tx As Long, ty As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).num <= 0 Then Exit Sub
    
    ' Only used if ever want to switch to blt rather then bltfast
    'With rec_pos
        '.Top = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset
        '.Bottom = .Top + PIC_Y
        '.Left = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).XOffset
        '.Right = .Left + PIC_X
    'End With
    
    ' Check for animation
    Anim = 1
    If MapNpc(MapNpcNum).Attacking = 0 Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).YOffset > PIC_Y / 2) Then Anim = PNJAnim(MapNpcNum)
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2 * -1) Then Anim = PNJAnim(MapNpcNum)
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).XOffset > PIC_Y / 2) Then Anim = PNJAnim(MapNpcNum)
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2 * -1) Then Anim = PNJAnim(MapNpcNum)
        End Select
    Else
        If MapNpc(MapNpcNum).AttackTimer + 500 > GetTickCount Then Anim = 2
    End If
    
    ' Check to see if we want to stop making him attack
    If MapNpc(MapNpcNum).AttackTimer + 1000 < GetTickCount Then MapNpc(MapNpcNum).Attacking = 0: MapNpc(MapNpcNum).AttackTimer = 0
        ty = DDSD_Character(Npc(MapNpc(MapNpcNum).num).sprite).lHeight / 4
        tx = DDSD_Character(Npc(MapNpc(MapNpcNum).num).sprite).lWidth / 4
        
        rec.Top = MapNpc(MapNpcNum).Dir * ty + (ty / 2)
        rec.Bottom = rec.Top + (ty / 2)
        rec.Left = Anim * tx + tx
        rec.Right = rec.Left + tx
        
        x = MapNpc(MapNpcNum).x * PIC_X + sx + MapNpc(MapNpcNum).XOffset - ((tx / 2) - 16) '(tx / 4) - ((tx / 4) / 2)
        y = MapNpc(MapNpcNum).y * PIC_Y + sx + MapNpc(MapNpcNum).YOffset
        
        If x < 0 Then rec.Left = rec.Left - x: rec.Right = rec.Left + (tx + x): x = 0
        If y < 0 Then rec.Top = rec.Top + (ty / 2): rec.Bottom = rec.Top: y = MapNpc(MapNpcNum).YOffset + sy
            
        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_SpriteSurf(Npc(MapNpc(MapNpcNum).num).sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltNpcTop(ByVal MapNpcNum As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim tx As Long, ty As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).num <= 0 Then Exit Sub
        
    ' Only used if ever want to switch to blt rather then bltfast
    'With rec_pos
        '.Top = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset
        '.Bottom = .Top + PIC_Y
        '.Left = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).XOffset
        '.Right = .Left + PIC_X
    'End With
    
    ' Check for animation
    Anim = 1
    If MapNpc(MapNpcNum).Attacking = 0 Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).YOffset > PIC_Y / 2) Then Anim = PNJAnim(MapNpcNum)
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2 * -1) Then Anim = PNJAnim(MapNpcNum)
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).XOffset > PIC_Y / 2) Then Anim = PNJAnim(MapNpcNum)
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2 * -1) Then Anim = PNJAnim(MapNpcNum)
        End Select
    Else
        If MapNpc(MapNpcNum).AttackTimer + 500 > GetTickCount Then Anim = 2
    End If
    
    ' Check to see if we want to stop making him attack
    If MapNpc(MapNpcNum).AttackTimer + 1000 < GetTickCount Then MapNpc(MapNpcNum).Attacking = 0: MapNpc(MapNpcNum).AttackTimer = 0
    
    ty = DDSD_Character(Npc(MapNpc(MapNpcNum).num).sprite).lHeight / 4
    tx = DDSD_Character(Npc(MapNpc(MapNpcNum).num).sprite).lWidth / 4
    
    rec.Top = MapNpc(MapNpcNum).Dir * ty
    rec.Bottom = rec.Top + (ty / 2)
    rec.Left = Anim * tx + tx
    rec.Right = rec.Left + tx
    
    x = MapNpc(MapNpcNum).x * PIC_X + sx + MapNpc(MapNpcNum).XOffset - ((tx / 2) - 16) '(tx / 4) - ((tx / 4) / 2)
    y = MapNpc(MapNpcNum).y * PIC_Y + sx + MapNpc(MapNpcNum).YOffset - (ty / 2)
    
    If x < 0 Then rec.Left = rec.Left - x: rec.Right = rec.Left + (tx + x): x = 0
    If y < 0 Then rec.Top = rec.Top + (ty / 2): rec.Bottom = rec.Top: y = MapNpc(MapNpcNum).YOffset + sy
    
    Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_SpriteSurf(Npc(MapNpc(MapNpcNum).num).sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub


Sub BltPlayerLevelUp(ByVal Index As Long)
Dim x As Long
Dim y As Long
    rec.Top = (32 \ TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (32 - (32 \ TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + 96
    
    If Index = MyIndex Then
        x = NewX + sx
        y = NewY + sy
        Call DD_BackBuffer.BltFast(x - 32, y - 10 - Player(Index).LevelUp, DD_TileSurf(ExtraSheets), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset
        y = GetPlayerY(Index) * PIC_Y + sy + Player(Index).YOffset
        Call DD_BackBuffer.BltFast(x - NewPlayerPicX - 32 - NewXOffset, y - NewPlayerPicY - 10 - Player(Index).LevelUp - NewYOffset, DD_TileSurf(ExtraSheets), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    If Player(Index).LevelUp >= 3 Then Player(Index).LevelUp = Player(Index).LevelUp - 1 Else If Player(Index).LevelUp >= 1 Then Player(Index).LevelUp = Player(Index).LevelUp + 1
End Sub
                
Sub BltPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
    ' Check access level
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerAccess(Index)
            Case 0
                Color = QBColor(Brown)
            Case 1
                Color = AccModo
            Case 2
                Color = AccMapeur
            Case 3
                Color = AccDevelopeur
            Case 4
                Color = AccAdmin
        End Select
    Else
        Color = QBColor(BrightRed)
    End If
    
If Index = MyIndex Then
    TextX = NewX + sx + (PIC_X \ 2) - ((Len(GetPlayerName(MyIndex)) / 2) * 8)
    If DDSD_Character(GetPlayerSprite(Index)).lHeight = 128 And DDSD_Character(GetPlayerSprite(Index)).lWidth = 128 Then
        TextY = NewY + sy - 34 - ((PIC_NPC1 - 1) * 16) + 16
    Else
        TextY = NewY + sy - 34 - ((PIC_NPC1 - 1) * 16)
    End If
    Call DrawText(TexthDC, TextX, TextY, GetPlayerName(MyIndex), Color)
Else
    ' Draw name
    TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset + (PIC_X \ 2) - ((Len(GetPlayerName(Index)) / 2) * 8)

    If DDSD_Character(GetPlayerSprite(Index)).lHeight = 128 And DDSD_Character(GetPlayerSprite(Index)).lWidth = 128 Then
        TextY = GetPlayerY(Index) * PIC_Y + sy + Player(Index).YOffset - (PIC_Y \ 2) - 50
    Else
        TextY = GetPlayerY(Index) * PIC_Y + sy + Player(Index).YOffset - (PIC_Y \ 2) - 34
    End If
    
    Call DrawText(TexthDC, TextX - NewPlayerPOffsetX, TextY - NewPlayerPOffsetY, GetPlayerName(Index), Color)
End If
End Sub

Sub BltPlayerGuildName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
    ' Check access level
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerGuildAccess(Index)
            Case 0
                Color = QBColor(Red)
            Case 1
                Color = QBColor(BrightCyan)
            Case 2
                Color = QBColor(Yellow)
            Case 3
                Color = QBColor(BrightGreen)
            Case 4
                Color = QBColor(Yellow)
        End Select
    Else
        Color = QBColor(BrightRed)
    End If

If Index = MyIndex Then
    TextX = NewX + sx + (PIC_X \ 2) - ((Len(GetPlayerGuild(MyIndex)) / 2) * 8)
    TextY = NewY + sy - (PIC_Y \ 4) - 20 - ((PIC_NPC1 - 1) * 10)
    
    Call DrawText(TexthDC, TextX, TextY, GetPlayerGuild(MyIndex), Color)
Else
    ' Draw name
    TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset + (PIC_X \ 2) - ((Len(GetPlayerGuild(Index)) / 2) * 8)
    TextY = GetPlayerY(Index) * PIC_Y + sy + Player(Index).YOffset - (PIC_Y \ 2) - 12
    Call DrawText(TexthDC, TextX - NewPlayerPOffsetX, TextY - NewPlayerPOffsetY, GetPlayerGuild(Index), Color)
End If
End Sub

Sub ProcessMovement(ByVal Index As Long)
' vérifier si le joueur(sprite) ne va pas trop loin
If Player(Index).XOffset > PIC_X Or Player(Index).XOffset < PIC_X * -1 Then Player(Index).XOffset = 0: Player(Index).Moving = 0: Exit Sub
If Player(Index).YOffset > PIC_Y Or Player(Index).YOffset < PIC_Y * -1 Then Player(Index).YOffset = 0: Player(Index).Moving = 0: Exit Sub

' Verifier si le joueur à une monture
If Player(Index).ArmorSlot > 0 And Player(Index).ArmorSlot < MAX_INV Then
If GetPlayerInvItemNum(Index, Player(Index).ArmorSlot) > 0 And GetPlayerInvItemNum(Index, Player(Index).ArmorSlot) < MAX_ITEMS Then
If (Player(Index).Moving = MOVING_WALKING Or Player(Index).Moving = MOVING_RUNNING) And Item(GetPlayerInvItemNum(Index, Player(Index).ArmorSlot)).Type = ITEM_TYPE_MONTURE Then
        If Player(Index).Access > 0 Then
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - (GM_WALK_SPEED * Item(GetPlayerInvItemNum(Index, Player(Index).ArmorSlot)).Data2)
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + (GM_WALK_SPEED * Item(GetPlayerInvItemNum(Index, Player(Index).ArmorSlot)).Data2)
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - (GM_WALK_SPEED * Item(GetPlayerInvItemNum(Index, Player(Index).ArmorSlot)).Data2)
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + (GM_WALK_SPEED * Item(GetPlayerInvItemNum(Index, Player(Index).ArmorSlot)).Data2)
            End Select
            If Player(Index).PetSlot <> 0 Then
                If Player(Index).pet.YOffset <> 0 Or Player(Index).pet.XOffset <> 0 Then
                    Select Case Player(Index).pet.Dir
                        Case DIR_UP
                            Player(Index).pet.YOffset = Player(Index).pet.YOffset - (GM_WALK_SPEED * Item(GetPlayerInvItemNum(Index, Player(Index).ArmorSlot)).Data2)
                        Case DIR_DOWN
                            Player(Index).pet.YOffset = Player(Index).pet.YOffset + (GM_WALK_SPEED * Item(GetPlayerInvItemNum(Index, Player(Index).ArmorSlot)).Data2)
                        Case DIR_LEFT
                            Player(Index).pet.XOffset = Player(Index).pet.XOffset - (GM_WALK_SPEED * Item(GetPlayerInvItemNum(Index, Player(Index).ArmorSlot)).Data2)
                        Case DIR_RIGHT
                            Player(Index).pet.XOffset = Player(Index).pet.XOffset + (GM_WALK_SPEED * Item(GetPlayerInvItemNum(Index, Player(Index).ArmorSlot)).Data2)
                    End Select
                End If
            End If
        Else
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - (WALK_SPEED + ((WALK_SPEED / 100) * Item(GetPlayerInvItemNum(Index, Player(Index).ArmorSlot)).Data2))
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + (WALK_SPEED + ((WALK_SPEED / 100) * Item(GetPlayerInvItemNum(Index, Player(Index).ArmorSlot)).Data2))
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - (WALK_SPEED + ((WALK_SPEED / 100) * Item(GetPlayerInvItemNum(Index, Player(Index).ArmorSlot)).Data2))
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + (WALK_SPEED + ((WALK_SPEED / 100) * Item(GetPlayerInvItemNum(Index, Player(Index).ArmorSlot)).Data2))
            End Select
            If Player(Index).PetSlot <> 0 Then
                If Player(Index).pet.YOffset <> 0 Or Player(Index).pet.XOffset <> 0 Then
                    Select Case Player(Index).pet.Dir
                        Case DIR_UP
                            Player(Index).pet.YOffset = Player(Index).pet.YOffset - (WALK_SPEED + ((WALK_SPEED / 100) * Item(GetPlayerInvItemNum(Index, Player(Index).ArmorSlot)).Data2))
                        Case DIR_DOWN
                            Player(Index).pet.YOffset = Player(Index).pet.YOffset + (WALK_SPEED + ((WALK_SPEED / 100) * Item(GetPlayerInvItemNum(Index, Player(Index).ArmorSlot)).Data2))
                        Case DIR_LEFT
                            Player(Index).pet.XOffset = Player(Index).pet.XOffset - (WALK_SPEED + ((WALK_SPEED / 100) * Item(GetPlayerInvItemNum(Index, Player(Index).ArmorSlot)).Data2))
                        Case DIR_RIGHT
                            Player(Index).pet.XOffset = Player(Index).pet.XOffset + (WALK_SPEED + ((WALK_SPEED / 100) * Item(GetPlayerInvItemNum(Index, Player(Index).ArmorSlot)).Data2))
                    End Select
                End If
            End If
        End If

        ' Check if completed walking over to the next tile
        If (Player(Index).XOffset = 0) And (Player(Index).YOffset = 0) Then Player(Index).Moving = 0
        Exit Sub
End If
End If
End If

' Check if player is walking, and if so process moving them over
If Player(Index).Moving = MOVING_WALKING Then
        If Player(Index).Access > 0 Then
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - GM_WALK_SPEED
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + GM_WALK_SPEED
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - GM_WALK_SPEED
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + GM_WALK_SPEED
            End Select
            If Player(Index).PetSlot <> 0 Then
                If Player(Index).pet.YOffset <> 0 Or Player(Index).pet.XOffset <> 0 Then
                    Select Case Player(Index).pet.Dir
                        Case DIR_UP
                            Player(Index).pet.YOffset = Player(Index).pet.YOffset - GM_WALK_SPEED
                        Case DIR_DOWN
                            Player(Index).pet.YOffset = Player(Index).pet.YOffset + GM_WALK_SPEED
                        Case DIR_LEFT
                            Player(Index).pet.XOffset = Player(Index).pet.XOffset - GM_WALK_SPEED
                        Case DIR_RIGHT
                            Player(Index).pet.XOffset = Player(Index).pet.XOffset + GM_WALK_SPEED
                    End Select
                End If
            End If
        Else
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - WALK_SPEED
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + WALK_SPEED
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - WALK_SPEED
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + WALK_SPEED
            End Select
            If Player(Index).PetSlot <> 0 Then
                If Player(Index).pet.YOffset <> 0 Or Player(Index).pet.XOffset <> 0 Then
                    Select Case Player(Index).pet.Dir
                        Case DIR_UP
                            Player(Index).pet.YOffset = Player(Index).pet.YOffset - WALK_SPEED
                        Case DIR_DOWN
                            Player(Index).pet.YOffset = Player(Index).pet.YOffset + WALK_SPEED
                        Case DIR_LEFT
                            Player(Index).pet.XOffset = Player(Index).pet.XOffset - WALK_SPEED
                        Case DIR_RIGHT
                            Player(Index).pet.XOffset = Player(Index).pet.XOffset + WALK_SPEED
                    End Select
                End If
            End If
        End If
        
        ' Check if completed walking over to the next tile
        If (Player(Index).XOffset = 0) And (Player(Index).YOffset = 0) Then Player(Index).Moving = 0
   ' Check if player is running, and if so process moving them over
ElseIf Player(Index).Moving = MOVING_RUNNING Then
        If Player(Index).Access > 0 Then
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - GM_RUN_SPEED
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + GM_RUN_SPEED
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - GM_RUN_SPEED
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + GM_RUN_SPEED
            End Select
            If Player(Index).PetSlot <> 0 Then
                If Player(Index).pet.YOffset <> 0 Or Player(Index).pet.XOffset <> 0 Then
                    Select Case Player(Index).pet.Dir
                        Case DIR_UP
                            Player(Index).pet.YOffset = Player(Index).pet.YOffset - GM_RUN_SPEED
                        Case DIR_DOWN
                            Player(Index).pet.YOffset = Player(Index).pet.YOffset + GM_RUN_SPEED
                        Case DIR_LEFT
                            Player(Index).pet.XOffset = Player(Index).pet.XOffset - GM_RUN_SPEED
                        Case DIR_RIGHT
                            Player(Index).pet.XOffset = Player(Index).pet.XOffset + GM_RUN_SPEED
                    End Select
                End If
            End If
        Else
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - RUN_SPEED
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + RUN_SPEED
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - RUN_SPEED
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + RUN_SPEED
            End Select
            If Player(Index).PetSlot <> 0 Then
                If Player(Index).pet.YOffset <> 0 Or Player(Index).pet.XOffset <> 0 Then
                    Select Case Player(Index).pet.Dir
                        Case DIR_UP
                            Player(Index).pet.YOffset = Player(Index).pet.YOffset - RUN_SPEED
                        Case DIR_DOWN
                            Player(Index).pet.YOffset = Player(Index).pet.YOffset + RUN_SPEED
                        Case DIR_LEFT
                            Player(Index).pet.XOffset = Player(Index).pet.XOffset - RUN_SPEED
                        Case DIR_RIGHT
                            Player(Index).pet.XOffset = Player(Index).pet.XOffset + RUN_SPEED
                    End Select
                End If
            End If
        End If
        
        ' Check if completed walking over to the next tile
        If (Player(Index).XOffset = 0) And (Player(Index).YOffset = 0) Then Player(Index).Moving = 0
End If
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
    ' Check if npc is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset - WALK_SPEED
            Case DIR_DOWN
                MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset + WALK_SPEED
            Case DIR_LEFT
                MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset - WALK_SPEED
            Case DIR_RIGHT
                MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset + WALK_SPEED
        End Select
        
        ' Check if completed walking over to the next tile
        If (MapNpc(MapNpcNum).XOffset = 0) And (MapNpc(MapNpcNum).YOffset = 0) Then MapNpc(MapNpcNum).Moving = 0
    End If
End Sub

Sub HandleKeypresses(ByVal KeyAscii As Integer)
Dim ChatText As String
Dim name As String
Dim i As Long
Dim n As Long
Dim Cod As String
Dim PX As Long
Dim PY As Long
Dim tp As Long
'MyText = frmMirage.txtMyTextBox.Text
If Len(frmMirage.txtMyTextBox.Text) > 200 Then
    MyText = Left(frmMirage.txtMyTextBox.Text, 200)
Else
    MyText = frmMirage.txtMyTextBox.Text
End If
' Handle when the player presses the return key
    If (KeyAscii = vbKeyReturn) Then
    frmMirage.txtMyTextBox.Text = vbNullString
        On Error GoTo er:
        
        PX = 0
        PY = 0
        If Player(MyIndex).y - 1 > -1 And PX = 0 And PY = 0 Then
            tp = Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type
            If tp = TILE_TYPE_COFFRE Or tp = TILE_TYPE_PORTE_CODE And Player(MyIndex).Dir = DIR_UP Then PX = 0: PY = -1
        End If
                
        If Player(MyIndex).y + 1 < MAX_MAPY + 1 And PX = 0 And PY = 0 Then
            tp = Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type
            If tp = TILE_TYPE_COFFRE Or tp = TILE_TYPE_PORTE_CODE And Player(MyIndex).Dir = DIR_DOWN Then PX = 0: PY = 1
        End If
                
        If Player(MyIndex).x - 1 > -1 And PX = 0 And PY = 0 Then
            tp = Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type
            If tp = TILE_TYPE_COFFRE Or tp = TILE_TYPE_PORTE_CODE And Player(MyIndex).Dir = DIR_LEFT Then PX = -1: PY = 0
        End If
        
        If Player(MyIndex).x + 1 < MAX_MAPX + 1 And PX = 0 And PY = 0 Then
            tp = Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type
            If tp = TILE_TYPE_COFFRE Or tp = TILE_TYPE_PORTE_CODE And Player(MyIndex).Dir = DIR_RIGHT Then PX = 1: PY = 0
        End If
        
        If PX <> 0 Or PY <> 0 Then
        With Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY)
            If .String1 > vbNullString And TempTile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).DoorOpen = NO Then
                Dim Packet As String
                Cod = InputBox("Veuillez entre le mot de passe :", "Code")
                If Cod = .String1 Then
                    TempTile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).DoorOpen = YES
                    Packet = "OUVRIRE" & SEP_CHAR & GetPlayerX(MyIndex) + PX & SEP_CHAR & GetPlayerY(MyIndex) + PY & END_CHAR
                    Call SendData(Packet)
                    If .Type = TILE_TYPE_COFFRE Then
                        i = FindOpenInvSlot(Val(.Data3))
                        If i > 0 Then
                            Call SetPlayerInvItemNum(MyIndex, i, Val(.Data3))
                            Call SetPlayerInvItemValue(MyIndex, i, GetPlayerInvItemValue(MyIndex, i) + 1)
                            Call SetPlayerInvItemDur(MyIndex, i, Item(Val(.Data3)).Data1)
                            Call UpdateVisInv
                            Packet = "ACOFFRE" & SEP_CHAR & i & SEP_CHAR & Val(.Data3) & SEP_CHAR & 1 & SEP_CHAR & Item(Val(.Data3)).Data1 & END_CHAR
                            Call SendData(Packet)
                        End If
                    End If
                Else
                    Call MsgBox("Mauvais code!", vbCritical)
                End If
            End If
        End With
        End If
        
        If GetPlayerY(MyIndex) - 1 > 0 And GetPlayerY(MyIndex) - 1 < MAX_MAPY Then
        With Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1)
            If .Type = TILE_TYPE_SIGN And Player(MyIndex).Dir = DIR_UP Then
                If Trim$(.String1) <> vbNullString Then Call QueteMsg(MyIndex, "Il est marqué :" & Trim$(.String1))
                If Trim$(.String2) <> vbNullString Then Call QueteMsg(MyIndex, "Il est marqué :" & Trim$(.String2))
                If Trim$(.String3) <> vbNullString Then Call QueteMsg(MyIndex, "Il est marqué :" & Trim$(.String3))
                Exit Sub
            End If
        End With
        End If
        
        ' Broadcast message
        If Mid$(MyText, 1, 1) = "'" Then
            ChatText = Mid$(MyText, 2, Len(MyText) - 1)
            If Len(Trim$(ChatText)) > 0 Then Call BroadcastMsg(ChatText)
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Emote message
        If Mid$(MyText, 1, 1) = "-" Then
            ChatText = Mid$(MyText, 2, Len(MyText) - 1)
            If Len(Trim$(ChatText)) > 0 Then Call EmoteMsg(ChatText)
            MyText = vbNullString
            Exit Sub
        End If
        
        ' message de guilde
       If Mid$(MyText, 1, 1) = "*" Then
           ChatText = Mid$(MyText, 2, Len(MyText) - 1)
           If Len(Trim$(ChatText)) > 0 Then Call GuildeMsg(ChatText)
           MyText = vbNullString
           Exit Sub
       End If
       
        ' Player message
        If Mid$(MyText, 1, 1) = "!" Or Mid$(MyText, 1, 1) = "w" Or Mid$(MyText, 1, 1) = "/w " Then
            ChatText = Mid$(MyText, 2, Len(MyText) - 1)
            name = vbNullString
                    
            ' Get the desired player from the user text
            For i = 1 To Len(ChatText)
                If Mid$(ChatText, i, 1) <> " " Then name = name & Mid$(ChatText, i, 1) Else Exit For
            Next i
                    
            ' Make sure they are actually sending something
            If Len(ChatText) - i > 0 Then
                ChatText = Mid$(ChatText, i + 1, Len(ChatText) - i)
                    
                ' Send the message to the player
                Call PlayerMsg(ChatText, name)
            Else
                Call AddText("Utiliser: !nomjoueur msgici", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If
            
        ' // Commands //
                
        ' Verification User
        If LCase$(Mid$(MyText, 1, 5)) = "/info" Then
            ChatText = Mid$(MyText, 6, Len(MyText) - 5)
            Call SendData("playerinforequest" & SEP_CHAR & ChatText & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Whos Online
        If LCase$(Mid$(MyText, 1, 4)) = "/who" Or LCase$(Mid$(MyText, 1, 4)) = "/qui" Then
            Call SendWhosOnline
            MyText = vbNullString
            Exit Sub
        End If
                        
        ' Checking fps
        If LCase$(Mid$(MyText, 1, 4)) = "/fps" Then
            Call AddText("FPS: " & GameFPS, Pink)
            MyText = vbNullString
            Exit Sub
        End If
                
        ' Show inventory
        If LCase$(Mid$(MyText, 1, 4)) = "/inv" Then
            frmMirage.picInv3.Visible = True
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Request stats
        If LCase$(Mid$(MyText, 1, 6)) = "/stats" Then
            Call SendData("getstats" & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
         
        ' Refresh Player
        If LCase$(Mid$(MyText, 1, 8)) = "/refresh" Then
            ConOff = True
            Call SendData("refresh" & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Decline Chat
        If LCase$(Mid$(MyText, 1, 12)) = "/chatdecline" Or LCase$(Mid$(MyText, 1, 12)) = "/chatrefu" Then
            Call SendData("dchat" & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Accept Chat
        If LCase$(Mid$(MyText, 1, 5)) = "/chat" Then
            Call SendData("achat" & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
        
        If LCase$(Mid$(MyText, 1, 6)) = "/trade" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 7 Then
                ChatText = Mid$(MyText, 8, Len(MyText) - 7)
                Call SendTradeRequest(ChatText)
            Else
                Call AddText("Utiliser : /echange nomdujoueur", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If
        
        If LCase$(Mid$(MyText, 1, 8)) = "/echange" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 9 Then
                ChatText = Mid$(MyText, 9, Len(MyText) - 8)
                Call SendTradeRequest(ChatText)
            Else
                Call AddText("Utiliser : /echange nomdujoueur", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Accept Trade
        If LCase$(Mid$(MyText, 1, 7)) = "/accept" Then
            Call SendAcceptTrade
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Decline Trade
        If LCase$(Mid$(MyText, 1, 8)) = "/decline" Or LCase$(Mid$(MyText, 1, 8)) = "/refu" Then
            Call SendDeclineTrade
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Party request
        If LCase$(Mid$(MyText, 1, 6)) = "/party" Or LCase$(Mid$(MyText, 1, 6)) = "/group" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 7 Then
                ChatText = Mid$(MyText, 8, Len(MyText) - 7)
                Call SendPartyRequest(ChatText)
            Else
                Call AddText("Utiliser : /groupe nomdujoueur", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Join party
        If LCase$(Mid$(MyText, 1, 5)) = "/join" Or LCase$(Mid$(MyText, 1, 5)) = "/rejoin" Then
            Call SendJoinParty
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Leave party
        If LCase$(Mid$(MyText, 1, 6)) = "/leave" Or LCase$(Mid$(MyText, 1, 6)) = "/quitte" Then
            Call SendLeaveParty
            MyText = vbNullString
            Exit Sub
        End If
        
        ' // Moniter Admin Commands //
        If GetPlayerAccess(MyIndex) > 0 Then
            ' day night command
            If LCase$(Mid$(MyText, 1, 9)) = "/daynight" Or LCase$(Mid$(MyText, 1, 9)) = "/journuit" Then
                If GameTime = TIME_DAY Then GameTime = TIME_NIGHT: Call InitNightAndFog(Player(MyIndex).Map) Else GameTime = TIME_DAY
                Call SendGameTime
                MyText = vbNullString
                Exit Sub
            End If
            
            ' weather command
            If LCase$(Mid$(MyText, 1, 8)) = "/weather" Or LCase$(Mid$(MyText, 1, 8)) = "/temp" Then
                If Len(MyText) > 8 Then
                    MyText = Mid$(MyText, 9, Len(MyText) - 8)
                    If IsNumeric(MyText) Then
                        Call SendData("weather" & SEP_CHAR & Val(MyText) & END_CHAR)
                    Else
                        If Trim$(LCase$(MyText)) = "none" Then i = 0
                        If Trim$(LCase$(MyText)) = "rain" Then i = 1
                        If Trim$(LCase$(MyText)) = "snow" Then i = 2
                        If Trim$(LCase$(MyText)) = "thunder" Then i = 3
                        Call SendData("weather" & SEP_CHAR & i & END_CHAR)
                    End If
                End If
                MyText = vbNullString
                Exit Sub
            End If

            ' Global Message
            If Mid$(MyText, 1, 1) = """" Then
                ChatText = Mid$(MyText, 2, Len(MyText) - 1)
                If Len(Trim$(ChatText)) > 0 Then Call GlobalMsg(ChatText)
                MyText = vbNullString
                Exit Sub
            End If
        
            ' Admin Message
            If Mid$(MyText, 1, 1) = "=" Then
                ChatText = Mid$(MyText, 2, Len(MyText) - 1)
                If Len(Trim$(ChatText)) > 0 Then Call AdminMsg(ChatText)
                MyText = vbNullString
                Exit Sub
            End If
        End If
        
        ' // Mapper Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
            ' Location
            If LCase$(Mid$(MyText, 1, 4)) = "/loc" Then
                Call SendRequestLocation
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Setting sprite
            If LCase$(Mid$(MyText, 1, 10)) = "/setsprite" Then
                If Len(MyText) > 11 Then
                    ' Get sprite #
                    MyText = Mid$(MyText, 12, Len(MyText) - 11)
                    Call SendSetSprite(Val(MyText))
                End If
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Setting player sprite
            If LCase$(Mid$(MyText, 1, 16)) = "/setplayersprite" Then
                If Len(MyText) > 19 Then
                    i = Val(Mid$(MyText, 17, 1))
                    MyText = Mid$(MyText, 18, Len(MyText) - 17)
                    Call SendSetPlayerSprite(i, Val(MyText))
                End If
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Changement de nom de joueur
            If LCase$(Mid$(MyText, 1, 16)) = "/setplayername" Then
                If Len(MyText) > 19 Then
                    i = Val(Mid$(MyText, 17, 1))
                    MyText = Mid$(MyText, 18, Len(MyText) - 17)
                    Call SendSetPlayerName(i, Val(MyText))
                End If
                MyText = vbNullString
                Exit Sub
            End If
        
            ' Respawn request
            If Mid$(MyText, 1, 8) = "/respawn" Then
                Call SendMapRespawn
                MyText = vbNullString
                Exit Sub
            End If
        
            ' MOTD change
            If Mid$(MyText, 1, 5) = "/motd" Then
                If Len(MyText) > 6 Then
                    MyText = Mid$(MyText, 7, Len(MyText) - 6)
                    If Trim$(MyText) <> vbNullString Then Call SendMOTDChange(MyText)
                End If
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Check the ban list
            If Mid$(MyText, 1, 3) = "/banlist" Then
                Call SendBanList
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Banning a player
            If LCase$(Mid$(MyText, 1, 4)) = "/ban" Then
                If Len(MyText) > 5 Then
                    MyText = Mid$(MyText, 6, Len(MyText) - 5)
                    Call SendBan(MyText)
                    MyText = vbNullString
                End If
                Exit Sub
            End If
        End If
            
        ' // Developer Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
            ' Editing item request
            If Mid$(MyText, 1, 9) = "/edititem" Then
                Call SendRequestEditItem
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Editing emoticon request
            If Mid$(MyText, 1, 13) = "/editemoticon" Then
                Call SendRequestEditEmoticon
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Editing arrow request
            If Mid$(MyText, 1, 13) = "/editarrow" Then
                Call SendRequestEditArrow
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Editing npc request
            If Mid$(MyText, 1, 8) = "/editnpc" Then
                Call SendRequestEditNpc
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Editing shop request
            If Mid$(MyText, 1, 9) = "/editshop" Then
                Call SendRequestEditShop
                MyText = vbNullString
                Exit Sub
            End If
        
            ' Editing spell request
            If LCase$(Trim$(MyText)) = "/editspell" Then
                Call SendRequestEditSpell
                MyText = vbNullString
                Exit Sub
            End If
        End If
        
        ' // Creator Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
            ' Giving another player access
            If LCase$(Mid$(MyText, 1, 10)) = "/setaccess" Then
                ' Get access #
                i = Val(Mid$(MyText, 12, 1))
                MyText = Mid$(MyText, 14, Len(MyText) - 13)
                Call SendSetAccess(MyText, i)
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Ban destroy
            If LCase$(Mid$(MyText, 1, 15)) = "/destroybanlist" Then
                Call SendBanDestroy
                MyText = vbNullString
                Exit Sub
            End If
        End If
        
        ' Tell them its not a valid command
        If Left$(Trim$(MyText), 1) = "/" Then
            For i = 0 To MAX_EMOTICONS
                If Trim$(Emoticons(i).Command) = Trim$(MyText) And Trim$(Emoticons(i).Command) <> "/" Then
                    Call SendData("checkemoticons" & SEP_CHAR & i & END_CHAR)
                    MyText = vbNullString
                Exit Sub
                End If
            Next i
            Call SendData("checkcommands" & SEP_CHAR & MyText & END_CHAR)
            MyText = vbNullString
        Exit Sub
        End If
            
        ' Say message
        If Len(Trim$(MyText)) > 0 Then
        '//Début du code de canaux
    If Len(Trim$(MyText)) > 0 Then
            If frmMirage.Canal.ListIndex = 1 Then
                Call BroadcastMsg(MyText)
                MyText = vbNullString
                Exit Sub
            ElseIf frmMirage.Canal.ListIndex = 2 Then
                Call GuildeMsg(MyText)
                MyText = vbNullString
                Exit Sub
            ElseIf frmMirage.Canal.ListIndex = 3 Then
                name = vbNullString
                   
                For i = 1 To Len(MyText)
                    If Mid$(MyText, i, 1) <> " " Then name = name & Mid$(MyText, i, 1) Else Exit For
                Next i
                   
                If Len(MyText) - i > 0 Then
                    MyText = Mid$(MyText, i + 1, Len(MyText) - i)
                   
                    Call PlayerMsg(MyText, name)
                Else
                    Call AddText("Vous avez oublié le nom du joueur", AlertColor)
                End If
                    MyText = vbNullString
                    Exit Sub
            ElseIf frmMirage.Canal.ListIndex = 4 Then
                MyText = "Commerce : " & MyText
                Call BroadcastMsg(MyText)
                MyText = vbNullString
                Exit Sub
            ElseIf frmMirage.Canal.ListIndex = 0 Then
                Call SayMsg(MyText)
            Else
                Call SayMsg(MyText)
            End If
    End If
'//Fin du code de canaux
        MyText = vbNullString
        Exit Sub
    End If
Exit Sub
er:
MsgBox "Erreur dans le code de Texte(" & Err.Number & " : " & Err.description & ")" & vbCrLf & "Merci de la rapporter sur le forum de FRoG Creator si elle persiste."
End If
End Sub

Sub Horsligne()
Dim i As Long
If HORS_LIGNE < 1 Then Exit Sub
    MyIndex = 1
    Call ClearConstante
    If ReadINI("INFO", "GameName", App.Path & "\config.ini") > vbNullString Then GAME_NAME = ReadINI("INFO", "GameName", App.Path & "\config.ini")
    If Val(ReadINI("INFO", "Maxplayers", App.Path & "\config.ini")) > 0 Then MAX_PLAYERS = ReadINI("INFO", "Maxplayers", App.Path & "\config.ini")
    If Val(ReadINI("INFO", "Maxitems", App.Path & "\config.ini")) > 0 Then MAX_ITEMS = ReadINI("INFO", "Maxitems", App.Path & "\config.ini")
    If Val(ReadINI("INFO", "Maxnpcs", App.Path & "\config.ini")) > 0 Then MAX_NPCS = ReadINI("INFO", "Maxnpcs", App.Path & "\config.ini")
    If Val(ReadINI("INFO", "Maxshops", App.Path & "\config.ini")) > 0 Then MAX_SHOPS = ReadINI("INFO", "Maxshops", App.Path & "\config.ini")
    If Val(ReadINI("INFO", "Maxspells", App.Path & "\config.ini")) > 0 Then MAX_SPELLS = ReadINI("INFO", "Maxspells", App.Path & "\config.ini")
    If Val(ReadINI("INFO", "Maxmaps", App.Path & "\config.ini")) > 0 Then MAX_MAPS = ReadINI("INFO", "Maxmaps", App.Path & "\config.ini")
    If Val(ReadINI("INFO", "Maxmapitems", App.Path & "\config.ini")) > 0 Then MAX_MAP_ITEMS = ReadINI("INFO", "Maxmapitems", App.Path & "\config.ini")
    If Val(ReadINI("INFO", "Maxmapx", App.Path & "\config.ini")) > 0 Then MAX_MAPX = ReadINI("INFO", "Maxmapx", App.Path & "\config.ini")
    If Val(ReadINI("INFO", "Maxmapy", App.Path & "\config.ini")) > 0 Then MAX_MAPY = ReadINI("INFO", "Maxmapy", App.Path & "\config.ini")
    If Val(ReadINI("INFO", "Maxemots", App.Path & "\config.ini")) > 0 Then MAX_EMOTICONS = ReadINI("INFO", "Maxemots", App.Path & "\config.ini")
    If Val(ReadINI("INFO", "Maxclasses", App.Path & "\config.ini")) > 0 Then Max_Classes = Val(ReadINI("INFO", "Maxclasses", App.Path & "\config.ini"))
    If Val(ReadINI("INFO", "Maxlevel", App.Path & "\config.ini")) > 0 Then MAX_LEVEL = Val(ReadINI("INFO", "Maxlevel", App.Path & "\config.ini"))
    If Val(ReadINI("INFO", "Maxquet", App.Path & "\config.ini")) > 0 Then MAX_QUETES = Val(ReadINI("INFO", "Maxquet", App.Path & "\config.ini"))
    If Val(ReadINI("INFO", "Maxnpcspell", App.Path & "\config.ini")) > 0 Then MAX_NPC_SPELLS = Val(ReadINI("INFO", "Maxnpcspell", App.Path & "\config.ini")) Else MAX_NPC_SPELLS = 10
    If Val(ReadINI("INFO", "Maxinv", App.Path & "\config.ini")) > 0 Then MAX_INV = Val(ReadINI("INFO", "Maxinv", App.Path & "\config.ini"))
    If Val(ReadINI("INFO", "Maxpets", App.Path & "\config.ini")) > 0 Then MAX_PETS = ReadINI("INFO", "Maxpets", App.Path & "\config.ini")
    If Val(ReadINI("INFO", "Maxmetier", App.Path & "\config.ini")) > 0 Then MAX_METIER = ReadINI("INFO", "Maxmetier", App.Path & "\config.ini")
    
    PIC_PL = Val(ReadINI("INFO", "PIC_PL", App.Path & "\config.ini"))
    PIC_NPC1 = Val(ReadINI("INFO", "PIC_NPC1", App.Path & "\config.ini"))
    PIC_NPC2 = Val(ReadINI("INFO", "PIC_NPC2", App.Path & "\config.ini"))
    'MAX_PLAYER_SPELLS = Val(ReadINI("INFO", "Maxpspel", App.Path & "\Editeur\config.ini"))
    ReDim Pets(1 To MAX_PETS) As PetsRec
    ReDim recette(1 To MAX_RECETTE) As RecetteRec
    ReDim Metier(1 To MAX_METIER) As MetierRec
    ReDim Class(0 To Max_Classes) As ClassRec
    ReDim quete(1 To MAX_QUETES) As QueteRec
    ReDim Map(1 To MAX_MAPS) As MapRec
    ReDim Player(1 To MAX_PLAYERS) As PlayerRec
    ReDim Item(1 To MAX_ITEMS) As ItemRec
    ReDim Npc(1 To MAX_NPCS) As NpcRec
    ReDim MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    ReDim Bubble(1 To MAX_PLAYERS) As ChatBubble
    ReDim SaveMapItem(1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim Experience(1 To MAX_LEVEL) As Long
    For i = 1 To MAX_MAPS
        ReDim Map(i).tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        ReDim Map(i).tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Next i
    For i = 0 To 5
        ReDim TempMap(i).tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Next i
    ReDim TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
    ReDim Emoticons(0 To MAX_EMOTICONS) As EmoRec
    ReDim MapReport(1 To MAX_MAPS) As MapRec
    MAX_SPELL_ANIM = MAX_MAPX * MAX_MAPY
    MAX_BLT_LINE = 6
    ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    
    For i = 1 To MAX_PLAYERS
        ReDim Player(i).SpellAnim(1 To MAX_SPELL_ANIM) As SpellAnimRec
        ReDim Player(i).inv(1 To MAX_INV) As PlayerInvRec
    Next i
    
    For i = 1 To MAX_NPCS
        ReDim Npc(i).Spell(1 To MAX_NPC_SPELLS) As Integer
    Next i
    
    For i = 0 To MAX_EMOTICONS
        Emoticons(i).Pic = 0
        Emoticons(i).Command = vbNullString
    Next i
    
    Call ClearTempTile
    
    ' Clear out players
    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
    Next i
    
    For i = 1 To MAX_MAPS
        DoEvents
        Call LoadMap(i)
    Next i
    
    frmMirage.Caption = "Editeur pour le jeu : " & Trim$(GAME_NAME) & " Mettez votre souris sur un élément pour plus de détails."
    App.Title = GAME_NAME
    If Not FileExiste("Stats.ini") Then
        Call WriteINI("HP", "AddPerLevel", 10, App.Path & "\Stats.ini")
        Call WriteINI("HP", "AddPerStr", 10, App.Path & "\Stats.ini")
        Call WriteINI("HP", "AddPerDef", 0, App.Path & "\Stats.ini")
        Call WriteINI("HP", "AddPerMagi", 0, App.Path & "\Stats.ini")
        Call WriteINI("HP", "AddPerSpeed", 0, App.Path & "\Stats.ini")
        Call WriteINI("MP", "AddPerLevel", 10, App.Path & "\Stats.ini")
        Call WriteINI("MP", "AddPerStr", 0, App.Path & "\Stats.ini")
        Call WriteINI("MP", "AddPerDef", 0, App.Path & "\Stats.ini")
        Call WriteINI("MP", "AddPerMagi", 10, App.Path & "\Stats.ini")
        Call WriteINI("MP", "AddPerSpeed", 0, App.Path & "\Stats.ini")
        Call WriteINI("SP", "AddPerLevel", 10, App.Path & "\Stats.ini")
        Call WriteINI("SP", "AddPerStr", 0, App.Path & "\Stats.ini")
        Call WriteINI("SP", "AddPerDef", 0, App.Path & "\Stats.ini")
        Call WriteINI("SP", "AddPerMagi", 0, App.Path & "\Stats.ini")
        Call WriteINI("SP", "AddPerSpeed", 20, App.Path & "\Stats.ini")
    End If
    
    Call StopMidi
    Call ChargerJoueure(MyIndex)
    Call ChargerCartes
    Call ChargerObjets(MyIndex)
    Call ChargerFleche
    Call ChargerEmots
    Call ChargerExps
    Call ChargerPnjs
    Call ChargerMagasins
    Call ChargerRecette
    Call ChargerSorts
    Call ChargerQuetes
    Call ChargerLeTemps
    Call ChargerClasses
    Call InitMirage
    Call PlayerMsg("Bienvenue dans " & GAME_NAME & "!", 15)
End Sub

Public Sub InitMirageVars()
    PicScWidth = frmMirage.picScreen.Width
    PicScHeight = frmMirage.picScreen.Height
End Sub

Sub InitMirage()
Dim i As Long
    frmMirage.Toolbar1.buttons(1).Enabled = False
    frmMirage.test.Enabled = False
    frmMirage.envoicarte.Enabled = False
    frmMirage.comtest.Enabled = False
    frmMirage.modo.Visible = False
    frmMirage.admin.Visible = False
    frmMirage.envserv.Enabled = False
    frmMirage.opti.Enabled = False
    Call frmMirage.NetPic
    Call StopMidi
    frmMirage.lstIndex.Clear
    For i = 1 To MAX_MAPS
        frmMirage.lstIndex.AddItem i & " : " & Map(i).name
    Next i
    frmsplash.Visible = False
    InGame = True
    Call InitDirectX
    Call EditorInit
    If ExtraSheets < frmMirage.Tiles.Count - 1 Then
        For i = ExtraSheets To 5
            Unload frmMirage.Tiles(i)
            Call frmMirage.tilescmb.RemoveItem(i)
        Next i
    Else
        For i = 0 To ExtraSheets
            If i > frmMirage.Tiles.Count - 1 Then Load frmMirage.Tiles(i): frmMirage.Tiles(i).Caption = "Tiles" & i: frmMirage.Tiles(i).Checked = False
            If i > frmMirage.tilescmb.ListCount - 1 Then Call frmMirage.tilescmb.AddItem("Tiles" & i, i)
        Next i
    End If
    frmMirage.Show
    Call GameLoop
End Sub
Sub ClearConstante()
GAME_NAME = "Frog Creator"
MAX_PLAYERS = 50
MAX_ITEMS = 100
MAX_NPCS = 100
MAX_SHOPS = 100
MAX_SPELLS = 100
MAX_MAPS = 255
MAX_MAP_ITEMS = 20
MAX_MAPX = 30
MAX_MAPY = 30
MAX_EMOTICONS = 10
Max_Classes = 3
MAX_LEVEL = 100
MAX_QUETES = 100
MAX_PETS = 10
MAX_METIER = 100
MAX_RECETTE = 200
End Sub
Sub ChargerCartes()
Dim FileName As String
Dim f As Long
Dim MapNum As Long
If HORS_LIGNE = 1 Then
    Call ClearMap
    For MapNum = 1 To MAX_MAPS
        FileName = App.Path & "\maps\map" & MapNum & ".fcc"
        If FileExiste("maps\map" & MapNum & ".fcc") Then
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , Map(MapNum)
            Close #f
        End If
    Next MapNum
    
    Call InitPano(Player(MyIndex).Map)
    Call InitNightAndFog(Player(MyIndex).Map)
    
    Dim i As Long
    For i = 1 To MAX_MAP_ITEMS
        MapItem(i) = SaveMapItem(i)
    Next i
        
    For i = 1 To MAX_MAP_NPCS
        MapNpc(i) = SaveMapNpc(i)
    Next i
        
    GettingMap = False
    Call Unload(frmmsg)
        
    ' Play music
    If OldMap > 0 Then
        If Trim$(Map(Player(MyIndex).Map).Music) = Trim$(Map(OldMap).Music) Then Exit Sub
        If Trim$(Map(Player(MyIndex).Map).Music) <> "Aucune" Then Call PlayMidi(Trim$(Map(Player(MyIndex).Map).Music)) Else Call StopMidi
    Else
        If Trim$(Map(Player(MyIndex).Map).Music) <> "Aucune" Then Call PlayMidi(Trim$(Map(Player(MyIndex).Map).Music)) Else Call StopMidi
    End If
Else
    For MapNum = 1 To MAX_MAPS
        FileName = App.Path & "\maps\map" & MapNum & ".fcc"
        If FileExiste("maps\map" & MapNum & ".fcc") Then
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , Map(MapNum)
            Close #f
        End If
    Next MapNum
End If
End Sub
Sub ChargerCarte(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long

If HORS_LIGNE = 1 Then
    FileName = App.Path & "\maps\map" & MapNum & ".fcc"
        
    If FileExiste("maps\map" & MapNum & ".fcc") Then
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Map(MapNum)
        Close #f
    End If
    
    Call InitPano(Player(MyIndex).Map)
    Call InitNightAndFog(Player(MyIndex).Map)

    Dim i As Long
    For i = 1 To MAX_MAP_ITEMS
        MapItem(i) = SaveMapItem(i)
    Next i
        
    For i = 1 To MAX_MAP_NPCS
        MapNpc(i) = SaveMapNpc(i)
    Next i
        
    GettingMap = False
    Call Unload(frmmsg)
        
    ' Play music
    If OldMap > 0 Then
        If Trim$(Map(Player(MyIndex).Map).Music) = Trim$(Map(OldMap).Music) Then Exit Sub
        If Trim$(Map(Player(MyIndex).Map).Music) <> "Aucune" Then Call PlayMidi(Trim$(Map(Player(MyIndex).Map).Music)) Else Call StopMidi
    Else
        If Trim$(Map(Player(MyIndex).Map).Music) <> "Aucune" Then Call PlayMidi(Trim$(Map(Player(MyIndex).Map).Music)) Else Call StopMidi
    End If
Else
    FileName = App.Path & "\maps\map" & MapNum & ".fcc"
    If FileExiste("maps\map" & MapNum & ".fcc") Then
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Map(MapNum)
        Close #f
    End If
End If
End Sub
Sub ChargerJoueure(Index As Long)
Dim FileName As String
Dim i As Long
Dim n As Long

    Call ClearPlayer(Index)
    
        ' General
        Player(Index).name = "Testeur"
        Player(Index).Class = 1
        Player(Index).sprite = 1
        Player(Index).Level = 1
        Player(Index).exp = 0
        Player(Index).Access = 5
        Player(Index).PK = 0
        Player(Index).guild = "Aucune Guild"
        Player(Index).Guildaccess = 0
        
        ' Vitals
        Player(Index).HP = 50 'a metre dan option choi
        Player(Index).MP = 50 'a metre dan option choi
        Player(Index).SP = 0 'a metre dan option choi
        
        ' Stats
        Player(Index).STR = 5 'a metre dan option choi
        Player(Index).def = 5 'a metre dan option choi
        Player(Index).speed = 5 'a metre dan option choi
        Player(Index).magi = 5 'a metre dan option choi
        Player(Index).POINTS = 0 'a metre dan option choi
        
        ' Worn equipment
        Player(Index).ArmorSlot = 0
        Player(Index).WeaponSlot = 0
        Player(Index).HelmetSlot = 0
        Player(Index).ShieldSlot = 0
        Player(Index).PetSlot = 0
        
        Player(Index).pet.Dir = DIR_DOWN
        Player(Index).pet.y = 1
        Player(Index).pet.y = 1
        
        ' Position
        Player(Index).Map = 1 'a metre dan option choi
        Player(Index).x = 1 'a metre dan option choi
        Player(Index).y = 1 'a metre dan option choi
        Player(Index).Dir = DIR_DOWN
        
        ' Check to make sure that they aren't on map 0, if so reset'm
        If Player(Index).Map = 0 Then Player(Index).Map = 1: Player(Index).x = 1: Player(Index).y = 1
        
        ' Inventory
        For n = 1 To MAX_INV
            Player(Index).inv(n).num = 0
            Player(Index).inv(n).value = 0
            Player(Index).inv(n).dur = 0
        Next n
        
        ' Spells
        For n = 1 To MAX_PLAYER_SPELLS
            Player(Index).Spell(n) = 0
        Next n
        
        'diver
        Player(Index).Access = 0
        Player(Index).MaxHp = 50 'metre dan option
        Player(Index).MaxMP = 50 'metre dan option
        Player(Index).MaxSP = 50 'metre dan option
End Sub
Sub ChargerLeTemps()
    GameWeather = 0
End Sub

Sub ChargerSorts()
Dim FileName As String
Dim i As Long
Dim f As Long
    For i = 1 To MAX_SPELLS
        If FileExiste("spells\spells" & i & ".fcg") Then
            FileName = App.Path & "\spells\spells" & i & ".fcg"
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , Spell(i)
            Close #f
        
            DoEvents
        End If
    Next i
End Sub
Sub ChargerQuetes()
Dim FileName As String
Dim i As Long
Dim f As Long
    For i = 1 To MAX_QUETES
        If FileExiste("quetes\quete" & i & ".fcq") Then
            FileName = App.Path & "\quetes\quete" & i & ".fcq"
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , quete(i)
            Close #f
        
            DoEvents
        End If
    Next i
End Sub
Sub ChargerMagasins()
Dim FileName As String
Dim i As Long, f As Long
    For i = 1 To MAX_SHOPS
        If FileExiste("shops\shop" & i & ".fcm") Then
            FileName = App.Path & "\shops\shop" & i & ".fcm"
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , Shop(i)
            Close #f
            
            DoEvents
        End If
    Next i
End Sub

Sub ChargerRecette()
Dim i As Long
Dim FileName As String
Dim f As Long
    For i = 1 To MAX_RECETTE
        Call ClearRecette(i)
    Next i
End Sub

Sub ChargerObjets(Index As Long)
Dim i As Long
Dim FileName As String
Dim f As Long
If HORS_LIGNE = 1 Then
    For i = 1 To MAX_ITEMS
        If FileExiste("Items\Item" & i & ".fco") Then
            FileName = App.Path & "\Items\Item" & i & ".fco"
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , Item(i)
            Close #f
            
            DoEvents
        Else
            Call ClearItem(i)
        End If
    Next i
    
Dim x As Long
Dim y As Long

Call ClearMapItems
i = 1
    For x = 1 To MAX_MAPX
        For y = 1 To MAX_MAPY
            If Map(Player(MyIndex).Map).tile(x, y).Type = TILE_TYPE_ITEM Then
                If Item(Map(Player(MyIndex).Map).tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY And Map(Player(MyIndex).Map).tile(x, y).Data2 <= 0 Then
                    MapItem(i).num = Map(Player(MyIndex).Map).tile(x, y).Data1
                    MapItem(i).value = 1
                    MapItem(i).dur = Item(Map(Player(MyIndex).Map).tile(x, y).Data1).Data1
                    MapItem(i).x = x
                    MapItem(i).y = y
                    i = i + 1
                Else
                    MapItem(i).num = Map(Player(MyIndex).Map).tile(x, y).Data1
                    MapItem(i).value = Map(Player(MyIndex).Map).tile(x, y).Data2
                    MapItem(i).dur = Item(Map(Player(MyIndex).Map).tile(x, y).Data1).Data1
                    MapItem(i).x = x
                    MapItem(i).y = y
                    i = i + 1
                End If
            End If
        Next y
    Next x
Else
    For i = 1 To MAX_ITEMS
        If FileExiste("Items\Item" & i & ".fco") Then
            FileName = App.Path & "\Items\Item" & i & ".fco"
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , Item(i)
            Close #f
            
            DoEvents
        End If
    Next i
End If
End Sub
Sub ChargerFleche()
Dim i As Long
    If Not FileExiste("Arrows.ini") Then
        For i = 1 To MAX_ARROWS
            DoEvents
            Call WriteINI("Arrow" & i, "ArrowName", vbNullString, App.Path & "\Arrows.ini")
            Call WriteINI("Arrow" & i, "ArrowPic", 0, App.Path & "\Arrows.ini")
            Call WriteINI("Arrow" & i, "ArrowRange", 0, App.Path & "\Arrows.ini")
        Next i
    End If

Dim FileName As String
    FileName = App.Path & "\Arrows.ini"
    
    For i = 1 To MAX_ARROWS
        Arrows(i).name = ReadINI("Arrow" & i, "ArrowName", FileName)
        Arrows(i).Pic = ReadINI("Arrow" & i, "ArrowPic", FileName)
        Arrows(i).Range = ReadINI("Arrow" & i, "ArrowRange", FileName)
        DoEvents
    Next i
End Sub

Sub ChargerEmots()
 Dim i As Long
    If Not FileExiste("emoticons.ini") Then
        For i = 0 To MAX_EMOTICONS
            DoEvents
            Call WriteINI("EMOTICONS", "Emoticon" & i, 0, App.Path & "\emoticons.ini")
            Call WriteINI("EMOTICONS", "EmoticonC" & i, vbNullString, App.Path & "\emoticons.ini")
        Next i
    End If

Dim FileName As String
    FileName = App.Path & "\emoticons.ini"
    
    For i = 0 To MAX_EMOTICONS
        Emoticons(i).Pic = Val(GetVar(FileName, "EMOTICONS", "Emoticon" & i))
        Emoticons(i).Command = GetVar(FileName, "EMOTICONS", "EmoticonC" & i)
        DoEvents
    Next i
End Sub

Sub ChargerExps()
Dim i As Long
    If Not FileExiste("experience.ini") Then
        For i = 1 To MAX_LEVEL
            DoEvents
            Call WriteINI("EXPERIENCE", "Exp" & i, i * 1500, App.Path & "\experience.ini")
        Next i
    End If
    
    For i = 1 To MAX_LEVEL
        DoEvents
        Experience(i) = Val(ReadINI("EXPERIENCE", "Exp" & i, App.Path & "\experience.ini"))
    Next i
End Sub

Sub ChargerPnjs()
Dim FileName As String
Dim i As Long
Dim z As Long
Dim f As Long
If HORS_LIGNE = 1 Then
    For i = 1 To MAX_NPCS
        If FileExiste("pnjs\npc" & i & ".fcp") Then
            FileName = App.Path & "\pnjs\npc" & i & ".fcp"
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , Npc(i)
            Close #f
            
            DoEvents
        End If
    Next i
    For i = 1 To MAX_MAP_NPCS
        MapNpc(i).num = Map(Player(MyIndex).Map).Npc(i)
        If MapNpc(i).num > 0 Then
            Randomize
            MapNpc(i).Dir = Int(4 * Rnd)
            MapNpc(i).HP = Npc(MapNpc(i).num).MaxHp
            MapNpc(i).MaxHp = Npc(MapNpc(i).num).MaxHp
            Randomize
            MapNpc(i).x = Map(Player(MyIndex).Map).Npcs(i).x
            MapNpc(i).y = Map(Player(MyIndex).Map).Npcs(i).y
        End If
    Next i
Else
    For i = 1 To MAX_NPCS
        If FileExiste("pnjs\npc" & i & ".fcp") Then
            FileName = App.Path & "\pnjs\npc" & i & ".fcp"
            f = FreeFile
            Open FileName For Binary As #f
                    Get #f, , Npc(i)
            Close #f
            
            DoEvents
        End If
    Next i
End If
End Sub

Sub ChargerClasses()
Dim FileName As String
Dim i As Long
        
    FileName = App.Path & "\Classes\info.ini"
    
    Max_Classes = Val(GetVar(FileName, "INFO", "MaxClasses"))
    
    ReDim Class(0 To Max_Classes) As ClassRec
    
    For i = 0 To Max_Classes
        FileName = App.Path & "\Classes\Class" & i & ".ini"
        Class(i).name = GetVar(FileName, "CLASS", "Name")
        Class(i).MaleSprite = Val(GetVar(FileName, "CLASS", "MaleSprite"))
        Class(i).FemaleSprite = Val(GetVar(FileName, "CLASS", "FemaleSprite"))
        Class(i).STR = Val(GetVar(FileName, "CLASS", "STR"))
        Class(i).def = Val(GetVar(FileName, "CLASS", "DEF"))
        Class(i).speed = Val(GetVar(FileName, "CLASS", "SPEED"))
        Class(i).magi = Val(GetVar(FileName, "CLASS", "MAGI"))
        Class(i).Locked = Val(GetVar(FileName, "CLASS", "Locked"))
        DoEvents
    Next i
End Sub

Sub CheckMapGetItem()
    If GetTickCount > Player(MyIndex).MapGetTimer + 250 And Trim$(MyText) = vbNullString Then
        Player(MyIndex).MapGetTimer = GetTickCount
        Call SendData("mapgetitem" & END_CHAR)
    End If
End Sub

Sub CheckAttack()
Dim AttackSpeed As Long
    If GetPlayerWeaponSlot(MyIndex) > 0 Then AttackSpeed = Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AttackSpeed Else AttackSpeed = 1000
    
    If ControlDown = True And Player(MyIndex).AttackTimer + AttackSpeed < GetTickCount And Player(MyIndex).Attacking = 0 Then
        Player(MyIndex).Attacking = 1
        Player(MyIndex).AttackTimer = GetTickCount
        Call SendData("attack" & END_CHAR)
    End If
End Sub

Sub CheckInput(ByVal KeyState As Byte, ByVal KeyCode As Integer, ByVal Shift As Integer)
    If Not GettingMap And Not InEditor Then
        If KeyState = 1 Then
            If KeyCode = vbKeyReturn Then Call CheckMapGetItem
            If KeyCode = vbKeyControl Then ControlDown = True
            If KeyCode = vbKeyUp Then
                DirUp = True
                DirDown = False
                DirLeft = False
                DirRight = False
            End If
            If KeyCode = vbKeyDown Then
                DirUp = False
                DirDown = True
                DirLeft = False
                DirRight = False
            End If
            If KeyCode = vbKeyLeft Then
                DirUp = False
                DirDown = False
                DirLeft = True
                DirRight = False
            End If
            If KeyCode = vbKeyRight Then
                DirUp = False
                DirDown = False
                DirLeft = False
                DirRight = True
            End If
            If KeyCode = vbKeyShift Then ShiftDown = True
        Else
            If KeyCode = vbKeyUp Then DirUp = False
            If KeyCode = vbKeyDown Then DirDown = False
            If KeyCode = vbKeyLeft Then DirLeft = False
            If KeyCode = vbKeyRight Then DirRight = False
            If KeyCode = vbKeyShift Then ShiftDown = False
            If KeyCode = vbKeyControl Then ControlDown = False
        End If
    End If
End Sub

Function IsTryingToMove() As Boolean
    If (DirUp = True) Or (DirDown = True) Or (DirLeft = True) Or (DirRight = True) Then IsTryingToMove = True Else IsTryingToMove = False
End Function

Function ObjetPos(ByVal x As Long, ByVal y As Long) As Boolean
Dim i As Long

ObjetPos = False

For i = 1 To MAX_MAP_ITEMS
    If MapItem(i).x = x And MapItem(i).y = y And MapItem(i).num > 0 Then ObjetPos = True
Next i

End Function

Function ObjetNumPos(ByVal x As Long, ByVal y As Long) As Long
Dim i As Long

ObjetNumPos = 0

For i = 1 To MAX_MAP_ITEMS
    If MapItem(i).x = x And MapItem(i).y = y And MapItem(i).num > 0 Then ObjetNumPos = MapItem(i).num
Next i

End Function

Function ObjetValPos(ByVal x As Long, ByVal y As Long) As Long
Dim i As Long

ObjetValPos = 0

For i = 1 To MAX_MAP_ITEMS
    If MapItem(i).x = x And MapItem(i).y = y And MapItem(i).num > 0 Then ObjetValPos = MapItem(i).value
Next i

End Function

Sub CaseChange(ByVal CX, ByVal CY)
Dim ONum As Long

If Val(ReadINI("CONFIG", "NomObjet", App.Path & "\Config\Account.ini")) = 0 Then frmMirage.ObjNm.Visible = False: Exit Sub

ONum = ObjetNumPos(CX, CY)

If ObjetPos(CX, CY) Then
    If Item(ONum).Type = ITEM_TYPE_CURRENCY Then frmMirage.OName.Caption = Trim$(Item(ONum).name) & "(" & ObjetValPos(CX, CY) & ")" Else frmMirage.OName.Caption = Trim$(Item(ONum).name) & "(1)"
    frmMirage.OName.ForeColor = Item(ONum).NCoul
    frmMirage.ObjNm.Left = PotX + 10
    frmMirage.ObjNm.Top = PotY - 30
    frmMirage.ObjNm.Width = frmMirage.OName.Width / Screen.TwipsPerPixelY + 240 / Screen.TwipsPerPixelY
    frmMirage.OName.Left = 120
    frmMirage.ObjNm.Visible = True
Else
    frmMirage.ObjNm.Visible = False
End If

End Sub

Function CanMove() As Boolean
Dim i As Long, d As Long
Dim x As Long, y As Long
Dim PX As Long, PY As Long
Dim Dire As Long

    CanMove = True
    
    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then CanMove = False: Exit Function
    
    ' Make sure they haven't just casted a spell
    If Player(MyIndex).CastedSpell = YES Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            Player(MyIndex).CastedSpell = NO
        Else
            CanMove = False
            Exit Function
        End If
    End If
           
    d = GetPlayerDir(MyIndex)
    PX = 0
    PY = 0
    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)
        Dire = DIR_UP
        If GetPlayerY(MyIndex) > 0 Then
            PX = 0
            PY = -1
        Else
            ' Check if they can warp to a new map
            If Map(Player(MyIndex).Map).Up > 0 Then Call SendPlayerRequestNewMap: GettingMap = True
            CanMove = False
            Exit Function
        End If
    End If
    If DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)
        Dire = DIR_DOWN
        If GetPlayerY(MyIndex) < MAX_MAPY Then
            PX = 0
            PY = 1
        Else
            ' Check if they can warp to a new map
            If Map(Player(MyIndex).Map).Down > 0 Then Call SendPlayerRequestNewMap: GettingMap = True
            CanMove = False
            Exit Function
        End If
    End If
    If DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)
        Dire = DIR_LEFT
        If GetPlayerX(MyIndex) > 0 Then
            PX = -1
            PY = 0
        Else
            ' Check if they can warp to a new map
            If Map(Player(MyIndex).Map).Left > 0 Then Call SendPlayerRequestNewMap: GettingMap = True
            CanMove = False
            Exit Function
        End If
    End If
    If DirRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)
        Dire = DIR_RIGHT
        If GetPlayerX(MyIndex) < MAX_MAPX Then
            PX = 1
            PY = 0
        Else
            ' Check if they can warp to a new map
            If Map(Player(MyIndex).Map).Right > 0 Then Call SendPlayerRequestNewMap: GettingMap = True
            CanMove = False
            Exit Function
        End If
    End If
    If PX = 0 And PY = 0 Then CanMove = False: Exit Function
        ' Check to see if the map tile is blocked or not
            If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCKED Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_SIGN Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_NIVEAUX Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_MONTURE Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_GUILDE Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_TOIT Then
                If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_MONTURE Then
                    If Player(MyIndex).ArmorSlot > 0 Then
                        If Item(Player(MyIndex).ArmorSlot).Type = ITEM_TYPE_MONTURE Then CanMove = False Else CanMove = True
                    Else
                        CanMove = True
                    End If
                ElseIf Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_NIVEAUX Then
                    If Player(MyIndex).Level < Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Data1 Then CanMove = False Else CanMove = True
                ElseIf Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_GUILDE Then
                    If Trim$(Player(MyIndex).guild) = Trim$(Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).String1) Then CanMove = True Else CanMove = False
                Else
                    CanMove = False
                End If
                
                ' Set the new direction if they weren't facing that direction
                If d <> Dire Then Call SendPlayerDir
                Exit Function
            End If
            
            If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_CBLOCK Then
                If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Data1 = Player(MyIndex).Class Then CanMove = True: Exit Function
                If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Data2 = Player(MyIndex).Class Then CanMove = True: Exit Function
                If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Data3 = Player(MyIndex).Class Then CanMove = True: Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> Dire Then Call SendPlayerDir
            End If
            
            If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_DIR Then
                If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Data1 = Player(MyIndex).Dir Then CanMove = True: Exit Function
                If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Data2 = Player(MyIndex).Dir Then CanMove = True: Exit Function
                If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Data3 = Player(MyIndex).Dir Then CanMove = True: Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> Dire Then Call SendPlayerDir
            End If
            
        ' verif atribut toit
        Call SuprTileToit(PY, PX)
                                                    
            ' Check to see if the key door is open or not
            If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_KEY Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_COFFRE Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_PORTE_CODE Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).DoorOpen = NO Then
                    CanMove = False
                    
                    ' Set the new direction if they weren't facing that direction
                    If d <> Dire Then Call SendPlayerDir
                Else
                    If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_COFFRE Then CanMove = False
                    Exit Function
                End If
            End If
                        
            ' Check to see if a player is already on that tile
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    If Player(i).Map = Player(MyIndex).Map Then
                        If Map(Player(MyIndex).Map).guildSoloView = 1 Then
                            If Map(Player(MyIndex).Map).traversable = 1 Then
                                If Player(MyIndex).guild = Player(i).guild Then
                                    If (GetPlayerX(i) = GetPlayerX(MyIndex) + PX) And (GetPlayerY(i) = GetPlayerY(MyIndex) + PY) Then
                                        CanMove = False
                                    
                                        ' Set the new direction if they weren't facing that direction
                                        If d <> Dire Then Call SendPlayerDir
                                        Exit Function
                                    End If
                                End If
                            End If
                        Else
                            If Map(Player(MyIndex).Map).traversable = 1 Then
                                If (GetPlayerX(i) = GetPlayerX(MyIndex) + PX) And (GetPlayerY(i) = GetPlayerY(MyIndex) + PY) Then
                                    CanMove = False
                                
                                    ' Set the new direction if they weren't facing that direction
                                    If d <> Dire Then Call SendPlayerDir
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
        
            ' Check to see if a npc is already on that tile
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).num > 0 Then
                    If (MapNpc(i).x = GetPlayerX(MyIndex) + PX) And (MapNpc(i).y = GetPlayerY(MyIndex) + PY) And Npc(MapNpc(i).num).vol = 0 Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> Dire Then Call SendPlayerDir
                        Exit Function
                    End If
                End If
            Next i
End Function

Sub SuprTileToit(ByVal Dy As Long, ByVal Dx As Long)
' verif atribut toit
On Error Resume Next
                
            If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, GetPlayerY(MyIndex) + Dy).Type <> TILE_TYPE_WALKABLE And Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_TOIT Then
            If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, GetPlayerY(MyIndex) + Dy).Fringe > 0 Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, GetPlayerY(MyIndex) + Dy).Fringe2 > 0 Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, GetPlayerY(MyIndex) + Dy).F2Anim > 0 Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, GetPlayerY(MyIndex) + Dy).FAnim > 0 Then
                Dim MX As Long
                Dim MY As Long
                Dim er As Long
                Dim i As Long
            
                
                If Not InEditor And Not InToit Then
                
                For er = Player(MyIndex).y To MAX_MAPY
                If er < MAX_MAPY Then
                If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    For i = Player(MyIndex).x To MAX_MAPX
                    If i < MAX_MAPX Then
                        If Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).F3Anim = 0
                        Else
                        If Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Type = TILE_TYPE_KEY Then
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).F3Anim = 0
                        End If
                    Else
                        If Map(Player(MyIndex).Map).tile(i, er + Dy).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(i, er + Dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(Player(MyIndex).Map).tile(i, er + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).F3Anim = 0
                        Else
                        If Map(Player(MyIndex).Map).tile(i, er + Dy).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(i, er + Dy).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(i, er + Dy).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(i, er + Dy).Type = TILE_TYPE_KEY Then
                            Map(Player(MyIndex).Map).tile(i, er + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(Player(MyIndex).Map).tile(i, er + Dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(Player(MyIndex).Map).tile(i, er + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).F3Anim = 0
                        End If
                    End If
                    Next i
                        MX = Player(MyIndex).x
                    For i = 0 To Player(MyIndex).x
                        If Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).F3Anim = 0
                        Else
                        If Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Type = TILE_TYPE_KEY Then
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Type = TILE_TYPE_BLOCKED Then Exit For
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).F3Anim = 0
                        End If
                        MX = MX - 1
                    Next i
                Else
                If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Type = TILE_TYPE_KEY Then
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Fringe = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Fringe2 = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Fringe3 = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).FAnim = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).F2Anim = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).F3Anim = 0
                        Exit For
                    End If
                    If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Fringe = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Fringe2 = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Fringe3 = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).FAnim = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).F2Anim = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).F3Anim = 0
                End If
                Else
                If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).Type = TILE_TYPE_BLOCK_TOIT Then
                    For i = Player(MyIndex).x To MAX_MAPX
                    If i < MAX_MAPX Then
                        If Map(Player(MyIndex).Map).tile(i + Dx, er).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(i + Dx, er).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(Player(MyIndex).Map).tile(i + Dx, er).Fringe = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er).FAnim = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er).F3Anim = 0
                        Else
                        If Map(Player(MyIndex).Map).tile(i + Dx, er).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(i + Dx, er).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(i + Dx, er).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(i + Dx, er).Type = TILE_TYPE_KEY Then
                            Map(Player(MyIndex).Map).tile(i + Dx, er).Fringe = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er).FAnim = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er).F3Anim = 0
                            Exit For
                        End If
                            If Map(Player(MyIndex).Map).tile(i + Dx, er).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(Player(MyIndex).Map).tile(i + Dx, er).Fringe = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er).FAnim = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er).F3Anim = 0
                        End If
                    Else
                        If Map(Player(MyIndex).Map).tile(i, er).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(i, er).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(Player(MyIndex).Map).tile(i, er).Fringe = 0
                            Map(Player(MyIndex).Map).tile(i, er).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(i, er).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(i, er).FAnim = 0
                            Map(Player(MyIndex).Map).tile(i, er).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(i, er).F3Anim = 0
                        Else
                        If Map(Player(MyIndex).Map).tile(i, er).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(i, er).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(i, er).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(i, er).Type = TILE_TYPE_KEY Then
                            Map(Player(MyIndex).Map).tile(i, er).Fringe = 0
                            Map(Player(MyIndex).Map).tile(i, er).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(i, er).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(i, er).FAnim = 0
                            Map(Player(MyIndex).Map).tile(i, er).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(i, er).F3Anim = 0
                            Exit For
                        End If
                            If Map(Player(MyIndex).Map).tile(i, er).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(Player(MyIndex).Map).tile(i, er).Fringe = 0
                            Map(Player(MyIndex).Map).tile(i, er).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(i, er).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(i, er).FAnim = 0
                            Map(Player(MyIndex).Map).tile(i, er).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(i, er).F3Anim = 0
                        End If
                    End If
                    Next i
                        MX = Player(MyIndex).x
                    For i = 0 To Player(MyIndex).x
                        If Map(Player(MyIndex).Map).tile(MX + Dx, er).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(MX + Dx, er).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(Player(MyIndex).Map).tile(MX + Dx, er).Fringe = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er).FAnim = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er).F3Anim = 0
                        Else
                        If Map(Player(MyIndex).Map).tile(MX + Dx, er).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(MX + Dx, er).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(MX + Dx, er).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(MX + Dx, er).Type = TILE_TYPE_KEY Then
                            Map(Player(MyIndex).Map).tile(MX + Dx, er).Fringe = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er).FAnim = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er).F3Anim = 0
                            Exit For
                        End If
                            If Map(Player(MyIndex).Map).tile(MX + Dx, er).Type = TILE_TYPE_BLOCKED Then Exit For
                            Map(Player(MyIndex).Map).tile(MX + Dx, er).Fringe = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er).FAnim = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er).F3Anim = 0
                        End If
                        MX = MX - 1
                    Next i
                Else
                If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).Type = TILE_TYPE_KEY Then
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).Fringe = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).Fringe2 = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).Fringe3 = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).FAnim = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).F2Anim = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).F3Anim = 0
                        Exit For
                    End If
                    If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).Fringe = 0
                    Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).Fringe2 = 0
                    Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).Fringe3 = 0
                    Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).FAnim = 0
                    Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).F2Anim = 0
                    Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er).F3Anim = 0
                End If
                End If
                Next er
                
                er = Player(MyIndex).y
                For MY = 0 To Player(MyIndex).y
                If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    For i = Player(MyIndex).x To MAX_MAPX
                    If i < MAX_MAPX Then
                        If Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).F3Anim = 0
                        Else
                        If Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Type = TILE_TYPE_KEY Then
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(i + Dx, er + Dy).F3Anim = 0
                        End If
                        Else
                        If Map(Player(MyIndex).Map).tile(i, er + Dy).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(i, er + Dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(Player(MyIndex).Map).tile(i, er + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).F3Anim = 0
                        Else
                        If Map(Player(MyIndex).Map).tile(i, er + Dy).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(i, er + Dy).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(i, er + Dy).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(i, er + Dy).Type = TILE_TYPE_KEY Then
                            Map(Player(MyIndex).Map).tile(i, er + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(Player(MyIndex).Map).tile(i, er + Dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(Player(MyIndex).Map).tile(i, er + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(i, er + Dy).F3Anim = 0
                        End If
                    End If
                    Next i
                        MX = Player(MyIndex).x
                    For i = 0 To Player(MyIndex).x
                        If Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).F3Anim = 0
                        Else
                        If Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Type = TILE_TYPE_KEY Then
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(MX + Dx, er + Dy).F3Anim = 0
                        End If
                        MX = MX - 1
                    Next i
                Else
                    If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Type = TILE_TYPE_KEY Then
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Fringe = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Fringe2 = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Fringe3 = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).FAnim = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).F2Anim = 0
                        Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).F3Anim = 0
                        Exit For
                    End If
                    If Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Fringe = 0
                    Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Fringe2 = 0
                    Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).Fringe3 = 0
                    Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).FAnim = 0
                    Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).F2Anim = 0
                    Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex) + Dx, er + Dy).F3Anim = 0
                End If
                er = er - 1
                Next MY
                
                For er = Player(MyIndex).x To MAX_MAPX
                If er < MAX_MAPX Then
                If Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    For i = Player(MyIndex).y To MAX_MAPY
                    If i < MAX_MAPY Then
                        If Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).F3Anim = 0
                        Else
                        If Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Type = TILE_TYPE_KEY Then
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).F3Anim = 0
                        End If
                    Else
                    If Map(Player(MyIndex).Map).tile(er + Dx, i).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(er + Dx, i).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(Player(MyIndex).Map).tile(er + Dx, i).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).F3Anim = 0
                        Else
                        If Map(Player(MyIndex).Map).tile(er + Dx, i).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(er + Dx, i).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(er + Dx, i).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(er + Dx, i).Type = TILE_TYPE_KEY Then
                            Map(Player(MyIndex).Map).tile(er + Dx, i).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).F3Anim = 0
                            Exit For
                        End If
                            If Map(Player(MyIndex).Map).tile(er + Dx, i).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(Player(MyIndex).Map).tile(er + Dx, i).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).F3Anim = 0
                        End If
                    End If
                    Next i
                        MY = Player(MyIndex).y
                    For i = 0 To Player(MyIndex).y
                        If Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).F3Anim = 0
                        Else
                        If Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Type = TILE_TYPE_KEY Then
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).F3Anim = 0
                        End If
                        MY = MY - 1
                    Next i
                Else
                    If Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_KEY Then
                        Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Fringe = 0
                        Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Fringe2 = 0
                        Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Fringe3 = 0
                        Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).FAnim = 0
                        Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).F2Anim = 0
                        Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).F3Anim = 0
                        Exit For
                    End If
                    If Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Fringe = 0
                    Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Fringe2 = 0
                    Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Fringe3 = 0
                    Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).FAnim = 0
                    Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).F2Anim = 0
                    Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).F3Anim = 0
                End If
                Else
                If Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    For i = Player(MyIndex).y To MAX_MAPY
                    If i < MAX_MAPY Then
                        If Map(Player(MyIndex).Map).tile(er, i + Dy).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(er, i + Dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(Player(MyIndex).Map).tile(er, i + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er, i + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er, i + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er, i + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er, i + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er, i + Dy).F3Anim = 0
                        Else
                        If Map(Player(MyIndex).Map).tile(er, i + Dy).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(er, i + Dy).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(er, i + Dy).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(er, i + Dy).Type = TILE_TYPE_KEY Then
                            Map(Player(MyIndex).Map).tile(er, i + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er, i + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er, i + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er, i + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er, i + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er, i + Dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(Player(MyIndex).Map).tile(er, i + Dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(Player(MyIndex).Map).tile(er, i + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er, i + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er, i + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er, i + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er, i + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er, i + Dy).F3Anim = 0
                        End If
                    Else
                    If Map(Player(MyIndex).Map).tile(er, i).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(er, i).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(Player(MyIndex).Map).tile(er, i).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er, i).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er, i).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er, i).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er, i).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er, i).F3Anim = 0
                        Else
                        If Map(Player(MyIndex).Map).tile(er, i).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(er, i).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(er, i).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(er, i).Type = TILE_TYPE_KEY Then
                            Map(Player(MyIndex).Map).tile(er, i).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er, i).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er, i).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er, i).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er, i).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er, i).F3Anim = 0
                            Exit For
                        End If
                            If Map(Player(MyIndex).Map).tile(er, i).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(Player(MyIndex).Map).tile(er, i).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er, i).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er, i).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er, i).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er, i).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er, i).F3Anim = 0
                        End If
                    End If
                    Next i
                        MY = Player(MyIndex).y
                    For i = 0 To Player(MyIndex).y
                        If Map(Player(MyIndex).Map).tile(er, MY + Dy).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(er, MY + Dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(Player(MyIndex).Map).tile(er, MY + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er, MY + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er, MY + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er, MY + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er, MY + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er, MY + Dy).F3Anim = 0
                        Else
                        If Map(Player(MyIndex).Map).tile(er, MY + Dy).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(er, MY + Dy).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(er, MY + Dy).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(er, MY + Dy).Type = TILE_TYPE_KEY Then
                            Map(Player(MyIndex).Map).tile(er, MY + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er, MY + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er, MY + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er, MY + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er, MY + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er, MY + Dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(Player(MyIndex).Map).tile(er, MY + Dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(Player(MyIndex).Map).tile(er, MY + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er, MY + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er, MY + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er, MY + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er, MY + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er, MY + Dy).F3Anim = 0
                        End If
                        MY = MY - 1
                    Next i
                Else
                    If Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_KEY Then
                        Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).Fringe = 0
                        Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).Fringe2 = 0
                        Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).Fringe3 = 0
                        Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).FAnim = 0
                        Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).F2Anim = 0
                        Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).F3Anim = 0
                        Exit For
                    End If
                    If Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).Fringe = 0
                    Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).Fringe2 = 0
                    Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).Fringe3 = 0
                    Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).FAnim = 0
                    Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).F2Anim = 0
                    Map(Player(MyIndex).Map).tile(er, GetPlayerY(MyIndex) + Dy).F3Anim = 0
                End If
                End If
                Next er
                
                er = Player(MyIndex).x
                For MX = 0 To Player(MyIndex).x
                If Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    For i = Player(MyIndex).y To MAX_MAPY
                    If i < MAX_MAPY Then
                        If Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).F3Anim = 0
                        Else
                        If Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Type = TILE_TYPE_KEY Then
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i + Dy).F3Anim = 0
                        End If
                    Else
                        If Map(Player(MyIndex).Map).tile(er + Dx, i).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(er + Dx, i).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(Player(MyIndex).Map).tile(er + Dx, i).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).F3Anim = 0
                        Else
                        If Map(Player(MyIndex).Map).tile(er + Dx, i).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(er + Dx, i).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(er + Dx, i).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(er + Dx, i).Type = TILE_TYPE_KEY Then
                            Map(Player(MyIndex).Map).tile(er + Dx, i).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).F3Anim = 0
                            Exit For
                        End If
                            If Map(Player(MyIndex).Map).tile(er + Dx, i).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(Player(MyIndex).Map).tile(er + Dx, i).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, i).F3Anim = 0
                        End If
                    End If
                    Next i
                        MY = Player(MyIndex).y
                    For i = 0 To Player(MyIndex).y
                        If Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Type = TILE_TYPE_TOIT Or Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).F3Anim = 0
                        Else
                        If Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Type = TILE_TYPE_KEY Then
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Fringe = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Fringe2 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).Fringe3 = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).FAnim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).F2Anim = 0
                            Map(Player(MyIndex).Map).tile(er + Dx, MY + Dy).F3Anim = 0
                        End If
                        MY = MY - 1
                    Next i
                Else
                    If Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_DOOR Or Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_PORTE_CODE Or Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_WARP Or Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_KEY Then
                        Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Fringe = 0
                        Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Fringe2 = 0
                        Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Fringe3 = 0
                        Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).FAnim = 0
                        Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).F2Anim = 0
                        Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).F3Anim = 0
                        Exit For
                    End If
                    If Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Fringe = 0
                    Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Fringe2 = 0
                    Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).Fringe3 = 0
                    Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).FAnim = 0
                    Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).F2Anim = 0
                    Map(Player(MyIndex).Map).tile(er + Dx, GetPlayerY(MyIndex) + Dy).F3Anim = 0
                End If
                er = er - 1
                Next MX
                InToit = True
                Else
                If InToit And Not InEditor Then Call LoadMap(Player(MyIndex).Map)
                InToit = False
                End If
            End If
            Else
                If InToit And Not InEditor Then Call LoadMap(Player(MyIndex).Map)
                InToit = False
            End If
End Sub

Sub CheckMovement()
    If Not GettingMap And IsTryingToMove And CanMove Then
        ' Check if player has the shift key down for running
        If ShiftDown Then Player(MyIndex).Moving = MOVING_RUNNING Else Player(MyIndex).Moving = MOVING_WALKING
        
        If Player(MyIndex).PetSlot <> 0 Then
            Select Case Player(MyIndex).pet.Dir
                Case DIR_UP
                    If Player(MyIndex).pet.Anim = 0 Then Player(MyIndex).pet.Anim = 2 Else Player(MyIndex).pet.Anim = 0
            
                Case DIR_DOWN
                    If Player(MyIndex).pet.Anim = 0 Then Player(MyIndex).pet.Anim = 2 Else Player(MyIndex).pet.Anim = 0
            
                Case DIR_LEFT
                    If Player(MyIndex).pet.Anim = 0 Then Player(MyIndex).pet.Anim = 2 Else Player(MyIndex).pet.Anim = 0
            
                Case DIR_RIGHT
                    If Player(MyIndex).pet.Anim = 0 Then Player(MyIndex).pet.Anim = 2 Else Player(MyIndex).pet.Anim = 0
            End Select
        End If
        
        Select Case GetPlayerDir(MyIndex)
            Case DIR_UP
                Call SendPlayerMove
                Player(MyIndex).YOffset = PIC_Y
                Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                If Player(MyIndex).Anim = 0 Then Player(MyIndex).Anim = 2 Else Player(MyIndex).Anim = 0
        
            Case DIR_DOWN
                Call SendPlayerMove
                Player(MyIndex).YOffset = PIC_Y * -1
                Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                If Player(MyIndex).Anim = 0 Then Player(MyIndex).Anim = 2 Else Player(MyIndex).Anim = 0
        
            Case DIR_LEFT
                Call SendPlayerMove
                Player(MyIndex).XOffset = PIC_X
                Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                If Player(MyIndex).Anim = 0 Then Player(MyIndex).Anim = 2 Else Player(MyIndex).Anim = 0
        
            Case DIR_RIGHT
                Call SendPlayerMove
                Player(MyIndex).XOffset = PIC_X * -1
                Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
                If Player(MyIndex).Anim = 0 Then Player(MyIndex).Anim = 2 Else Player(MyIndex).Anim = 0
        End Select
    
        ' Gotta check :)
        If Not InEditor And Map(Player(MyIndex).Map).tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then GettingMap = True
    End If
End Sub

Function FindPlayer(ByVal name As String) As Long
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(name)))) = UCase$(Trim$(name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    FindPlayer = 0
End Function

Function FindOpenInvSlot(ByVal ItemNum As Long) As Long
Dim i As Long
    
    FindOpenInvSlot = 0
    
    ' Check for subscript out of range
    If IsPlaying(MyIndex) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function
    
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        ' If currency then check to see if they already have an guildSoloView of the item and add it to that
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(MyIndex, i) = ItemNum Then FindOpenInvSlot = i: Exit Function
        Next i
    End If
    
    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(MyIndex, i) <= 0 Then FindOpenInvSlot = i: Exit Function
    Next i
End Function

Public Sub EditorInit()
    Call EcrireEtat("Initialisation de l'éditeur")
    InEditor = True
    EditorSet = 0
    Call EcrireEtat("Initialisation de l'éditeur : Affichage des tiles")
    Call AffTilesPic(EditorSet, frmMirage.scrlPicture.value * PIC_Y)
    frmMirage.picBackSelect.Refresh
    'frmMirage.picBackSelect.Picture = LoadPNG(App.Path + "\GFX\tiles0.png")
    frmMirage.scrlPicture.Max = Int((DDSD_Tile(EditorSet).lHeight - frmMirage.picBackSelect.Height) \ PIC_Y)
    frmMirage.picBack.Width = frmMirage.picBackSelect.Width
    Call EcrireEtat("Initialisation de l'éditeur : Terminer")
End Sub

Public Sub EditorMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1, y1 As Long
Dim x2 As Long, y2 As Long, PicX As Long
On Error GoTo er:

If InMouvEditor Then
    If Map(Player(MyIndex).Map).tile((x \ 32 / VZoom * 3), (y \ 32 / VZoom * 3)).Type <> TILE_TYPE_WALKABLE And Map(Player(MyIndex).Map).tile((x \ 32 / VZoom * 3), (y \ 32 / VZoom * 3)).Type <> TILE_TYPE_TOIT Then If frmCpnjmouv.imob.value = Unchecked Then Call MsgBox("Le pnj ne peut pas marcher ou apparaître sur cette case veuillez en sélectionner une autre"): Exit Sub
    frmCpnjmouv.x(cordo).Text = (x \ 32 / VZoom * 3)
    frmCpnjmouv.y(cordo).Text = (y \ 32 / VZoom * 3)
    frmCpnjmouv.SetFocus
    Exit Sub
End If

If InQuetesEditor Then
    If Map(Player(MyIndex).Map).tile((x \ 32 / VZoom * 3), (y \ 32 / VZoom * 3)).Type <> TILE_TYPE_WALKABLE Then Call MsgBox("Le joueur ne peut pas marcher ou apparaître sur cette case veuillez en sélectionner une autre"): Exit Sub
    frmEditeurQuetes.carted.Text = Player(MyIndex).Map
    frmEditeurQuetes.xd.Text = (x \ 32 / VZoom * 3)
    frmEditeurQuetes.yd.Text = (y \ 32 / VZoom * 3)
    frmEditeurQuetes.Show
    Exit Sub
End If

If InDefTel Then
    frmMapWarp.txtMap.Text = Player(MyIndex).Map
    frmMapWarp.scrlX.value = (x \ 32 / VZoom * 3)
    frmMapWarp.scrlY.value = (y \ 32 / VZoom * 3)
    frmMapWarp.lblX.Caption = (x \ 32 / VZoom * 3)
    frmMapWarp.lblY.Caption = (y \ 32 / VZoom * 3)
    frmMapWarp.Show
    Exit Sub
End If

If InDefKey Then
    frmKeyOpen.scrlX.value = (x \ 32 / VZoom * 3)
    frmKeyOpen.scrlY.value = (y \ 32 / VZoom * 3)
    frmKeyOpen.lblX.Caption = (x \ 32 / VZoom * 3)
    frmKeyOpen.lblY.Caption = (y \ 32 / VZoom * 3)
    frmKeyOpen.Show
    Exit Sub
End If

    If InEditor Then
        x1 = (x \ PIC_X / VZoom * 3)
        y1 = (y \ PIC_Y / VZoom * 3)
        
        If x1 < 0 Or x1 > MAX_MAPX Or y1 < 0 Or y1 > MAX_MAPY Then Exit Sub
        
        If frmMirage.MousePointer = 2 Then
            If frmMirage.tp(1).Checked Then
                With Map(Player(MyIndex).Map).tile(x1, y1)
                    If frmMirage.Toolbar1.buttons(5).value = tbrPressed Then
                        PicX = .Ground
                        EditorSet = .GroundSet
                    ElseIf frmMirage.Toolbar1.buttons(6).value = tbrPressed Then
                        PicX = .Mask
                        EditorSet = .MaskSet
                    ElseIf frmMirage.Toolbar1.buttons(13).value = tbrPressed Then
                        PicX = .Anim
                        EditorSet = .AnimSet
                    ElseIf frmMirage.Toolbar1.buttons(7).value = tbrPressed Then
                        PicX = .Mask2
                        EditorSet = .Mask2Set
                    ElseIf frmMirage.Toolbar1.buttons(14).value = tbrPressed Then
                        PicX = .M2Anim
                        EditorSet = .M2AnimSet
                    ElseIf frmMirage.Toolbar1.buttons(8).value = tbrPressed Then '<--
                        PicX = .Mask3
                        EditorSet = .Mask3Set
                    ElseIf frmMirage.Toolbar1.buttons(15).value = tbrPressed Then '<--
                        PicX = .M3Anim
                        EditorSet = .M3AnimSet
                    ElseIf frmMirage.Toolbar1.buttons(9).value = tbrPressed Then
                        PicX = .Fringe
                        EditorSet = .FringeSet
                   ElseIf frmMirage.Toolbar1.buttons(16).value = tbrPressed Then
                        PicX = .FAnim
                        EditorSet = .FAnimSet
                    ElseIf frmMirage.Toolbar1.buttons(10).value = tbrPressed Then
                        PicX = .Fringe2
                        EditorSet = .Fringe2Set
                    ElseIf frmMirage.Toolbar1.buttons(17).value = tbrPressed Then
                        PicX = .F2Anim
                        EditorSet = .F2AnimSet
                    ElseIf frmMirage.Toolbar1.buttons(11).value = tbrPressed Then '<--
                        PicX = .Fringe3
                        EditorSet = .Fringe3Set
                    ElseIf frmMirage.Toolbar1.buttons(18).value = tbrPressed Then '<--
                        PicX = .F3Anim
                        EditorSet = .F3AnimSet
                    End If
                    
                    EditorTileY = (PicX \ TilesInSheets)
                    EditorTileX = (PicX - (PicX \ TilesInSheets) * TilesInSheets)
                    frmMirage.shpSelected.Top = Int(EditorTileY * PIC_Y)
                    frmMirage.shpSelected.Left = Int(EditorTileX * PIC_Y)
                    frmMirage.shpSelected.Height = PIC_Y
                    frmMirage.shpSelected.Width = PIC_X
                    If frmMirage.Tiles(EditorSet).Checked = False Then
                        frmMirage.Tiles(EditorSet).Checked = True
                        Call AffTilesPic(EditorSet, frmMirage.scrlPicture.value * PIC_Y)
                        frmMirage.scrlPicture.Max = ((DDSD_Tile(EditorSet).lHeight - frmMirage.picBackSelect.Height) \ PIC_Y)
                        frmMirage.HScroll1.Max = frmMirage.picBackSelect.Width / 32
                        frmMirage.picBack.Width = frmMirage.picBackSelect.Width
                        frmMirage.tilescmb.ListIndex = EditorSet
                    End If
                    
                    Dim i As Byte
                    For i = 0 To ExtraSheets
                        If i <> EditorSet Then frmMirage.Tiles(i).Checked = False
                    Next i
                    If frmMirage.previsu.Checked And InEditor And frmMirage.tp(1).Checked And frmMirage.MousePointer <> 99 Then Call PreVisua
                End With
            ElseIf frmMirage.tp(3).Checked Then
                EditorTileY = (Map(Player(MyIndex).Map).tile(x1, y1).Light \ TilesInSheets)
                EditorTileX = (Map(Player(MyIndex).Map).tile(x1, y1).Light - (Map(Player(MyIndex).Map).tile(x1, y1).Light \ TilesInSheets) * TilesInSheets)
                frmMirage.shpSelected.Top = Int(EditorTileY * PIC_Y)
                frmMirage.shpSelected.Left = Int(EditorTileX * PIC_Y)
                frmMirage.shpSelected.Height = PIC_Y
                frmMirage.shpSelected.Width = PIC_X
            ElseIf frmMirage.tp(2).Checked Then
                With Map(Player(MyIndex).Map).tile(x1, y1)
                    If .Type = TILE_TYPE_BLOCKED Then frmMirage.optBlocked.value = True
                    If .Type = TILE_TYPE_WARP Then
                        EditorWarpMap = .Data1
                        EditorWarpX = .Data2
                        EditorWarpY = .Data3
                        frmMirage.optWarp.value = True
                    End If
                    If .Type = TILE_TYPE_HEAL Then frmMirage.optHeal.value = True
                    If .Type = TILE_TYPE_KILL Then frmMirage.optKill.value = True
                    If .Type = TILE_TYPE_ITEM Then
                        ItemEditorNum = .Data1
                        ItemEditorValue = .Data2
                        frmMirage.optItem.value = True
                    End If
                    If .Type = TILE_TYPE_NPCAVOID Then frmMirage.optNpcAvoid.value = True
                    If .Type = TILE_TYPE_KEY Then
                        KeyEditorNum = .Data1
                        KeyEditorTake = .Data2
                        frmMirage.optKey.value = True
                    ElseIf .Type = TILE_TYPE_KEYOPEN Then
                        KeyOpenEditorX = .Data1
                        KeyOpenEditorY = .Data2
                        KeyOpenEditorMsg = .String1
                        frmMirage.optKeyOpen.value = True
                    ElseIf .Type = TILE_TYPE_SHOP Then
                        EditorShopNum = .Data1
                        frmMirage.optShop.value = True
                    ElseIf .Type = TILE_TYPE_CBLOCK Then
                        EditorItemNum1 = .Data1
                        EditorItemNum2 = .Data2
                        EditorItemNum3 = .Data3
                        frmMirage.optCBlock.value = True
                    ElseIf .Type = TILE_TYPE_ARENA Then
                        Arena1 = .Data1
                        Arena2 = .Data2
                        Arena3 = .Data3
                        frmMirage.optArena.value = True
                    ElseIf .Type = TILE_TYPE_SOUND Then
                        SoundFileName = .String1
                        frmMirage.optSound.value = True
                    ElseIf .Type = TILE_TYPE_SPRITE_CHANGE Then
                        SpritePic = .Data1
                        SpriteItem = .Data2
                        SpritePrice = .Data3
                        frmMirage.optSprite.value = True
                    ElseIf .Type = TILE_TYPE_SIGN Then
                        SignLine1 = .String1
                        frmMirage.optSign.value = True
                    End If
                    If .Type = TILE_TYPE_DOOR Then frmMirage.optDoor.value = True
                    If .Type = TILE_TYPE_NOTICE Then
                        NoticeTitle = .String1
                        NoticeText = .String2
                        NoticeSound = .String3
                        frmMirage.optNotice.value = True
                    ElseIf .Type = TILE_TYPE_CLASS_CHANGE Then
                        ClassChange = .Data1
                        ClassChangeReq = .Data2
                        frmMirage.optClassChange.value = True
                    ElseIf .Type = TILE_TYPE_SCRIPTED Then
                        ScriptNum = .Data1
                        frmMirage.optScripted.value = True
                    ElseIf .Type = TILE_TYPE_CRAFT Then
                        ScriptNum = .Data1
                        frmMirage.OptCraft.value = True
                    ElseIf .Type = TILE_TYPE_METIER Then
                        ScriptNum = .Data1
                        frmMirage.OptMetier.value = True
                    ElseIf .Type = TILE_TYPE_BANK Then
                        frmMirage.OptBank.value = True
                        bankmsg = .String1
                    ElseIf .Type = TILE_TYPE_COFFRE Then
                        frmMirage.optcoffre.value = True
                        CodeCoffre = .String1
                        CleCoffreNum = .Data1
                        CleCoffreSupr = .Data2
                        ObjCoffreNum = .Data3
                    ElseIf .Type = TILE_TYPE_PORTE_CODE Then
                        frmMirage.optportecode.value = True
                        CodePorte = .String1
                    End If
                    If .Type = TILE_TYPE_BLOCK_MONTURE Then frmMirage.optBmont.value = True
                    If .Type = TILE_TYPE_BLOCK_NIVEAUX Then
                        frmMirage.optBniv.value = True
                        NivMin = .Data1
                    End If
                    If .Type = TILE_TYPE_TOIT Then frmMirage.opttoit.value = True
                    If .Type = TILE_TYPE_BLOCK_GUILDE Then
                        frmMirage.optBguilde.value = True
                        NomGuilde = .String1
                    End If
                    If .Type = TILE_TYPE_BLOCK_TOIT Then frmMirage.optbtoit.value = True
                    If .Type = TILE_TYPE_BLOCK_DIR Then
                        frmMirage.optBDir.value = True
                        AccptDir1 = .Data1
                        AccptDir2 = .Data2
                        AccptDir3 = .Data3
                    End If
                End With
            End If
            frmMirage.MousePointer = 1
            frmMirage.Toolbar1.buttons(32).value = tbrUnpressed
        Else
            If (Button = 1) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
                If frmMirage.shpSelected.Height <= PIC_Y And frmMirage.shpSelected.Width <= PIC_X Then
                    If frmMirage.tp(1).Checked Then
                        With Map(Player(MyIndex).Map).tile(x1, y1)
                            If frmMirage.Toolbar1.buttons(5).value = tbrPressed Then
                                .Ground = EditorTileY * TilesInSheets + EditorTileX
                                .GroundSet = EditorSet
                            ElseIf frmMirage.Toolbar1.buttons(6).value = tbrPressed Then
                                .Mask = EditorTileY * TilesInSheets + EditorTileX
                                .MaskSet = EditorSet
                            ElseIf frmMirage.Toolbar1.buttons(13).value = tbrPressed Then
                                .Anim = EditorTileY * TilesInSheets + EditorTileX
                                .AnimSet = EditorSet
                            ElseIf frmMirage.Toolbar1.buttons(7).value = tbrPressed Then
                                .Mask2 = EditorTileY * TilesInSheets + EditorTileX
                                .Mask2Set = EditorSet
                            ElseIf frmMirage.Toolbar1.buttons(14).value = tbrPressed Then
                                .M2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .M2AnimSet = EditorSet
                            ElseIf frmMirage.Toolbar1.buttons(8).value = tbrPressed Then '<--
                                .Mask3 = EditorTileY * TilesInSheets + EditorTileX
                                .Mask3Set = EditorSet
                            ElseIf frmMirage.Toolbar1.buttons(15).value = tbrPressed Then '<--
                                .M3Anim = EditorTileY * TilesInSheets + EditorTileX
                                .M3AnimSet = EditorSet
                            ElseIf frmMirage.Toolbar1.buttons(9).value = tbrPressed Then
                                .Fringe = EditorTileY * TilesInSheets + EditorTileX
                                .FringeSet = EditorSet
                            ElseIf frmMirage.Toolbar1.buttons(16).value = tbrPressed Then
                                .FAnim = EditorTileY * TilesInSheets + EditorTileX
                                .FAnimSet = EditorSet
                            ElseIf frmMirage.Toolbar1.buttons(10).value = tbrPressed Then
                                .Fringe2 = EditorTileY * TilesInSheets + EditorTileX
                                .Fringe2Set = EditorSet
                            ElseIf frmMirage.Toolbar1.buttons(17).value = tbrPressed Then
                                .F2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .F2AnimSet = EditorSet
                            ElseIf frmMirage.Toolbar1.buttons(11).value = tbrPressed Then
                                .Fringe3 = EditorTileY * TilesInSheets + EditorTileX
                                .Fringe3Set = EditorSet
                            ElseIf frmMirage.Toolbar1.buttons(18).value = tbrPressed Then
                                .F3Anim = EditorTileY * TilesInSheets + EditorTileX
                                .F3AnimSet = EditorSet
                            End If
                        End With
                    ElseIf frmMirage.tp(3).Checked Then
                        Map(Player(MyIndex).Map).tile(x1, y1).Light = EditorTileY * TilesInSheets + EditorTileX
                    ElseIf frmMirage.tp(2).Checked Then
                        With Map(Player(MyIndex).Map).tile(x1, y1)
                            If frmMirage.optBlocked.value Then .Type = TILE_TYPE_BLOCKED
                            If frmMirage.optWarp.value Then
                                .Type = TILE_TYPE_WARP
                                .Data1 = EditorWarpMap
                                .Data2 = EditorWarpX
                                .Data3 = EditorWarpY
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optHeal.value Then
                                .Type = TILE_TYPE_HEAL
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optKill.value Then
                                .Type = TILE_TYPE_KILL
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optItem.value Then
                                .Type = TILE_TYPE_ITEM
                                .Data1 = ItemEditorNum
                                .Data2 = ItemEditorValue
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optNpcAvoid.value Then
                                .Type = TILE_TYPE_NPCAVOID
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optKey.value Then
                                .Type = TILE_TYPE_KEY
                                .Data1 = KeyEditorNum
                                .Data2 = KeyEditorTake
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optKeyOpen.value Then
                                .Type = TILE_TYPE_KEYOPEN
                                .Data1 = KeyOpenEditorX
                                .Data2 = KeyOpenEditorY
                                .Data3 = 0
                                .String1 = KeyOpenEditorMsg
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optShop.value Then
                                .Type = TILE_TYPE_SHOP
                                .Data1 = EditorShopNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optCBlock.value Then
                                .Type = TILE_TYPE_CBLOCK
                                .Data1 = EditorItemNum1
                                .Data2 = EditorItemNum2
                                .Data3 = EditorItemNum3
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optArena.value Then
                                .Type = TILE_TYPE_ARENA
                                .Data1 = Arena1
                                .Data2 = Arena2
                                .Data3 = Arena3
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optSound.value Then
                                .Type = TILE_TYPE_SOUND
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = SoundFileName
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optSprite.value Then
                                .Type = TILE_TYPE_SPRITE_CHANGE
                                .Data1 = SpritePic
                                .Data2 = SpriteItem
                                .Data3 = SpritePrice
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optSign.value Then
                                .Type = TILE_TYPE_SIGN
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = SignLine1
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optDoor.value Then
                                .Type = TILE_TYPE_DOOR
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optNotice.value Then
                                .Type = TILE_TYPE_NOTICE
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = NoticeTitle
                                .String2 = NoticeText
                                .String3 = NoticeSound
                            ElseIf frmMirage.optClassChange.value Then
                                .Type = TILE_TYPE_CLASS_CHANGE
                                .Data1 = ClassChange
                                .Data2 = ClassChangeReq
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optScripted.value Then
                                .Type = TILE_TYPE_SCRIPTED
                                .Data1 = ScriptNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.OptBank.value Then
                                .Type = TILE_TYPE_BANK
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = bankmsg
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optcoffre.value Then
                                .Type = TILE_TYPE_COFFRE
                                .Data1 = CleCoffreNum
                                .Data2 = CleCoffreSupr
                                .Data3 = ObjCoffreNum
                                .String1 = CodeCoffre
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optportecode.value Then
                                .Type = TILE_TYPE_PORTE_CODE
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = CodePorte
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optBmont.value Then
                                .Type = TILE_TYPE_BLOCK_MONTURE
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optBniv.value Then
                                .Type = TILE_TYPE_BLOCK_NIVEAUX
                                .Data1 = NivMin
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.opttoit.value Then
                                .Type = TILE_TYPE_TOIT
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optBguilde.value Then
                                .Type = TILE_TYPE_BLOCK_GUILDE
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = NomGuilde
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optbtoit.value Then
                                .Type = TILE_TYPE_BLOCK_TOIT
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.optBDir.value Then
                                .Type = TILE_TYPE_BLOCK_DIR
                                .Data1 = AccptDir1
                                .Data2 = AccptDir2
                                .Data3 = AccptDir3
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.OptCraft.value Then
                                .Type = TILE_TYPE_CRAFT
                                .Data1 = ScriptNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            ElseIf frmMirage.OptMetier.value Then
                                .Type = TILE_TYPE_METIER
                                .Data1 = ScriptNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                        End With
                    End If
                Else
                    For y2 = 0 To (frmMirage.shpSelected.Height \ PIC_Y) - 1
                        For x2 = 0 To (frmMirage.shpSelected.Width \ PIC_X) - 1
                            If x1 + x2 <= MAX_MAPX Then
                                If y1 + y2 <= MAX_MAPY Then
                                    If frmMirage.tp(1).Checked = True Then
                                        With Map(Player(MyIndex).Map).tile(x1 + x2, y1 + y2)
                                            If frmMirage.Toolbar1.buttons(5).value = tbrPressed Then
                                                .Ground = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .GroundSet = EditorSet
                                            ElseIf frmMirage.Toolbar1.buttons(6).value = tbrPressed Then
                                                .Mask = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .MaskSet = EditorSet
                                            ElseIf frmMirage.Toolbar1.buttons(13).value = tbrPressed Then
                                                .Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .AnimSet = EditorSet
                                            ElseIf frmMirage.Toolbar1.buttons(7).value = tbrPressed Then
                                                .Mask2 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Mask2Set = EditorSet
                                            ElseIf frmMirage.Toolbar1.buttons(14).value = tbrPressed Then
                                                .M2Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .M2AnimSet = EditorSet
                                            ElseIf frmMirage.Toolbar1.buttons(8).value = tbrPressed Then '<--
                                                .Mask3 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Mask3Set = EditorSet
                                            ElseIf frmMirage.Toolbar1.buttons(15).value = tbrPressed Then '<--
                                                .M3Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .M3AnimSet = EditorSet
                                            ElseIf frmMirage.Toolbar1.buttons(9).value = tbrPressed Then
                                                .Fringe = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .FringeSet = EditorSet
                                            ElseIf frmMirage.Toolbar1.buttons(16).value = tbrPressed Then
                                                .FAnim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .FAnimSet = EditorSet
                                            ElseIf frmMirage.Toolbar1.buttons(10).value = tbrPressed Then
                                                .Fringe2 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Fringe2Set = EditorSet
                                            ElseIf frmMirage.Toolbar1.buttons(17).value = tbrPressed Then
                                                .F2Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .F2AnimSet = EditorSet
                                            ElseIf frmMirage.Toolbar1.buttons(11).value = tbrPressed Then '<--
                                                .Fringe3 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Fringe3Set = EditorSet
                                            ElseIf frmMirage.Toolbar1.buttons(18).value = tbrPressed Then '<--
                                                .F3Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .F3AnimSet = EditorSet
                                            End If
                                        End With
                                    ElseIf frmMirage.tp(3).Checked = True Then
                                        Map(Player(MyIndex).Map).tile(x1 + x2, y1 + y2).Light = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                    End If
                                End If
                            End If
                        Next x2
                    Next y2
                End If
            End If
            End If
            If (Button = 2) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
                If frmMirage.tp(1).Checked = True Then
                    With Map(Player(MyIndex).Map).tile(x1, y1)
                        If frmMirage.Toolbar1.buttons(5).value = tbrPressed Then .Ground = 0
                        If frmMirage.Toolbar1.buttons(6).value = tbrPressed Then .Mask = 0
                        If frmMirage.Toolbar1.buttons(13).value = tbrPressed Then .Anim = 0
                        If frmMirage.Toolbar1.buttons(7).value = tbrPressed Then .Mask2 = 0
                        If frmMirage.Toolbar1.buttons(14).value = tbrPressed Then .M2Anim = 0
                        If frmMirage.Toolbar1.buttons(8).value = tbrPressed Then .Mask3 = 0 '<--
                        If frmMirage.Toolbar1.buttons(15).value = tbrPressed Then .M3Anim = 0 '<--
                        If frmMirage.Toolbar1.buttons(9).value = tbrPressed Then .Fringe = 0
                        If frmMirage.Toolbar1.buttons(16).value = tbrPressed Then .FAnim = 0
                        If frmMirage.Toolbar1.buttons(10).value = tbrPressed Then .Fringe2 = 0
                        If frmMirage.Toolbar1.buttons(17).value = tbrPressed Then .F2Anim = 0
                        If frmMirage.Toolbar1.buttons(11).value = tbrPressed Then .Fringe3 = 0 '<--
                        If frmMirage.Toolbar1.buttons(18).value = tbrPressed Then .F3Anim = 0 '<--
                    End With
                ElseIf frmMirage.tp(3).Checked = True Then
                    Map(Player(MyIndex).Map).tile(x1, y1).Light = 0
                ElseIf frmMirage.tp(2).Checked = True Then
                    With Map(Player(MyIndex).Map).tile(x1, y1)
                        .Type = 0
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End With
                End If
            End If
        End If
If TilesInSheets > 0 Then save = 1
Call WriteINI("modif", "carte" & Player(MyIndex).Map, "1", App.Path & "\config.ini")
Exit Sub
er:
MsgBox "Erreur dans le code d'édition de carte(" & Err.Number & " : " & Err.description & ")" & vbCrLf & "Merci de la rapporter sur le forum de FRoG Creator si elle persiste."
End Sub

Public Sub EditorChooseTile(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then EditorTileX = (x \ PIC_X): EditorTileY = (y \ PIC_Y) + frmMirage.scrlPicture.value
    frmMirage.shpSelected.Top = Int((EditorTileY - frmMirage.scrlPicture.value) * PIC_Y)
    frmMirage.shpSelected.Left = Int(EditorTileX * PIC_Y)
    frmMirage.shpSelected.Visible = True
    frmTile.shpSelected.Top = Int((EditorTileY - frmTile.Defile.value) * PIC_Y)
    frmTile.shpSelected.Left = Int(EditorTileX * PIC_Y)
    frmTile.shpSelected.Visible = True
End Sub

Public Sub EditorChooseTiles(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then EditorTileX = (x \ PIC_X): EditorTileY = (y \ PIC_Y) + frmTile.Defile.value
    frmTile.shpSelected.Top = Int((EditorTileY - frmTile.Defile.value) * PIC_Y)
    frmTile.shpSelected.Left = Int(EditorTileX * PIC_Y)
    frmTile.shpSelected.Visible = True
    frmMirage.shpSelected.Top = Int((EditorTileY - frmMirage.scrlPicture.value) * PIC_Y)
    frmMirage.shpSelected.Left = Int(EditorTileX * PIC_Y)
    frmMirage.shpSelected.Visible = True
End Sub

Public Sub EditorTileScroll()
On Error Resume Next
frmMirage.scrlPicture.Max = ((DDSD_Tile(EditorSet).lHeight - frmMirage.picBackSelect.Height) \ PIC_Y)
If (EditorTileY * PIC_Y) < frmMirage.picBack.Height + (frmMirage.scrlPicture.value * PIC_Y) And (EditorTileY * PIC_Y) > ((frmMirage.scrlPicture.value - 1) * PIC_Y) Then frmMirage.shpSelected.Top = Int((EditorTileY - frmMirage.scrlPicture.value) * PIC_Y): frmMirage.shpSelected.Visible = True Else frmMirage.shpSelected.Visible = False
If frmMirage.scrlPicture.value = 0 Then frmMirage.picBackSelect.Top = 55
Call AffTilesPic(EditorSet, frmMirage.scrlPicture.value * PIC_Y)
End Sub

Public Sub EditorSend()
save = 0
Call WriteINI("modif", "carte" & Player(MyIndex).Map, "0", App.Path & "\config.ini")
If CarteFTP Then
    Call SendData("ENVMAP" & END_CHAR)
Else
    frmmsg.Show
    Call SendMap
End If
End Sub

Public Sub EcrireEtat(Etat As String)
Dim filepath As String
Dim f As Long
On Error Resume Next

If Etat > vbNullString Then Etat = "le : " & Date & "     à : " & Time & "        ..." & Etat & "..."
f = FreeFile
filepath = App.Path & "\LOG.txt"

If FileExiste(filepath) Then If FileLen(filepath) > 6000000 Then Call Kill(filepath)

Open filepath For Append As #f
    Print #f, Etat
Close #f

End Sub

Public Sub EditorCancel()
    InEditor = False
    frmMirage.Show
    frmMirage.MousePointer = 1
End Sub

Public Sub EditorClearLayer()
Dim YesNo As Long, x As Long, y As Long

    ' Ground layer
    If frmMirage.Toolbar1.buttons(5).value = tbrPressed Then
    YesNo = MsgBox("es-tu sur de vouloir éffacer le sol ?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            Call SauvTemp
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(Player(MyIndex).Map).tile(x, y).Ground = 0
                    Map(Player(MyIndex).Map).tile(x, y).GroundSet = 0
                Next x
            Next y
        End If
    End If

    ' Mask layer
    If frmMirage.Toolbar1.buttons(6).value = tbrPressed Then
    YesNo = MsgBox("Es-tu certains de vouloir éffacer les masque ?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            Call SauvTemp
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(Player(MyIndex).Map).tile(x, y).Mask = 0
                    Map(Player(MyIndex).Map).tile(x, y).MaskSet = 0
                Next x
            Next y
        End If
    End If
    
    ' Mask Animation layer
    If frmMirage.Toolbar1.buttons(13).value = tbrPressed Then
    YesNo = MsgBox("Es-tu certains de vouloir éffacer les animations ?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            Call SauvTemp
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(Player(MyIndex).Map).tile(x, y).Anim = 0
                    Map(Player(MyIndex).Map).tile(x, y).AnimSet = 0
                Next x
            Next y
        End If
    End If
    
    ' Mask 2 layer
    If frmMirage.Toolbar1.buttons(7).value = tbrPressed Then
    YesNo = MsgBox("Es-tu certains de vouloir éffacer les masques 2?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            Call SauvTemp
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(Player(MyIndex).Map).tile(x, y).Mask2 = 0
                    Map(Player(MyIndex).Map).tile(x, y).Mask2Set = 0
                Next x
            Next y
        End If
    End If
    
    ' Mask 2 Animation layer
    If frmMirage.Toolbar1.buttons(14).value = tbrPressed Then
    YesNo = MsgBox("Es-tu certains de vouloir éffacer les animations 2?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            Call SauvTemp
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(Player(MyIndex).Map).tile(x, y).M2Anim = 0
                    Map(Player(MyIndex).Map).tile(x, y).M2AnimSet = 0
                Next x
            Next y
        End If
    End If
    
    ' Mask 2 layer
    If frmMirage.Toolbar1.buttons(8).value = tbrPressed Then '<--
    YesNo = MsgBox("Es-tu certains de vouloir éffacer les masques 3?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            Call SauvTemp
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(Player(MyIndex).Map).tile(x, y).Mask3 = 0
                    Map(Player(MyIndex).Map).tile(x, y).Mask3Set = 0
                Next x
            Next y
        End If
    End If
    
    ' Mask 3 Animation layer
    If frmMirage.Toolbar1.buttons(15).value = tbrPressed Then '<--
    YesNo = MsgBox("Es-tu certains de vouloir éffacer les animations 3?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            Call SauvTemp
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(Player(MyIndex).Map).tile(x, y).M3Anim = 0
                    Map(Player(MyIndex).Map).tile(x, y).M3AnimSet = 0
                Next x
            Next y
        End If
    End If
    
    ' Fringe layer
    If frmMirage.Toolbar1.buttons(9).value = tbrPressed Then
    YesNo = MsgBox("Es-tu certains de vouloir éffacer les franges?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            Call SauvTemp
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(Player(MyIndex).Map).tile(x, y).Fringe = 0
                    Map(Player(MyIndex).Map).tile(x, y).FringeSet = 0
                Next x
            Next y
        End If
    End If
    
    ' Fringe Animation layer
    If frmMirage.Toolbar1.buttons(16).value = tbrPressed Then
    YesNo = MsgBox("es-tu certains de vouloir éfface la frange animé?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            Call SauvTemp
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(Player(MyIndex).Map).tile(x, y).FAnim = 0
                    Map(Player(MyIndex).Map).tile(x, y).FAnimSet = 0
                Next x
            Next y
        End If
    End If
    
    ' Fringe 2 layer
    If frmMirage.Toolbar1.buttons(10).value = tbrPressed Then
    YesNo = MsgBox("Es-tu certains de vouloir éffacer la frange 2 ?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            Call SauvTemp
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(Player(MyIndex).Map).tile(x, y).Fringe2 = 0
                    Map(Player(MyIndex).Map).tile(x, y).Fringe2Set = 0
                Next x
            Next y
        End If
    End If
    
    ' Fringe 2 Animation layer
    If frmMirage.Toolbar1.buttons(17).value = tbrPressed Then
    YesNo = MsgBox("Es-tu certains de vouloir éffacer la frange 2 animés ?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            Call SauvTemp
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(Player(MyIndex).Map).tile(x, y).F2Anim = 0
                    Map(Player(MyIndex).Map).tile(x, y).F2AnimSet = 0
                Next x
            Next y
        End If
    End If
    
    ' Fringe 3 layer
    If frmMirage.Toolbar1.buttons(11).value = tbrPressed Then '<--
    YesNo = MsgBox("Es-tu certains de vouloir éffacer la frange 3 ?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            Call SauvTemp
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(Player(MyIndex).Map).tile(x, y).Fringe3 = 0
                    Map(Player(MyIndex).Map).tile(x, y).Fringe3Set = 0
                Next x
            Next y
        End If
    End If
    
    ' Fringe 3 Animation layer
    If frmMirage.Toolbar1.buttons(18).value = tbrPressed Then '<--
    YesNo = MsgBox("Es-tu certains de vouloir éffacer la frange 3 animés ?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            Call SauvTemp
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(Player(MyIndex).Map).tile(x, y).F3Anim = 0
                    Map(Player(MyIndex).Map).tile(x, y).F3AnimSet = 0
                Next x
            Next y
        End If
    End If
End Sub

Public Sub EditorClearAttribs()
Dim YesNo As Long, x As Long, y As Long

    YesNo = MsgBox("Es-tu certains de vouloir éffacer les attributs de la maps?", vbYesNo, GAME_NAME)
    If YesNo = vbYes Then
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Map(Player(MyIndex).Map).tile(x, y).Type = 0
            Next x
        Next y
    End If
End Sub

Public Sub EmoticonEditorInit()
    frmEmoticonEditor.scrlEmoticon.Max = MAX_EMOTICONS
    frmEmoticonEditor.scrlEmoticon.value = Emoticons(EditorIndex - 1).Pic
    frmEmoticonEditor.txtCommand.Text = Trim$(Emoticons(EditorIndex - 1).Command)
    'frmEmoticonEditor.picEmoticons.Picture = LoadPNG(App.Path & "\GFX\emoticons.png")
    frmEmoticonEditor.Show vbModeless, frmMirage
End Sub

Public Sub EmoticonEditorOk()
    Emoticons(EditorIndex - 1).Pic = frmEmoticonEditor.scrlEmoticon.value
    If frmEmoticonEditor.txtCommand.Text <> "/" Then
        Emoticons(EditorIndex - 1).Command = frmEmoticonEditor.txtCommand.Text
    Else
        Emoticons(EditorIndex - 1).Command = vbNullString
    End If
    Call SendSaveEmoticon(EditorIndex - 1)
    Call EmoticonEditorCancel
End Sub
Sub EnvoieServeur()
Dim valueini As String
Dim i As Long
Dim a As Long

a = 0
If frmenvoier.Visible = True Then frmenvoier.SetFocus: Exit Sub

Call frmenvoier.TreeView1.Nodes.Add(, , "obj", "Objets")
For i = 1 To MAX_ITEMS
    valueini = ReadINI("modif", "objet" & i, App.Path & "\config.ini")
    If valueini = "1" Then
        Call frmenvoier.TreeView1.Nodes.Add("obj", tvwChild, , "Objet" & i)
        a = a + 1
        frmenvoier.TreeView1.Nodes(frmenvoier.TreeView1.Nodes("obj").Index + a).Tag = i
    End If
Next i

a = 0

Call frmenvoier.TreeView1.Nodes.Add(, , "mag", "Magasins")
For i = 1 To MAX_SHOPS
    valueini = ReadINI("modif", "magasin" & i, App.Path & "\config.ini")
    If valueini = "1" Then
        Call frmenvoier.TreeView1.Nodes.Add("mag", tvwChild, , "Magasin" & i)
        a = a + 1
        frmenvoier.TreeView1.Nodes(frmenvoier.TreeView1.Nodes("mag").Index + a).Tag = i
    End If
Next i

a = 0
Call frmenvoier.TreeView1.Nodes.Add(, , "sort", "Sorts")
For i = 1 To MAX_SPELLS
    valueini = ReadINI("modif", "sort" & i, App.Path & "\config.ini")
    If valueini = "1" Then
        Call frmenvoier.TreeView1.Nodes.Add("sort", tvwChild, , "Sort" & i)
        a = a + 1
        frmenvoier.TreeView1.Nodes(frmenvoier.TreeView1.Nodes("sort").Index + a).Tag = i
    End If
Next i

a = 0
Call frmenvoier.TreeView1.Nodes.Add(, , "pnj", "PNJs")
For i = 1 To MAX_NPCS
    valueini = ReadINI("modif", "pnj" & i, App.Path & "\config.ini")
    If valueini = "1" Then
        Call frmenvoier.TreeView1.Nodes.Add("pnj", tvwChild, , "PNJ" & i)
        a = a + 1
        frmenvoier.TreeView1.Nodes(frmenvoier.TreeView1.Nodes("pnj").Index + a).Tag = i
    End If
Next i

a = 0
Call frmenvoier.TreeView1.Nodes.Add(, , "flc", "Flêches")
For i = 1 To MAX_ARROWS
    valueini = ReadINI("modif", "flêche" & i, App.Path & "\config.ini")
    If valueini = "1" Then
        Call frmenvoier.TreeView1.Nodes.Add("flc", tvwChild, , "Flêche" & i)
        a = a + 1
        frmenvoier.TreeView1.Nodes(frmenvoier.TreeView1.Nodes("flc").Index + a).Tag = i
    End If
Next i

a = 0
Call frmenvoier.TreeView1.Nodes.Add(, , "emot", "Emoticons")
For i = 1 To MAX_EMOTICONS
    valueini = ReadINI("modif", "emot" & i, App.Path & "\config.ini")
    If valueini = "1" Then
        Call frmenvoier.TreeView1.Nodes.Add("emot", tvwChild, , "Emoticon" & i)
        a = a + 1
        frmenvoier.TreeView1.Nodes(frmenvoier.TreeView1.Nodes("emot").Index + a).Tag = i
    End If
Next i

a = 0
Call frmenvoier.TreeView1.Nodes.Add(, , "quête", "Quêtes")
For i = 1 To MAX_QUETES
    valueini = ReadINI("modif", "quête" & i, App.Path & "\config.ini")
    If valueini = "1" Then
        Call frmenvoier.TreeView1.Nodes.Add("quête", tvwChild, Chr(i), "Quête" & i)
        a = a + 1
        frmenvoier.TreeView1.Nodes(frmenvoier.TreeView1.Nodes("quête").Index + a).Tag = i
    End If
Next i

frmenvoier.Show
End Sub
Public Sub EmoticonEditorCancel()
Dim i As Long
    Unload frmEmoticonEditor
    frmIndex.lstIndex.Clear
    For i = 0 To MAX_EMOTICONS
        frmIndex.lstIndex.AddItem i & " : " & Trim$(Emoticons(i).Command)
    Next i
    frmIndex.SetFocus
End Sub

Public Sub ArrowEditorInit()
    frmEditArrows.scrlArrow.Max = MAX_ARROWS
    If Arrows(EditorIndex).Pic = 0 Then Arrows(EditorIndex).Pic = 1
    frmEditArrows.scrlArrow.value = Arrows(EditorIndex).Pic
    frmEditArrows.txtName.Text = Arrows(EditorIndex).name
    If Arrows(EditorIndex).Range = 0 Then Arrows(EditorIndex).Range = 1
    frmEditArrows.scrlRange.value = Arrows(EditorIndex).Range
    Call AffSurfPic(DD_ArrowAnim, frmEditArrows.picArrows, 3 * PIC_X, frmEditArrows.scrlArrow.value * PIC_Y)
    'frmEditArrows.picArrows.Picture = LoadPNG(App.Path & "\GFX\arrows.png")
    frmEditArrows.Show vbModeless, frmMirage
End Sub

Public Sub ArrowEditorOk()
    Arrows(EditorIndex).Pic = frmEditArrows.scrlArrow.value
    Arrows(EditorIndex).Range = frmEditArrows.scrlRange.value
    Arrows(EditorIndex).name = frmEditArrows.txtName.Text
    Call SendSaveArrow(EditorIndex)
    Call ArrowEditorCancel
End Sub

Public Sub ArrowEditorCancel()
Dim i As Long
    Unload frmEditArrows
    frmIndex.lstIndex.Clear
    For i = 1 To MAX_ARROWS
        frmIndex.lstIndex.AddItem i & " : " & Trim$(Arrows(i).name)
    Next i
    frmIndex.SetFocus
End Sub

Public Sub PetEditorInit()
    ' EditorIndex
    frmPets.TxtNom = Pets(EditorIndex).nom
    frmPets.ScrlApp.Max = MAX_DX_PETS
    frmPets.ScrlApp.value = Pets(EditorIndex).sprite
    frmPets.ScrlForce.value = Pets(EditorIndex).addForce
    frmPets.ScrlDefence.value = Pets(EditorIndex).addDefence
    frmPets.PictApp.Picture = LoadPNG(App.Path & "\GFX\Pets\Pet" & Pets(EditorIndex).sprite & ".png")
    frmPets.Show vbModeless, frmMirage
End Sub

Public Sub PetEditorOk()
    Pets(EditorIndex).nom = frmPets.TxtNom
    Pets(EditorIndex).sprite = frmPets.ScrlApp.value
    Pets(EditorIndex).addForce = frmPets.ScrlForce.value
    Pets(EditorIndex).addDefence = frmPets.ScrlDefence.value
    Call SendSavePet(EditorIndex)
    Call PetEditorCancel
End Sub

Public Sub PetEditorCancel()
    Dim i As Long
    Unload frmPets
    frmIndex.lstIndex.Clear
    ' Add the names
    For i = 1 To MAX_PETS
        frmIndex.lstIndex.AddItem i & " : " & Trim$(Pets(i).nom)
    Next i
    frmIndex.SetFocus
End Sub

Public Sub MetierEditorInit()
    ' EditorIndex
    frmMetier.TxtNom = Metier(EditorIndex).nom
    frmMetier.txtDesc = Metier(EditorIndex).desc
    frmMetier.CMetier.ListIndex = Metier(EditorIndex).Type
    If frmMetier.CMetier.ListIndex = METIER_CHASSEUR Then
        frmMetier.frChasseur.Caption = "Tuer"
        frmMetier.Label4.Caption = "Numéro PNJ"
    ElseIf frmMetier.CMetier.ListIndex = METIER_CRAFT Then
        frmMetier.frChasseur.Caption = "Craft"
        frmMetier.Label4.Caption = "Numéro Recette"
    End If
    frmMetier.scrlCibleNPC.value = 0
    frmMetier.scrlNPCNum.value = Metier(EditorIndex).Data(0, 0)
    frmMetier.scrlExpNPC.value = Metier(EditorIndex).Data(0, 1)
    frmMetier.Show vbModeless, frmMirage
End Sub

Public Sub MetierEditorOk()
    Metier(EditorIndex).nom = frmMetier.TxtNom
    Metier(EditorIndex).desc = frmMetier.txtDesc
    Metier(EditorIndex).Type = frmMetier.CMetier.ListIndex
    Call SendSaveMetier(EditorIndex)
    Call MetierEditorCancel
End Sub

Public Sub MetierEditorCancel()
    Dim i As Long
    Unload frmMetier
    frmIndex.lstIndex.Clear
    ' Add the names
    For i = 1 To MAX_METIER
        frmIndex.lstIndex.AddItem i & " : " & Trim$(Metier(i).nom)
    Next i
    frmIndex.SetFocus
End Sub

Public Sub recetteEditorInit()
Dim i As Long
    frmRecette.TxtNom = recette(EditorIndex).nom
    For i = 0 To 9
        frmRecette.scrlItemNum1(i).value = recette(EditorIndex).InCraft(i, 0)
        If recette(EditorIndex).InCraft(i, 1) > 0 Then frmRecette.scrlItemQ(i).value = recette(EditorIndex).InCraft(i, 1)
        If recette(EditorIndex).InCraft(i, 0) > 0 Then
            If Item(recette(EditorIndex).InCraft(i, 0)).Type = ITEM_TYPE_CURRENCY Or Item(recette(EditorIndex).InCraft(i, 0)).Empilable <> 0 Then
                frmRecette.scrlItemQ(i).Enabled = True
            Else
                frmRecette.scrlItemQ(i).Enabled = False
            End If
        Else
        frmRecette.lblItemNum1(i).Caption = "Pas d'objet"
        End If
    Next i
    frmRecette.scrlItemNum1(10).value = recette(EditorIndex).craft(0)
    If recette(EditorIndex).craft(1) > 0 Then
        frmRecette.scrlItemQ(10).value = recette(EditorIndex).craft(1)
        
        If Item(frmRecette.scrlItemNum1(10).value).Type = ITEM_TYPE_CURRENCY Or Item(frmRecette.scrlItemNum1(10).value).Empilable <> 0 Then
            frmRecette.scrlItemQ(10).Enabled = True
        Else
            frmRecette.scrlItemQ(10).Enabled = False
        End If
    End If

    frmRecette.Show vbModeless, frmMirage
End Sub

Public Sub recetteEditorOk()
Dim i As Long
    recette(EditorIndex).nom = frmRecette.TxtNom
    For i = 0 To 9
        recette(EditorIndex).InCraft(i, 0) = frmRecette.scrlItemNum1(i).value
        recette(EditorIndex).InCraft(i, 1) = frmRecette.scrlItemQ(i).value
    Next i
    recette(EditorIndex).craft(0) = frmRecette.scrlItemNum1(10).value
    recette(EditorIndex).craft(1) = frmRecette.scrlItemQ(10).value
    Call SendSaverecette(EditorIndex)
    Call recetteEditorCancel
End Sub

Public Sub recetteEditorCancel()
    Dim i As Long
    Unload frmRecette
    frmIndex.lstIndex.Clear
    ' Add the names
    For i = 1 To MAX_RECETTE
        frmIndex.lstIndex.AddItem i & " : " & Trim$(recette(i).nom)
    Next i
    frmIndex.SetFocus
End Sub

Public Sub ItemEditorInit()
Dim i As Long
    EditorItemY = (Item(EditorIndex).Pic \ 6)
    EditorItemX = (Item(EditorIndex).Pic - (Item(EditorIndex).Pic \ 6) * 6)
    
    frmItemEditor.scrlClassReq.Max = Max_Classes
    
    frmItemEditor.VScroll1.Max = DDSD_Item.lHeight \ PIC_X
    'frmItemEditor.picItems.Picture = LoadPNG(App.Path & "\GFX\items.png")
    
    frmItemEditor.txtName.Text = Trim$(Item(EditorIndex).name)
    frmItemEditor.txtDesc.Text = Trim$(Item(EditorIndex).desc)
    frmItemEditor.cmbType.ListIndex = Item(EditorIndex).Type
    
    frmItemEditor.PicPD.Picture = LoadPNG(App.Path & "\GFX\Paperdolls\Paperdolls0.png")
    
    frmItemEditor.CheckEmpi.value = Item(EditorIndex).Empilable
    frmItemEditor.CheckEmpi.Enabled = True
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.fraAttributes.Visible = True
        frmItemEditor.fraBow.Visible = True
        
        frmItemEditor.scrlDurability.value = Item(EditorIndex).Data1
        frmItemEditor.scrlStrength.value = Item(EditorIndex).Data2
        frmItemEditor.scrlStrReq.value = Item(EditorIndex).StrReq
        frmItemEditor.scrlDefReq.value = Item(EditorIndex).DefReq
        frmItemEditor.scrlSpeedReq.value = Item(EditorIndex).SpeedReq
        frmItemEditor.scrlClassReq.value = Item(EditorIndex).ClassReq
        frmItemEditor.scrlAccessReq.value = Item(EditorIndex).AccessReq
        frmItemEditor.scrlAddHP.value = Item(EditorIndex).AddHP
        frmItemEditor.scrlAddMP.value = Item(EditorIndex).AddMP
        frmItemEditor.scrlAddSP.value = Item(EditorIndex).AddSP
        frmItemEditor.scrlAddStr.value = Item(EditorIndex).AddStr
        frmItemEditor.scrlAddDef.value = Item(EditorIndex).AddDef
        frmItemEditor.scrlAddMagi.value = Item(EditorIndex).AddMagi
        frmItemEditor.scrlAddSpeed.value = Item(EditorIndex).AddSpeed
        frmItemEditor.scrlAddEXP.value = Item(EditorIndex).AddEXP
        frmItemEditor.scrlAttackSpeed.value = Item(EditorIndex).AttackSpeed
        If Item(EditorIndex).Data3 > 0 Then frmItemEditor.chkBow.value = Checked Else frmItemEditor.chkBow.value = Unchecked
        
        frmItemEditor.cmbBow.Clear
        If frmItemEditor.chkBow.value = Checked Then
            For i = 1 To 100
                frmItemEditor.cmbBow.AddItem i & " : " & Arrows(i).name
            Next i
            frmItemEditor.cmbBow.ListIndex = Item(EditorIndex).Data3 - 1
            Call AffSurfPic(DD_ArrowAnim, frmItemEditor.picBow, 3 * PIC_X, Arrows(Item(EditorIndex).Data3).Pic * PIC_Y)
            'frmItemEditor.picBow.Top = (Arrows(Item(EditorIndex).Data3).Pic * 32) * -1
            frmItemEditor.cmbBow.Enabled = True
        Else
            frmItemEditor.cmbBow.AddItem "Aucune"
            frmItemEditor.cmbBow.ListIndex = 0
            frmItemEditor.cmbBow.Enabled = False
        End If
        
        'paperdoll
        frmItemEditor.FramePD.Visible = True
        frmItemEditor.CheckPD.value = Item(EditorIndex).paperdoll
        frmItemEditor.scrlPD.value = Item(EditorIndex).paperdollPic
        frmItemEditor.PicPD.Picture = LoadPNG(App.Path & "\GFX\Paperdolls\Paperdolls" & Item(EditorIndex).paperdollPic & ".png")
        
        
        frmItemEditor.CheckEmpi.value = 0
        frmItemEditor.CheckEmpi.Enabled = False
        
        frmItemEditor.scrlSexReq.value = Item(EditorIndex).Sex
    Else
        frmItemEditor.fraEquipment.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        frmItemEditor.fraVitals.Visible = True
        frmItemEditor.scrlVitalMod.value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraVitals.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        frmItemEditor.fraSpell.Visible = True
        frmItemEditor.scrlSpell.value = Item(EditorIndex).Data1
        frmItemEditor.scrlSpell.Max = MAX_SPELLS
        frmItemEditor.lblSpellName.Caption = Trim$(Spell(Item(EditorIndex).Data1).name)
    Else
        frmItemEditor.fraSpell.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SCRIPT) Then
        frmItemEditor.fraobjsc.Visible = True
        frmItemEditor.HScroll1.value = Item(EditorIndex).Data1
        frmItemEditor.disp.value = Item(EditorIndex).Data2
    Else
        frmItemEditor.fraobjsc.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_MONTURE) Then
        frmItemEditor.framonture.Visible = True
        frmItemEditor.skin.value = Item(EditorIndex).Data1
        frmItemEditor.vit.value = Item(EditorIndex).Data2
        
        frmItemEditor.CheckEmpi.value = 0
        frmItemEditor.CheckEmpi.Enabled = False
    Else
        frmItemEditor.framonture.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_PET) Then
        frmItemEditor.fraPet.Visible = True
        frmItemEditor.lblPetNom = Pets(Item(EditorIndex).Data1).nom
        frmItemEditor.scrlPet.value = Item(EditorIndex).Data1
        frmItemEditor.CheckEmpi.Enabled = False
    Else
        frmItemEditor.fraPet.Visible = False
    End If
    
    frmItemEditor.coul.BackColor = Item(EditorIndex).NCoul
    frmItemEditor.coul.Tag = Item(EditorIndex).NCoul
    frmItemEditor.txtName.ForeColor = Item(EditorIndex).NCoul

    frmItemEditor.Show vbModeless, frmMirage
End Sub

Public Sub ItemEditorOk()
    Item(EditorIndex).name = frmItemEditor.txtName.Text
    Item(EditorIndex).desc = frmItemEditor.txtDesc.Text
    Item(EditorIndex).Pic = EditorItemY * 6 + EditorItemX
    Item(EditorIndex).Type = frmItemEditor.cmbType.ListIndex
    Item(EditorIndex).NCoul = frmItemEditor.coul.Tag
    
    Item(EditorIndex).paperdoll = 0
    Item(EditorIndex).paperdollPic = 0
      
    Item(EditorIndex).Empilable = frmItemEditor.CheckEmpi.value
    Item(EditorIndex).tArme = 0

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.value
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.value
        If frmItemEditor.chkBow.value = Checked Then Item(EditorIndex).Data3 = frmItemEditor.cmbBow.ListIndex + 1 Else Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = frmItemEditor.scrlStrReq.value
        Item(EditorIndex).DefReq = frmItemEditor.scrlDefReq.value
        Item(EditorIndex).SpeedReq = frmItemEditor.scrlSpeedReq.value
        
        Item(EditorIndex).ClassReq = frmItemEditor.scrlClassReq.value
        Item(EditorIndex).AccessReq = frmItemEditor.scrlAccessReq.value
        
        Item(EditorIndex).AddHP = frmItemEditor.scrlAddHP.value
        Item(EditorIndex).AddMP = frmItemEditor.scrlAddMP.value
        Item(EditorIndex).AddSP = frmItemEditor.scrlAddSP.value
        Item(EditorIndex).AddStr = frmItemEditor.scrlAddStr.value
        Item(EditorIndex).AddDef = frmItemEditor.scrlAddDef.value
        Item(EditorIndex).AddMagi = frmItemEditor.scrlAddMagi.value
        Item(EditorIndex).AddSpeed = frmItemEditor.scrlAddSpeed.value
        Item(EditorIndex).AddEXP = frmItemEditor.scrlAddEXP.value
        Item(EditorIndex).AttackSpeed = frmItemEditor.scrlAttackSpeed.value
        
        Item(EditorIndex).paperdoll = frmItemEditor.CheckPD.value
        Item(EditorIndex).paperdollPic = frmItemEditor.scrlPD.value
        
        Item(EditorIndex).Empilable = 0
        Item(EditorIndex).Sex = frmItemEditor.scrlSexReq.value
        If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_WEAPON) Then
            Item(EditorIndex).tArme = frmItemEditor.CtArme.ListIndex
        End If
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlVitalMod.value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).ClassReq = -1
        Item(EditorIndex).AccessReq = 0
        
        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddStr = 0
        Item(EditorIndex).AddDef = 0
        Item(EditorIndex).AddMagi = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
        Item(EditorIndex).AttackSpeed = 0
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlSpell.value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).ClassReq = -1
        Item(EditorIndex).AccessReq = 0
        
        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddStr = 0
        Item(EditorIndex).AddDef = 0
        Item(EditorIndex).AddMagi = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
        Item(EditorIndex).AttackSpeed = 0
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_MONTURE) Then
        Item(EditorIndex).Data1 = Val(frmItemEditor.skin.value)
        Item(EditorIndex).Data2 = Val(frmItemEditor.vit.value)
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).ClassReq = -1
        Item(EditorIndex).AccessReq = 0
        
        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddStr = 0
        Item(EditorIndex).AddDef = 0
        Item(EditorIndex).AddMagi = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
        Item(EditorIndex).AttackSpeed = 0
        
        Item(EditorIndex).Empilable = 0
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SCRIPT) Then
        Item(EditorIndex).Data1 = frmItemEditor.HScroll1.value
        Item(EditorIndex).Data2 = frmItemEditor.disp.value
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).ClassReq = -1
        Item(EditorIndex).AccessReq = 0
        
        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddStr = 0
        Item(EditorIndex).AddDef = 0
        Item(EditorIndex).AddMagi = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
        Item(EditorIndex).AttackSpeed = 0
    End If
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_PET) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlPet.value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).ClassReq = -1
        Item(EditorIndex).AccessReq = 0
        
        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddStr = 0
        Item(EditorIndex).AddDef = 0
        Item(EditorIndex).AddMagi = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
        Item(EditorIndex).AttackSpeed = 0
    End If
    Call SendSaveItem(EditorIndex)
    Call ItemEditorCancel
End Sub

Public Sub ItemEditorCancel()
    Dim i As Long
    Unload frmItemEditor
    frmIndex.lstIndex.Clear
    ' Add the names
    For i = 1 To MAX_ITEMS
        frmIndex.lstIndex.AddItem i & " : " & Trim$(Item(i).name)
    Next i
    frmIndex.SetFocus
End Sub

Public Sub NpcEditorInit()
Dim i As Byte
If Not FileExiste("pnjs\npc" & EditorIndex & ".fcp") And HORS_LIGNE = 1 Then Call ClearNpc(EditorIndex)
    'frmNpcEditor.picSprites.Picture = LoadPNG(App.Path & "\GFX\sprites.png")
    
    frmNpcEditor.txtName.Text = Trim$(Npc(EditorIndex).name)
    frmNpcEditor.txtAttackSay.Text = Trim$(Npc(EditorIndex).AttackSay)
    frmNpcEditor.scrlSprite.value = Npc(EditorIndex).sprite
    frmNpcEditor.txtSpawnSecs.Text = CStr(Npc(EditorIndex).SpawnSecs)
    frmNpcEditor.cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
    frmNpcEditor.scrlRange.value = Npc(EditorIndex).Range
    frmNpcEditor.scrlSTR.value = Npc(EditorIndex).STR
    frmNpcEditor.scrlDEF.value = Npc(EditorIndex).def
    frmNpcEditor.scrlSPEED.value = Npc(EditorIndex).speed
    frmNpcEditor.scrlMAGI.value = Npc(EditorIndex).magi
    If Npc(EditorIndex).MaxHp = 0 And Trim$(Npc(EditorIndex).name) <> vbNullString Then frmNpcEditor.StartHP.value = 1 Else frmNpcEditor.StartHP.value = Npc(EditorIndex).MaxHp
    frmNpcEditor.ExpGive.value = Npc(EditorIndex).exp
    frmNpcEditor.txtChance.Text = CStr(Npc(EditorIndex).ItemNPC(1).Chance)
    If MAX_ITEMS <= 32000 Then frmNpcEditor.scrlNum.Max = MAX_ITEMS
    frmNpcEditor.scrlNum.value = Npc(EditorIndex).ItemNPC(1).ItemNum
    frmNpcEditor.scrlValue.value = Npc(EditorIndex).ItemNPC(1).ItemValue
    If frmNpcEditor.cmbBehavior.ListIndex >= 0 And frmNpcEditor.cmbBehavior.ListIndex <= 1 Then
        If Npc(EditorIndex).quetenum < 0 Then Npc(EditorIndex).quetenum = 0
        frmNpcEditor.quetenum.value = Npc(EditorIndex).quetenum
    Else
        If Npc(EditorIndex).quetenum <= 0 Then Npc(EditorIndex).quetenum = 1
        frmNpcEditor.quetenum.value = Npc(EditorIndex).quetenum
    End If
    
    If Npc(EditorIndex).SpawnTime = 0 Then
        frmNpcEditor.chkDay.value = Checked
        frmNpcEditor.chkNight.value = Checked
    ElseIf Npc(EditorIndex).SpawnTime = 1 Then
        frmNpcEditor.chkDay.value = Checked
        frmNpcEditor.chkNight.value = Unchecked
    ElseIf Npc(EditorIndex).SpawnTime = 2 Then
        frmNpcEditor.chkDay.value = Unchecked
        frmNpcEditor.chkNight.value = Checked
    End If
    For i = 1 To MAX_NPC_SPELLS
        If Npc(EditorIndex).Spell(i) > 0 Then frmNpcEditor.lstSpells.AddItem frmNpcEditor.lstSpells.ListCount + 1 & ". " & Spell(Npc(EditorIndex).Spell(i)).name: frmNpcEditor.lstSpells.ItemData(frmNpcEditor.lstSpells.ListCount - 1) = Npc(EditorIndex).Spell(i)
    Next
    If Npc(EditorIndex).inv <> 0 Then frmNpcEditor.invi.value = 1: frmNpcEditor.StartHP.Enabled = False Else frmNpcEditor.invi.value = 0: frmNpcEditor.StartHP.Enabled = True
    If Npc(EditorIndex).vol <> 0 Then frmNpcEditor.vol.value = 1 Else frmNpcEditor.vol.value = 0
    frmNpcEditor.Show vbModeless, frmMirage
End Sub

Public Sub NpcEditorOk()
Dim i As Byte

    Npc(EditorIndex).name = frmNpcEditor.txtName.Text
    If Trim$(Npc(EditorIndex).name) = vbNullString Then Npc(EditorIndex).name = "**"
    Npc(EditorIndex).AttackSay = frmNpcEditor.txtAttackSay.Text
    Npc(EditorIndex).sprite = frmNpcEditor.scrlSprite.value
    Npc(EditorIndex).SpawnSecs = Val(frmNpcEditor.txtSpawnSecs.Text)
    Npc(EditorIndex).Behavior = frmNpcEditor.cmbBehavior.ListIndex
    Npc(EditorIndex).Range = frmNpcEditor.scrlRange.value
    Npc(EditorIndex).STR = frmNpcEditor.scrlSTR.value
    Npc(EditorIndex).def = frmNpcEditor.scrlDEF.value
    Npc(EditorIndex).speed = frmNpcEditor.scrlSPEED.value
    Npc(EditorIndex).magi = frmNpcEditor.scrlMAGI.value
    Npc(EditorIndex).MaxHp = frmNpcEditor.StartHP.value
    Npc(EditorIndex).exp = frmNpcEditor.ExpGive.value
    
    If frmNpcEditor.chkDay.value = Checked And frmNpcEditor.chkNight.value = Checked Then
        Npc(EditorIndex).SpawnTime = 0
    ElseIf frmNpcEditor.chkDay.value = Checked And frmNpcEditor.chkNight.value = Unchecked Then
        Npc(EditorIndex).SpawnTime = 1
    ElseIf frmNpcEditor.chkDay.value = Unchecked And frmNpcEditor.chkNight.value = Checked Then
        Npc(EditorIndex).SpawnTime = 2
    End If
    Npc(EditorIndex).inv = CLng(frmNpcEditor.invi.value)
    Npc(EditorIndex).vol = CLng(frmNpcEditor.vol.value)
    
    For i = 1 To MAX_NPC_SPELLS
        If frmNpcEditor.lstSpells.ListCount >= i Then
            Npc(EditorIndex).Spell(i) = frmNpcEditor.lstSpells.ItemData(i - 1)
        Else
            Npc(EditorIndex).Spell(i) = 0
        End If
    Next
        
    
    Call SendSaveNpc(EditorIndex)
    Call NpcEditorCancel
End Sub

Public Sub NpcEditorCancel()
Dim i As Long
    Unload frmNpcEditor
    frmIndex.lstIndex.Clear
    ' Add the names
    For i = 1 To MAX_NPCS
        frmIndex.lstIndex.AddItem i & " : " & Trim$(Npc(i).name)
    Next i
    If frmIndex.Visible Then frmIndex.SetFocus
End Sub

Public Sub NpcEditorBltSprite()
    'If frmNpcEditor.BigNpc.value = Checked Then
        ''Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, 64, 64, frmNpcEditor.picSprites.hDC, 3 * 64, frmNpcEditor.scrlSprite.value * 64, SRCCOPY)
    'Else
        ''Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, PIC_X, PIC_Y * PIC_NPC1, frmNpcEditor.picSprites.hDC, 3 * PIC_X, frmNpcEditor.scrlSprite.value * (PIC_Y * PIC_NPC1), SRCCOPY)
    'End If
End Sub

Public Sub OptionSave()
Dim PathServ As String

PathServ = Mid$(App.Path, 1, Len(App.Path) - Len(Dir$(App.Path, vbDirectory))) & "Serveur"

If LCase$(Dir$(PathServ, vbDirectory)) <> "serveur" Then

    Call MsgBox("Dossier du serveur introuvable les modifications niveau serveur ne seront pas prises en comptes.")

    Call WriteINI("INFO", "MaxClasses", frmoptions.nbcls.Text, App.Path & "\Classes\info.ini")
    Call WriteINI("INFO", "HPRegen", frmoptions.PV, App.Path & "\config.ini")
    Call WriteINI("INFO", "MPRegen", frmoptions.pm, App.Path & "\config.ini")
    Call WriteINI("INFO", "SPRegen", frmoptions.ps, App.Path & "\config.ini")
    Call WriteINI("CONFIG", "Scrolling", frmoptions.defl, App.Path & "\config.ini")
    Call WriteINI("CONFIG", "Scripting", frmoptions.script, App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxplayers", Val(frmoptions.mj), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxitems", Val(frmoptions.mo), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxnpcs", Val(frmoptions.mpnj), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxshops", Val(frmoptions.mm), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxspells", Val(frmoptions.ms), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxmaps", Val(frmoptions.mc), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxmapitems", Val(frmoptions.moc), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxemots", Val(frmoptions.Me), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxlevel", Val(frmoptions.mn), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxquet", Val(frmoptions.mq), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxguilds", Val(frmoptions.mg), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxjguild", Val(frmoptions.mjg), App.Path & "\config.ini")
    Call WriteINI("INFO", "GameName", GAME_NAME, App.Path & "\config.ini")
    Call WriteINI("INFO", "motd", frmoptions.motd, App.Path & "\config.ini")
    
    frmMirage.Caption = "Editeur pour le jeu : " & Trim$(GAME_NAME) & " Mettez votre souris sur un élément pour plus de détails."
    App.Title = GAME_NAME
    If HORS_LIGNE = False Then Call SendMOTDChange(frmoptions.motd.Text)
    
    Call Unload(frmmsg)
Else
    GAME_NAME = frmoptions.nom
    WEBSITE = frmoptions.site
    
    Call WriteINI("INFO", "MaxClasses", frmoptions.nbcls.Text, App.Path & "\Classes\info.ini")
    Call WriteINI("INFO", "HPRegen", frmoptions.PV, App.Path & "\config.ini")
    Call WriteINI("INFO", "MPRegen", frmoptions.pm, App.Path & "\config.ini")
    Call WriteINI("INFO", "SPRegen", frmoptions.ps, App.Path & "\config.ini")
    Call WriteINI("CONFIG", "Scrolling", frmoptions.defl, App.Path & "\config.ini")
    Call WriteINI("CONFIG", "Scripting", frmoptions.script, App.Path & "\config.ini")
    Call WriteINI("INFO", "GameName", GAME_NAME, App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxplayers", Val(frmoptions.mj), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxitems", Val(frmoptions.mo), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxnpcs", Val(frmoptions.mpnj), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxshops", Val(frmoptions.mm), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxspells", Val(frmoptions.ms), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxmaps", Val(frmoptions.mc), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxmapitems", Val(frmoptions.moc), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxemots", Val(frmoptions.Me), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxlevel", Val(frmoptions.mn), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxquet", Val(frmoptions.mq), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxguilds", Val(frmoptions.mg), App.Path & "\config.ini")
    Call WriteINI("INFO", "Maxjguild", Val(frmoptions.mjg), App.Path & "\config.ini")
    Call WriteINI("INFO", "motd", frmoptions.motd, App.Path & "\config.ini")
    
    frmMirage.Caption = "Editeur pour le jeu : " & Trim$(GAME_NAME) & " Mettez votre souris sur un élément pour plus de détails."
    App.Title = GAME_NAME

    Call WriteINI("INFO", "MaxClasses", frmoptions.nbcls.Text, PathServ & "\Classes\info.ini")
    Call WriteINI("CONFIG", "GameName", frmoptions.nom, PathServ & "\Data.ini")
    Call WriteINI("CONFIG", "WebSite", frmoptions.site, PathServ & "\Data.ini")
    Call WriteINI("CONFIG", "HPRegen", frmoptions.PV, PathServ & "\Data.ini")
    Call WriteINI("CONFIG", "MPRegen", frmoptions.pm, PathServ & "\Data.ini")
    Call WriteINI("CONFIG", "SPRegen", frmoptions.ps, PathServ & "\Data.ini")
    Call WriteINI("INFO", "HPRegen", frmoptions.PV, PathServ & "\Data.ini")
    Call WriteINI("INFO", "MPRegen", frmoptions.pm, PathServ & "\Data.ini")
    Call WriteINI("INFO", "SPRegen", frmoptions.ps, PathServ & "\Data.ini")
    Call WriteINI("CONFIG", "Scrolling", frmoptions.defl, PathServ & "\Data.ini")
    Call WriteINI("CONFIG", "Scripting", frmoptions.script, PathServ & "\Data.ini")
    Call WriteINI("MAX", "MAX_PLAYERS", frmoptions.mj, PathServ & "\Data.ini")
    Call WriteINI("MAX", "MAX_ITEMS", frmoptions.mo, PathServ & "\Data.ini")
    Call WriteINI("MAX", "MAX_NPCS", frmoptions.mpnj, PathServ & "\Data.ini")
    Call WriteINI("MAX", "MAX_SHOPS", frmoptions.mm, PathServ & "\Data.ini")
    Call WriteINI("MAX", "MAX_SPELLS", frmoptions.ms, PathServ & "\Data.ini")
    Call WriteINI("MAX", "MAX_MAPS", frmoptions.mc, PathServ & "\Data.ini")
    Call WriteINI("MAX", "MAX_MAP_ITEMS", frmoptions.moc, PathServ & "\Data.ini")
    Call WriteINI("MAX", "MAX_GUILDS", frmoptions.mg, PathServ & "\Data.ini")
    Call WriteINI("MAX", "MAX_GUILD_MEMBERS", frmoptions.mjg, PathServ & "\Data.ini")
    Call WriteINI("MAX", "MAX_EMOTICONS", frmoptions.Me, PathServ & "\Data.ini")
    Call WriteINI("MAX", "MAX_LEVEL", frmoptions.mn, PathServ & "\Data.ini")
    Call WriteINI("MAX", "MAX_QUETES", frmoptions.mq, PathServ & "\Data.ini")
    If HORS_LIGNE = 0 Then Call SendMOTDChange(frmoptions.motd.Text)
    
    Call SendData("CHGCLASSES" & END_CHAR)

    Call Unload(frmmsg)
End If
End Sub

Public Sub ShopEditorInit()
Dim i As Long

    frmShopEditor.txtName.Text = Trim$(Shop(EditorIndex).name)
    frmShopEditor.txtJoinSay.Text = Trim$(Shop(EditorIndex).JoinSay)
    frmShopEditor.txtLeaveSay.Text = Trim$(Shop(EditorIndex).LeaveSay)
    frmShopEditor.chkFixesItems.value = Shop(EditorIndex).FixesItems
    
    frmShopEditor.cmbItemGive.Clear
    frmShopEditor.cmbItemGive.AddItem "Aucun"
    frmShopEditor.cmbItemGet.Clear
    frmShopEditor.cmbItemGet.AddItem "Aucun"
    frmShopEditor.cmbItemFix.Clear
    frmShopEditor.cmbItemFix.AddItem "Aucun"
    For i = 1 To MAX_ITEMS
        frmShopEditor.cmbItemGive.AddItem i & " : " & Trim$(Item(i).name)
        frmShopEditor.cmbItemGet.AddItem i & " : " & Trim$(Item(i).name)
        frmShopEditor.cmbItemFix.AddItem i & " : " & Trim$(Item(i).name)
    Next i
    frmShopEditor.cmbItemGive.ListIndex = 0
    frmShopEditor.cmbItemGet.ListIndex = 0
    frmShopEditor.cmbItemFix.ListIndex = Shop(EditorIndex).FixObjet
    
    
    Call UpdateShopTrade
    
    frmShopEditor.Show vbModeless, frmMirage
End Sub

Public Sub UpdateShopTrade()
Dim i As Long, GetItem As Long, GetValue As Long, GiveItem As Long, GiveValue As Long, c As Long
    
    For i = 0 To 5
        frmShopEditor.lstTradeItem(i).Clear
    Next i
    
    For c = 1 To 6
        For i = 1 To MAX_TRADES
            GetItem = Shop(EditorIndex).TradeItem(c).value(i).GetItem
            GetValue = Shop(EditorIndex).TradeItem(c).value(i).GetValue
            GiveItem = Shop(EditorIndex).TradeItem(c).value(i).GiveItem
            GiveValue = Shop(EditorIndex).TradeItem(c).value(i).GiveValue

            If GetItem > 0 And GiveItem > 0 Then
                frmShopEditor.lstTradeItem(c - 1).AddItem i & " : " & GiveValue & " " & Trim$(Item(GiveItem).name) & " pour " & GetValue & " " & Trim$(Item(GetItem).name)
            Else
                frmShopEditor.lstTradeItem(c - 1).AddItem "Slot vide"
            End If
        Next i
    Next c
    
    For i = 0 To 5
        frmShopEditor.lstTradeItem(i).ListIndex = 0
    Next i
End Sub

Public Sub ShopEditorOk()
    Shop(EditorIndex).name = frmShopEditor.txtName.Text
    Shop(EditorIndex).JoinSay = frmShopEditor.txtJoinSay.Text
    Shop(EditorIndex).LeaveSay = frmShopEditor.txtLeaveSay.Text
    Shop(EditorIndex).FixesItems = frmShopEditor.chkFixesItems.value
    If Shop(EditorIndex).FixesItems >= 1 Then Shop(EditorIndex).FixObjet = frmShopEditor.cmbItemFix.ListIndex Else Shop(EditorIndex).FixObjet = -1
    
    Call SendSaveShop(EditorIndex)
    Unload frmShopEditor
    frmIndex.SetFocus
End Sub

Public Sub ShopEditorCancel()
Dim i As Long
    Unload frmShopEditor
    frmIndex.lstIndex.Clear
    ' Add the names
    For i = 1 To MAX_SHOPS
        frmIndex.lstIndex.AddItem i & " : " & Trim$(Shop(i).name)
    Next i
    frmIndex.SetFocus
End Sub

Public Sub SpellEditorInit()
Dim i As Long
If Not FileExiste("spells\spells" & EditorIndex & ".fcg") And HORS_LIGNE = 1 Then Call ClearSpell(EditorIndex)

    EditorItemY = (Spell(EditorIndex).SpellIco \ 6)
    EditorItemX = (Spell(EditorIndex).SpellIco - (Spell(EditorIndex).SpellIco \ 6) * 6)
    
    frmSpellEditor.cmbClassReq.AddItem "Toutes les classes"
    For i = 0 To Max_Classes
        If HORS_LIGNE = 1 Then frmSpellEditor.cmbClassReq.AddItem Trim$("classe" & i) Else frmSpellEditor.cmbClassReq.AddItem Trim$(Class(i).name)
    Next i
    
    frmSpellEditor.txtName.Text = Trim$(Spell(EditorIndex).name)
    frmSpellEditor.cmbClassReq.ListIndex = Spell(EditorIndex).ClassReq
    frmSpellEditor.scrlLevelReq.value = Spell(EditorIndex).LevelReq
        
    frmSpellEditor.cmbType.ListIndex = Spell(EditorIndex).Type
    frmSpellEditor.scrlVitalMod.value = Spell(EditorIndex).Data1
    frmSpellEditor.HScroll1.value = Spell(EditorIndex).Data3
    
    frmSpellEditor.scrlCost.value = Spell(EditorIndex).MPCost
    frmSpellEditor.scrlSound.value = Spell(EditorIndex).Sound
    
    If Spell(EditorIndex).Range = 0 Then Spell(EditorIndex).Range = 1
    frmSpellEditor.scrlRange.value = Spell(EditorIndex).Range
    
    If Spell(EditorIndex).Big = 1 Then
        frmSpellEditor.CheckSpell.value = Checked
        frmSpellEditor.scrlSpellAnim.Max = MAX_DX_BIGSPELLS
        frmSpellEditor.picSpell.Width = 960
        frmSpellEditor.picSpell.Height = 960
        frmSpellEditor.picSpell.Left = 10680
        frmSpellEditor.picSpell.Top = 3540
    Else
        frmSpellEditor.CheckSpell.value = Unchecked
        frmSpellEditor.scrlSpellAnim.Max = MAX_DX_SPELLS
        frmSpellEditor.picSpell.Width = 480
        frmSpellEditor.picSpell.Height = 480
        frmSpellEditor.picSpell.Left = 10920
        frmSpellEditor.picSpell.Top = 3720
    End If
    
    frmSpellEditor.scrlSpellAnim.value = Spell(EditorIndex).SpellAnim
    frmSpellEditor.scrlSpellTime.value = Spell(EditorIndex).SpellTime
    frmSpellEditor.scrlSpellDone.value = Spell(EditorIndex).SpellDone
    
    frmSpellEditor.chkArea.value = Spell(EditorIndex).AE
    frmSpellEditor.scrlLevelReq.Max = MAX_LEVEL
    
    frmSpellEditor.Show vbModeless, frmMirage
End Sub

Public Sub QuetesEditorInit()
Dim i As Long
If Not FileExiste("quetes\quete" & EditorIndex & ".fcq") And HORS_LIGNE = 1 Then Call ClearQuete(EditorIndex)

If Not FileExiste("quetes\quete" & EditorIndex & ".fcq") And Trim$(quete(EditorIndex).nom) = vbNullString Then Call ClearQuete(EditorIndex)

    frmEditeurQuetes.Init = False
    
    frmEditeurQuetes.nom.Text = Trim$(quete(EditorIndex).nom)
    frmEditeurQuetes.tpe.ListIndex = Val(quete(EditorIndex).Type)
    frmEditeurQuetes.description.Text = Trim$(quete(EditorIndex).description)
    frmEditeurQuetes.reponse.Text = Trim$(quete(EditorIndex).reponse)
    frmEditeurQuetes.ro1.value = Val(quete(EditorIndex).Recompence.objn1)
    frmEditeurQuetes.ro2.value = Val(quete(EditorIndex).Recompence.objn2)
    frmEditeurQuetes.ro3.value = Val(quete(EditorIndex).Recompence.objn3)
    frmEditeurQuetes.rq1.value = Val(quete(EditorIndex).Recompence.objq1)
    frmEditeurQuetes.rq2.value = Val(quete(EditorIndex).Recompence.objq2)
    frmEditeurQuetes.rq3.value = Val(quete(EditorIndex).Recompence.objq3)
    frmEditeurQuetes.rexp.Text = Val(quete(EditorIndex).Recompence.exp)
    frmEditeurQuetes.cases.Text = Val(quete(EditorIndex).Case)
    
    If frmEditeurQuetes.tpe.ListIndex = QUETE_TYPE_RECUP Then
        For i = 1 To 6
            frmEditeurQuetes.frtp(i).Visible = False
        Next i
        frmEditeurQuetes.frtp(frmEditeurQuetes.tpe.ListIndex).Visible = True
        frmEditeurQuetes.tempr(frmEditeurQuetes.tpe.ListIndex).value = quete(EditorIndex).Temps
        frmEditeurQuetes.indo.value = Val(quete(EditorIndex).Data1)
        frmEditeurQuetes.numo.value = Val(quete(EditorIndex).Data2)
        frmEditeurQuetes.quant.value = Val(quete(EditorIndex).Data3)
    ElseIf frmEditeurQuetes.tpe.ListIndex = QUETE_TYPE_APORT Then
        For i = 1 To 6
            frmEditeurQuetes.frtp(i).Visible = False
        Next i
        frmEditeurQuetes.frtp(frmEditeurQuetes.tpe.ListIndex).Visible = True
        frmEditeurQuetes.tempr(frmEditeurQuetes.tpe.ListIndex).value = quete(EditorIndex).Temps
        frmEditeurQuetes.numod.value = Val(quete(EditorIndex).Data1)
        frmEditeurQuetes.numpnj.value = Val(quete(EditorIndex).Data2)
        frmEditeurQuetes.reppnj.Text = Trim$(quete(EditorIndex).String1)
        For i = 1 To 15
            quete(EditorIndex).indexe(i).Data1 = 1
            quete(EditorIndex).indexe(i).Data2 = 1
            quete(EditorIndex).indexe(i).Data3 = 1
            quete(EditorIndex).indexe(i).String1 = vbNullString
        Next i
    ElseIf frmEditeurQuetes.tpe.ListIndex = QUETE_TYPE_PARLER Then
        For i = 1 To 6
            frmEditeurQuetes.frtp(i).Visible = False
        Next i
        frmEditeurQuetes.frtp(frmEditeurQuetes.tpe.ListIndex).Visible = True
        frmEditeurQuetes.tempr(frmEditeurQuetes.tpe.ListIndex).value = quete(EditorIndex).Temps
        frmEditeurQuetes.numepnj.value = Val(quete(EditorIndex).Data1)
        For i = 1 To 15
            quete(EditorIndex).indexe(i).Data1 = 1
            quete(EditorIndex).indexe(i).Data2 = 1
            quete(EditorIndex).indexe(i).Data3 = 1
            quete(EditorIndex).indexe(i).String1 = vbNullString
        Next i
    ElseIf frmEditeurQuetes.tpe.ListIndex = QUETE_TYPE_TUER Then
        For i = 1 To 6
            frmEditeurQuetes.frtp(i).Visible = False
        Next i
        frmEditeurQuetes.frtp(frmEditeurQuetes.tpe.ListIndex).Visible = True
        frmEditeurQuetes.tempr(frmEditeurQuetes.tpe.ListIndex).value = quete(EditorIndex).Temps
        frmEditeurQuetes.indpnj.value = Val(quete(EditorIndex).Data1)
        frmEditeurQuetes.numopnj.value = Val(quete(EditorIndex).Data2)
        frmEditeurQuetes.nbt.value = Val(quete(EditorIndex).Data3)
    ElseIf frmEditeurQuetes.tpe.ListIndex = QUETE_TYPE_FINIR Then
        For i = 1 To 6
            frmEditeurQuetes.frtp(i).Visible = False
        Next i
        frmEditeurQuetes.frtp(frmEditeurQuetes.tpe.ListIndex).Visible = True
        frmEditeurQuetes.tempr(frmEditeurQuetes.tpe.ListIndex).value = quete(EditorIndex).Temps
        frmEditeurQuetes.xd.Text = Val(quete(EditorIndex).Data1)
        frmEditeurQuetes.yd.Text = Val(quete(EditorIndex).Data2)
        frmEditeurQuetes.carted.Text = Val(quete(EditorIndex).Data3)
        For i = 1 To 15
            quete(EditorIndex).indexe(i).Data1 = 1
            quete(EditorIndex).indexe(i).Data2 = 1
            quete(EditorIndex).indexe(i).Data3 = 1
            quete(EditorIndex).indexe(i).String1 = vbNullString
        Next i
    ElseIf frmEditeurQuetes.tpe.ListIndex = QUETE_TYPE_GAGNE_XP Then
        For i = 1 To 6
            frmEditeurQuetes.frtp(i).Visible = False
        Next i
        frmEditeurQuetes.frtp(frmEditeurQuetes.tpe.ListIndex).Visible = True
        frmEditeurQuetes.tempr(frmEditeurQuetes.tpe.ListIndex).value = quete(EditorIndex).Temps
        frmEditeurQuetes.nbxp.value = Val(quete(EditorIndex).Data1)
        For i = 1 To 15
            quete(EditorIndex).indexe(i).Data1 = 1
            quete(EditorIndex).indexe(i).Data2 = 1
            quete(EditorIndex).indexe(i).Data3 = 1
            quete(EditorIndex).indexe(i).String1 = vbNullString
        Next i
    ElseIf frmEditeurQuetes.tpe.ListIndex = QUETE_TYPE_SCRIPT Then
        For i = 1 To 6
            frmEditeurQuetes.frtp(i).Visible = False
        Next i
        frmEditeurQuetes.frtp(frmEditeurQuetes.tpe.ListIndex).Visible = True
        frmEditeurQuetes.tempr(frmEditeurQuetes.tpe.ListIndex).value = quete(EditorIndex).Temps
        For i = 1 To 15
            quete(EditorIndex).indexe(i).Data1 = 1
            quete(EditorIndex).indexe(i).Data2 = 1
            quete(EditorIndex).indexe(i).Data3 = 1
            quete(EditorIndex).indexe(i).String1 = vbNullString
        Next i
    End If
    frmEditeurQuetes.Label4.Caption = "Index de l'objet (pour la quête) : " & frmEditeurQuetes.indo.value
    frmEditeurQuetes.Label15.Caption = "Index du PNJ (pour la quête) : " & frmEditeurQuetes.indpnj.value
    frmEditeurQuetes.Label12.Caption = "Nombre de fois qu'il faut le tuer : " & frmEditeurQuetes.nbt.value
    frmEditeurQuetes.Label22.Caption = "Nombre de points d'experiences a gagner : " & frmEditeurQuetes.nbxp.value
    frmEditeurQuetes.Label26.Caption = "Case scripter a éxécuter : " & frmEditeurQuetes.cases.Text
    frmEditeurQuetes.Label11.Caption = "Numéro du PNJ : " & frmEditeurQuetes.numepnj.value
    frmEditeurQuetes.Label6.Caption = "Numéro de l'objet : " & frmEditeurQuetes.numo.value
    frmEditeurQuetes.Label8.Caption = "Numéro de l'objet donné : " & frmEditeurQuetes.numod.value
    frmEditeurQuetes.Label13.Caption = "Numéro du PNJ : " & frmEditeurQuetes.numopnj.value
    frmEditeurQuetes.Label9.Caption = "Numéro du PNJ : " & frmEditeurQuetes.numpnj.value
    frmEditeurQuetes.Label7.Caption = "Quantités à ramasser : " & frmEditeurQuetes.quant.value
    frmEditeurQuetes.Label16.Caption = "Numéro de l'objet1 : " & frmEditeurQuetes.ro1.value & " : " & Item(frmEditeurQuetes.ro1.value).name
    frmEditeurQuetes.Label18.Caption = "Numéro de l'objet2 : " & frmEditeurQuetes.ro2.value & " : " & Item(frmEditeurQuetes.ro2.value).name
    frmEditeurQuetes.Label23.Caption = "Numéro de l'objet3 : " & frmEditeurQuetes.ro3.value & " : " & Item(frmEditeurQuetes.ro3.value).name
    frmEditeurQuetes.Label17.Caption = "Quantités de l'objet1 : " & frmEditeurQuetes.rq1.value
    frmEditeurQuetes.Label21.Caption = "Quantités de l'objet2 : " & frmEditeurQuetes.rq2.value
    frmEditeurQuetes.Label24.Caption = "Quantités de l'objet3 : " & frmEditeurQuetes.rq3.value
    frmEditeurQuetes.Label14.Caption = "Points d'expérience gagnés : "
    For i = 1 To 6
        If frmEditeurQuetes.tempr(i).value > 0 Then frmEditeurQuetes.tp(i).Caption = "Temps pour réaliser la quête : " & frmEditeurQuetes.tempr(i).value & "s (" & (frmEditeurQuetes.tempr(i).value \ 60) & "min" & frmEditeurQuetes.tempr(i).value - ((frmEditeurQuetes.tempr(i).value \ 60) * 60) & "s)" Else frmEditeurQuetes.tp(i).Caption = "Temps pour réaliser la quête : Infini"
    Next i
    frmEditeurQuetes.numod.Max = MAX_ITEMS
    frmEditeurQuetes.numo.Max = MAX_ITEMS
    frmEditeurQuetes.numpnj.Max = MAX_NPCS
    frmEditeurQuetes.numopnj.Max = MAX_NPCS
    frmEditeurQuetes.numepnj.Max = MAX_NPCS
    frmEditeurQuetes.Caption = "Editeur de Quêtes : " & EditorIndex
    frmEditeurQuetes.Show vbModeless, frmMirage
    
    frmEditeurQuetes.Init = True
End Sub

Public Sub SpellEditorOk()
    Spell(EditorIndex).name = frmSpellEditor.txtName.Text
    Spell(EditorIndex).ClassReq = frmSpellEditor.cmbClassReq.ListIndex
    Spell(EditorIndex).LevelReq = frmSpellEditor.scrlLevelReq.value
    Spell(EditorIndex).Type = frmSpellEditor.cmbType.ListIndex
    Spell(EditorIndex).Data1 = frmSpellEditor.scrlVitalMod.value
    Spell(EditorIndex).Data3 = frmSpellEditor.HScroll1.value
    Spell(EditorIndex).MPCost = frmSpellEditor.scrlCost.value
    Spell(EditorIndex).Sound = frmSpellEditor.scrlSound.value
    Spell(EditorIndex).Range = frmSpellEditor.scrlRange.value
    Spell(EditorIndex).SpellIco = EditorItemY * 6 + EditorItemX
    
    If frmSpellEditor.CheckSpell.value = Checked Then
        Spell(EditorIndex).Big = 1
    Else
        Spell(EditorIndex).Big = 0
    End If
    
    Spell(EditorIndex).SpellAnim = frmSpellEditor.scrlSpellAnim.value
    Spell(EditorIndex).SpellTime = frmSpellEditor.scrlSpellTime.value
    Spell(EditorIndex).SpellDone = frmSpellEditor.scrlSpellDone.value
    
    Spell(EditorIndex).AE = frmSpellEditor.chkArea.value
    
    Call SendSaveSpell(EditorIndex)
    InShopEditor = False
    Call SpellEditorCancel
End Sub

Public Sub SpellEditorCancel()
Dim i As Long
    Unload frmSpellEditor
    frmIndex.lstIndex.Clear
    ' Add the names
    For i = 1 To MAX_SPELLS
        frmIndex.lstIndex.AddItem i & " : " & Trim$(Spell(i).name)
    Next i
    frmIndex.SetFocus
End Sub

Public Sub Tester()
Dim i As Long
If test = 0 Then
    Call SaveMap(Player(MyIndex).Map)
    InEditor = False
    If save = 1 Then Call MsgBox("La carte n'a pas été enregistrée sur le serveur certains attributs peuvent ne pas fonctionnés totalement et/ou les PNJ ne marcheront pas.")
    For i = 2 To 30
        If i <> 20 And i <> 21 And i <> 23 And i <> 24 And i <> 25 And i <> 27 And i <> 28 And i <> 29 Then frmMirage.Toolbar1.buttons(i).Enabled = False
    Next i
    frmMirage.carte.Enabled = False
    frmMirage.comtest.Enabled = True
    frmMirage.Editeurs.Enabled = False
    frmMirage.Toolbar1.buttons(1).Image = 19
    frmMirage.Toolbar1.buttons(1).ToolTipText = "Arreter le test de la carte"
    frmMirage.gauchedroite.Visible = False
    frmMirage.hautbas.Visible = False
    frmMirage.vie.Visible = True
    frmMirage.mana.Visible = True
    frmMirage.xp.Visible = True
    frmMirage.picScreen.SetFocus
    frmMirage.test.Caption = "Quitter le teste"
    ConOff = True
    Call SendData("refresh" & END_CHAR)
    test = 1
    Call InitNightAndFog(Player(MyIndex).Map)
Else
    Call SendData("mapreport" & END_CHAR)
    If frmMirage.tp(2).Checked Then
        For i = 2 To 22
            If (i < 4 Or i > 19) And i <> 20 And i <> 21 Then frmMirage.Toolbar1.buttons(i).Enabled = True
        Next i
    Else
        For i = 2 To 30
            If i <> 20 And i <> 21 Then frmMirage.Toolbar1.buttons(i).Enabled = True
        Next i
    End If
    frmMirage.test.Caption = "Tester"
    frmMirage.itmDesc.Visible = False
    InEditor = True
    frmMirage.scrlPicture.Max = ((DDSD_Tile(EditorSet).lHeight - frmMirage.picBackSelect.Height) \ PIC_Y)
    frmMirage.picBack.Width = frmMirage.picBackSelect.Width
    frmMirage.carte.Enabled = True
    frmMirage.comtest.Enabled = False
    frmMirage.Editeurs.Enabled = True
    frmMirage.Toolbar1.buttons(1).Image = 8
    frmMirage.Toolbar1.buttons(1).ToolTipText = "Tester la carte"
    frmMirage.gauchedroite.Visible = True
    frmMirage.hautbas.Visible = True
    frmMirage.quetetimersec.Enabled = False
    frmMirage.tmpsquete.Visible = False
    frmMirage.vie.Visible = False
    frmMirage.mana.Visible = False
    frmMirage.xp.Visible = False
    Call frmMirage.NetPic
    Call ChargerCarte(Player(MyIndex).Map)
    test = 0
    If ConOff Then ConOff = False
    Call InitNightAndFog(Player(MyIndex).Map)
End If
End Sub

Public Sub UpdateTradeInventory()
Dim i As Long
    frmPlayerTrade.PlayerInv1.Clear
    
For i = 1 To MAX_INV
    If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
            frmPlayerTrade.PlayerInv1.AddItem i & " : " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerPetSlot(MyIndex) = i Then
                frmPlayerTrade.PlayerInv1.AddItem i & " : " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (worn)"
            Else
                frmPlayerTrade.PlayerInv1.AddItem i & " : " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name)
            End If
        End If
    Else
        frmPlayerTrade.PlayerInv1.AddItem "<Rien>"
    End If
Next i
    
    frmPlayerTrade.PlayerInv1.ListIndex = 0
End Sub

Sub PlayerSearch(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1 As Long, y1 As Long

    x1 = (x \ PIC_X / VZoom * 3)
    y1 = (y \ PIC_Y / VZoom * 3)
    
    If (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then Call SendData("search" & SEP_CHAR & x1 & SEP_CHAR & y1 & END_CHAR)
    MouseDownX = x1
    MouseDownY = y1
End Sub

Sub BltTile2(ByVal x As Long, ByVal y As Long, ByVal tile As Long)
    rec.Top = (tile \ TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (tile - (tile \ TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
    Call DD_BackBuffer.BltFast(x - NewPlayerPicX + sx - NewXOffset, y - NewPlayerPicY + sy - NewYOffset, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPlayerText(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim intLoop As Integer
Dim intLoop2 As Integer

Dim bytLineCount As Byte
Dim bytLineLength As Byte
Dim strLine(0 To MAX_LINES - 1) As String
Dim strWords() As String
    strWords() = Split(Bubble(Index).Text, " ")
    
    If Len(Bubble(Index).Text) < MAX_LINE_LENGTH Then
        DISPLAY_BUBBLE_WIDTH = 2 + ((Len(Bubble(Index).Text) * 9) \ PIC_X)
        If DISPLAY_BUBBLE_WIDTH > MAX_BUBBLE_WIDTH Then DISPLAY_BUBBLE_WIDTH = MAX_BUBBLE_WIDTH
    Else
        DISPLAY_BUBBLE_WIDTH = 6
    End If
    
    TextX = GetPlayerX(Index) * PIC_X + Player(Index).XOffset + Int(PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2) - 6
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - Int(PIC_Y) + 85
    
    Call DD_BackBuffer.ReleaseDC(TexthDC)
    
    ' Draw the fancy box with tiles.
    Call BltTile2(TextX - 10, TextY - 10, 1)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY - 10, 2)
    
    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY - 10, 16)
    Next intLoop
    
    TexthDC = DD_BackBuffer.GetDC
    
    ' Loop through all the words.
    For intLoop = 0 To UBound(strWords)
        ' Increment the line length.
        bytLineLength = bytLineLength + Len(strWords(intLoop)) + 1
            
        ' If we have room on the current line.
        If bytLineLength < MAX_LINE_LENGTH Then
            ' Add the text to the current line.
            strLine(bytLineCount) = strLine(bytLineCount) & strWords(intLoop) & " "
        Else
            bytLineCount = bytLineCount + 1
            
            If bytLineCount = MAX_LINES Then bytLineCount = bytLineCount - 1: Exit For
            
            strLine(bytLineCount) = Trim$(strWords(intLoop)) & " "
            bytLineLength = 0
        End If
    Next intLoop
    
    Call DD_BackBuffer.ReleaseDC(TexthDC)
    
    If bytLineCount > 0 Then
        For intLoop = 6 To (bytLineCount - 2) * PIC_Y + 6
            Call BltTile2(TextX - 10, TextY - 10 + intLoop, 19)
            Call BltTile2(TextX - 10 + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X, TextY - 10 + intLoop, 17)
            
            For intLoop2 = 1 To DISPLAY_BUBBLE_WIDTH - 2
                Call BltTile2(TextX - 10 + (intLoop2 * PIC_X), TextY + intLoop - 10, 5)
            Next intLoop2
        Next intLoop
    End If

    Call BltTile2(TextX - 10, TextY + (bytLineCount * 16) - 4, 3)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY + (bytLineCount * 16) - 4, 4)
    
    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY + (bytLineCount * 16) - 4, 15)
    Next intLoop
    
    TexthDC = DD_BackBuffer.GetDC
    
    For intLoop = 0 To (MAX_LINES - 1)
        If strLine(intLoop) <> vbNullString Then
            Call DrawText(TexthDC, TextX - NewPlayerPicX + sx - NewXOffset + (((DISPLAY_BUBBLE_WIDTH) * PIC_X) / 2) - ((Len(strLine(intLoop)) * 8) \ 2) - 7, TextY - NewPlayerPicY + sy - NewYOffset, strLine(intLoop), QBColor(DarkGrey))
            TextY = TextY + 16
        End If
    Next intLoop
End Sub

Sub BltPlayerBar() 's(ByVal Index As Long)
Dim x As Long, y As Long, Index As Long, ty As Long

Index = MyIndex

If Player(Index).HP <> 0 Then
    ty = (DDSD_Character(GetPlayerSprite(Index)).lHeight / 4) / 2
    x = (GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset) - NewPlayerPOffsetX
    y = (GetPlayerY(Index) * PIC_Y + sy + Player(Index).YOffset) - NewPlayerPOffsetY + ty + 3
    'draws the back bars
    Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
    Call DD_BackBuffer.DrawBox(x, y + 2, x + 32, y - 2)
    'draws HP
    Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
    Call DD_BackBuffer.DrawBox(x, y + 2, x + ((Player(Index).HP / 100) / (Player(Index).MaxHp / 100) * 32), y - 2)
End If
End Sub
Sub BltNpcBars(ByVal Index As Long)
Dim x As Long, y As Long, ty As Long

If MapNpc(Index).HP <= 0 Or MapNpc(Index).MaxHp <= 0 Or MapNpc(Index).num < 1 Then Exit Sub

        ty = (DDSD_Character(Npc(MapNpc(Index).num).sprite).lHeight / 4) / 2
        x = (MapNpc(Index).x * PIC_X + sx + MapNpc(Index).XOffset) - NewPlayerPOffsetX
        y = (MapNpc(Index).y * PIC_Y + sy + MapNpc(Index).YOffset) - NewPlayerPOffsetY + ty + 3
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(x, y, x + 32, y + 4)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        Call DD_BackBuffer.DrawBox(x, y, x + ((MapNpc(Index).HP / 100) / (MapNpc(Index).MaxHp / 100) * 32), y + 4)
End Sub

Public Sub AffInv()
Dim Q As Long
Dim Qq As Long
    For Q = 0 To MAX_INV - 1
        Qq = Player(MyIndex).inv(Q + 1).num
        If Qq = 0 Then frmMirage.picInv(Q).Picture = LoadPicture() Else Call AffSurfPic(DD_ItemSurf, frmMirage.picInv(Q), (Item(Qq).Pic - (Item(Qq).Pic \ 6) * 6) * PIC_X, (Item(Qq).Pic \ 6) * PIC_Y)
    Next Q
End Sub

Public Sub UpdateVisInv()
Dim Index As Long
Dim d As Long

frmMirage.ShieldImage.Picture = LoadPicture()
frmMirage.WeaponImage.Picture = LoadPicture()
frmMirage.HelmetImage.Picture = LoadPicture()
frmMirage.ArmorImage.Picture = LoadPicture()
frmMirage.PetImage.Picture = LoadPicture()

On Error GoTo mont:
    For Index = 1 To MAX_INV
        If GetPlayerShieldSlot(MyIndex) = Index Then Call AffSurfPic(DD_ItemSurf, frmMirage.ShieldImage, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic \ 6) * 6) * PIC_X, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic \ 6) * PIC_Y)
        'Call BitBlt(frmMirage.ShieldImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerWeaponSlot(MyIndex) = Index Then Call AffSurfPic(DD_ItemSurf, frmMirage.WeaponImage, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic \ 6) * 6) * PIC_X, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic \ 6) * PIC_Y)
        'Call BitBlt(frmMirage.WeaponImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerHelmetSlot(MyIndex) = Index Then Call AffSurfPic(DD_ItemSurf, frmMirage.HelmetImage, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic \ 6) * 6) * PIC_X, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic \ 6) * PIC_Y)
        'Call BitBlt(frmMirage.HelmetImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerArmorSlot(MyIndex) = Index Then Call AffSurfPic(DD_ItemSurf, frmMirage.ArmorImage, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic \ 6) * 6) * PIC_X, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic \ 6) * PIC_Y)
        'Call BitBlt(frmMirage.ArmorImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerPetSlot(MyIndex) = Index Then Call AffSurfPic(DD_ItemSurf, frmMirage.PetImage, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic \ 6) * 6) * PIC_X, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic \ 6) * PIC_Y)
    Next Index
mont:
    frmMirage.EquipS(0).Visible = False
    frmMirage.EquipS(1).Visible = False
    frmMirage.EquipS(2).Visible = False
    frmMirage.EquipS(3).Visible = False
    frmMirage.EquipS(4).Visible = False
    
    For d = 0 To MAX_INV - 1
        If Player(MyIndex).inv(d + 1).num > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type <> ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, d + 1)).Empilable = 0 Then
                If GetPlayerWeaponSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(0).Visible = True
                    frmMirage.EquipS(0).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(0).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerArmorSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(1).Visible = True
                    frmMirage.EquipS(1).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(1).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerHelmetSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(2).Visible = True
                    frmMirage.EquipS(2).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(2).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerShieldSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(3).Visible = True
                    frmMirage.EquipS(3).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(3).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerPetSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(4).Visible = True
                    frmMirage.EquipS(4).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(4).Left = frmMirage.picInv(d).Left - 2
                End If
            End If
        End If
    Next d
    Call AffInv
End Sub

Public Sub QueteMsg(ByVal Index As Long, ByVal Msg As String)
    frmMirage.txtQ.Visible = True
    frmMirage.TxtQ2.Text = Msg
End Sub

Sub BltSpriteChange(ByVal x As Long, ByVal y As Long)
Dim x2 As Long, y2 As Long
    ' Only used if ever want to switch to blt rather then bltfast
    'With rec_pos
        '.Top = y * PIC_Y
        '.Bottom = .Top + PIC_Y
        '.Left = x * PIC_X
        '.Right = .Left + PIC_X
    'End With
                                    
    rec.Top = Map(Player(MyIndex).Map).tile(x, y).Data1 * (PIC_NPC1 * 32) + PIC_NPC2
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = 128
    rec.Right = rec.Left + PIC_X
    
    x2 = x * PIC_X + sx
    y2 = y * PIC_Y + sy
                                       
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(x2 - NewPlayerPOffsetX, y2 - NewPlayerPOffsetY, DD_SpriteSurf(GetPlayerSprite(MyIndex)), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltSpriteChange2(ByVal x As Long, ByVal y As Long)
Dim x2 As Long, y2 As Long
    ' Only used if ever want to switch to blt rather then bltfast
    'With rec_pos
        '.Top = y * PIC_Y
        '.Bottom = .Top + PIC_Y
        '.Left = x * PIC_X
        '.Right = .Left + PIC_X
    'End With
                                    
    rec.Top = Map(Player(MyIndex).Map).tile(x, y).Data1 * 64
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = 128
    rec.Right = rec.Left + PIC_X
    
    x2 = x * PIC_X + sx
    y2 = y * PIC_Y + sy - 32
    If x2 < 0 Then x2 = 0
    If y2 < 0 Then y2 = 0
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(x2 - NewPlayerPOffsetX, y2 - NewPlayerPOffsetY, DD_SpriteSurf(GetPlayerSprite(MyIndex)), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub SaveMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".fcc"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Map(MapNum)
    Close #f
End Sub

Sub SavePet(ByVal PetNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\pets\pet" & PetNum & ".fcf"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Pets(PetNum)
    Close #f
End Sub

Sub SauvTemp()
Dim i As Long
If InMouvEditor Then Exit Sub
If InQuetesEditor Then Exit Sub
If InDefTel Then Exit Sub
If InDefKey Then Exit Sub

If TempNum > 0 Then
    'TempMap(1) = TempMap(TempNum)
    For i = 0 To 5 - TempNum
        TempMap(1 + i) = TempMap(TempNum + i)
    Next i
    For i = 7 - TempNum To 5
        Call NetTempMap(i)
    Next i
Else
    TempMap(5) = TempMap(4)
    TempMap(4) = TempMap(3)
    TempMap(3) = TempMap(2)
    TempMap(2) = TempMap(1)
    TempMap(1) = Map(Player(MyIndex).Map)
End If
TempNum = 0
frmMirage.Toolbar1.buttons(20).Enabled = True
frmMirage.Toolbar1.buttons(21).Enabled = False
End Sub

Sub SauveAuto()
Dim FileName As String
Dim f As Long

    If Not IsPlaying(MyIndex) Then Exit Sub
    
    FileName = App.Path & "\Maps\map" & Player(MyIndex).Map & "BACKUP.fcc"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Map(Player(MyIndex).Map)
    Close #f
End Sub

Sub SaveMapVide(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long
Call VidercttMap(MapNum)
    FileName = App.Path & "\maps\map" & MapNum & ".fcc"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Map(MapNum)
    Close #f
End Sub

Sub SendGameTime()
Dim Packet As String

Packet = "GmTime" & SEP_CHAR & GameTime & END_CHAR
Call SendData(Packet)
End Sub

Sub ItemSelected(ByVal Index As Long, ByVal Selected As Long)
Dim index2 As Long
index2 = Trade(Selected).Items(Index).ItemGetNum

    frmTrade.shpSelect.Top = frmTrade.picItem(Index - 1).Top - 1
    frmTrade.shpSelect.Left = frmTrade.picItem(Index - 1).Left - 1

    If index2 <= 0 Then Call ClearItemSelected: Exit Sub

    frmTrade.descName.Caption = Trim$(Item(index2).name)
    frmTrade.descName.ForeColor = Item(index2).NCoul
    frmTrade.descQuantity.Caption = Trade(Selected).Items(Index).ItemGetVal
    
    frmTrade.descStr.Caption = Item(index2).StrReq
    frmTrade.descDef.Caption = Item(index2).DefReq
    frmTrade.descSpeed.Caption = Item(index2).SpeedReq
    If Item(index2).Type = ITEM_TYPE_SPELL Then
        If Spell(Item(index2).Data1).ClassReq = 0 Then frmTrade.descClasse.Caption = "Toute" Else frmTrade.descClasse.Caption = Class(Spell(Item(index2).Data1).ClassReq - 1).name
    Else
        If Item(index2).ClassReq = -1 Then frmTrade.descClasse.Caption = "Toute" Else frmTrade.descClasse.Caption = Class(Item(index2).ClassReq).name
    End If
    
    frmTrade.descAStr.Caption = Item(index2).AddStr
    frmTrade.descADef.Caption = Item(index2).AddDef
    frmTrade.descAMagi.Caption = Item(index2).AddMagi
    frmTrade.descASpeed.Caption = Item(index2).AddSpeed
    
    frmTrade.descHp.Caption = Item(index2).AddHP
    frmTrade.descMp.Caption = Item(index2).AddMP
    frmTrade.descSp.Caption = Item(index2).AddSP

    frmTrade.descAExp.Caption = Item(index2).AddEXP
    frmTrade.desc.Caption = Trim$(Item(index2).desc)
    
    frmTrade.lblTradeFor.Caption = Trim$(Item(Trade(Selected).Items(Index).ItemGiveNum).name)
    frmTrade.lblTradeFor.ForeColor = Item(Trade(Selected).Items(Index).ItemGiveNum).NCoul
    frmTrade.lblQuantity.Caption = Trade(Selected).Items(Index).ItemGiveVal
End Sub

Sub ClearItemSelected()
    frmTrade.lblTradeFor.Caption = vbNullString
    frmTrade.lblQuantity.Caption = vbNullString
    
    frmTrade.descName.Caption = vbNullString
    frmTrade.descQuantity.Caption = vbNullString
    
    frmTrade.descStr.Caption = 0
    frmTrade.descDef.Caption = 0
    frmTrade.descSpeed.Caption = 0
    frmTrade.descClasse.Caption = vbNullString
    
    frmTrade.descAStr.Caption = 0
    frmTrade.descADef.Caption = 0
    frmTrade.descAMagi.Caption = 0
    frmTrade.descASpeed.Caption = 0
    
    frmTrade.descHp.Caption = 0
    frmTrade.descMp.Caption = 0
    frmTrade.descSp.Caption = 0

    frmTrade.descAExp.Caption = 0
    frmTrade.desc.Caption = vbNullString
End Sub

Sub AffTilesPic(ByVal Tnum As Byte, ByVal AScr As Long)
Dim sRECT As RECT
Dim dRECT As RECT
    frmMirage.picBackSelect.Picture = LoadPicture()
    frmMirage.picBackSelect.Width = Int(DDSD_Tile(Tnum).lWidth)
    frmMirage.scrlPicture.Max = Int((DDSD_Tile(Tnum).lHeight - frmMirage.picBackSelect.Height) \ PIC_Y)
    frmMirage.picBack.Width = Int(frmMirage.picBackSelect.Width)
    With dRECT
        .Top = 0
        .Bottom = frmMirage.picBackSelect.Height
        .Left = 0
        .Right = frmMirage.picBackSelect.Width
    End With
    With sRECT
        .Top = AScr
        .Bottom = .Top + frmMirage.picBackSelect.Height
        .Left = 0
        .Right = frmMirage.picBackSelect.Width
    End With
    Call DD_TileSurf(Tnum).BltToDC(frmMirage.picBackSelect.hDC, sRECT, dRECT)
    frmMirage.picBackSelect.Refresh
End Sub

Sub AffOutilPic(ByVal AScr As Long)
Dim sRECT As RECT
Dim dRECT As RECT
    frmMirage.picBackSelect.Picture = LoadPicture()
    frmMirage.picBackSelect.Width = Int(DDSD_Outil.lWidth)
    frmMirage.scrlPicture.Max = Int((DDSD_Outil.lHeight - frmMirage.picBackSelect.Height) \ PIC_Y)
    frmMirage.picBack.Width = Int(frmMirage.picBackSelect.Width)
    With dRECT
        .Top = 0
        .Bottom = frmMirage.picBackSelect.Height
        .Left = 0
        .Right = frmMirage.picBackSelect.Width
    End With
    With sRECT
        .Top = AScr
        .Bottom = .Top + frmMirage.picBackSelect.Height
        .Left = 0
        .Right = frmMirage.picBackSelect.Width
    End With
    Call DD_OutilSurf.BltToDC(frmMirage.picBackSelect.hDC, sRECT, dRECT)
    frmMirage.picBackSelect.Refresh
End Sub

Sub AffSurfPic(ByVal DD_Surf As DirectDrawSurface7, ByVal PicBox As PictureBox, ByVal x As Long, ByVal y As Long)
On Error Resume Next
Dim sRECT As RECT
Dim dRECT As RECT
If Not (DD_Surf Is Nothing) Then
    If DD_Surf Is Nothing Then Exit Sub
    PicBox.Picture = LoadPicture()
    With dRECT
        .Top = 0
        .Bottom = PicBox.Height
        .Left = 0
        .Right = PicBox.Width
    End With
    With sRECT
        .Top = y
        .Bottom = .Top + PicBox.Height
        .Left = x
        .Right = .Left + PicBox.Width
    End With
    Call DD_Surf.BltToDC(PicBox.hDC, sRECT, dRECT)
    PicBox.Refresh
    End If
End Sub

Sub NetInEditor()
InItemsEditor = False
InNpcEditor = False
InShopEditor = False
InSpellEditor = False
InEmoticonEditor = False
InArrowEditor = False
InQuetesEditor = False
InMetierEditor = False
InRecetteEditor = False
End Sub

Sub CheckErr()
Dim MapErr As Long
    MapErr = Val(ReadINI("CONFIG", "ERR", App.Path & "\Config.ini"))
    If FileExistes(App.Path & "\Maps\map" & MapErr & "BACKUP.fcc") Then Call frmMapErr.Init(MapErr)
End Sub
