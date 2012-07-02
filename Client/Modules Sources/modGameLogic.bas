Attribute VB_Name = "modGameLogic"

'***************************************************************************************************************************************************'
'ATTENTION : PENSER A NOTER LES MODIFICATIONS QUE VOUS APPORTER AU SOURCES POUR POUVOIR LES REFAIRE PLUS TARD SI VOUS DESIRER ACTUALISER LES SOURCES'
'***************************************************************************************************************************************************'

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
Public Const MENU_STATE_NEWACCOUNT As Byte = 0
Public Const MENU_STATE_DELACCOUNT As Byte = 1
Public Const MENU_STATE_LOGIN As Byte = 2
Public Const MENU_STATE_GETCHARS As Byte = 3
Public Const MENU_STATE_NEWCHAR As Byte = 4
Public Const MENU_STATE_ADDCHAR As Byte = 5
Public Const MENU_STATE_DELCHAR As Byte = 6
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

' Used to check if in Toit or not
Public InToit As Boolean

' Map for local use
Public SaveMapItem() As MapItemRec
Public SaveMapNpc(1 To MAX_MAP_NPCS) As MapNpcRec

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
Public DmgAddRem As Long
Public NPCDmgAddRem As Long

Public ii As Long, iii As Long
Public sx As Long
Public sy As Long

Public MouseDownX As Long
Public MouseDownY As Long

Public SpritePic As Long
Public SpriteItem As Long
Public SpritePrice As Long

Public SoundFileName As String

Public Connucted As Boolean

'Pour la banque
Public bankmsg As String

'Pour les quetes
Public Accepter As Boolean

'Pour les controlles
Public ConOff As Boolean
Public OldMap As Long
Public Rep_Theme As String
Public NumShop As Long

'Pour le mouvement des fenetre
Public drx As Long
Public dry As Long
Public dr As Boolean

Public cychat As Integer

'Pour les couleurs personalisables
Public AccModo As Long
Public AccMapeur As Long
Public AccDevelopeur As Long
Public AccAdmin As Long

'Mouvement des PNJs
Public PNJAnim(1 To MAX_MAP_NPCS) As Byte

'Variables pour FrmMirage
Public PicScWidth As Single
Public PicScHeight As Single

Public MaxSprite As Integer
Public MaxPaperdoll As Integer
Public MaxSpell As Integer
Public MaxBigSpell As Integer
Public MaxPet As Integer

                    
Sub Main()
Dim i As Long
Dim Ending As String
Dim t As Currency

On Error GoTo er:
    
    If Not FileExiste("Client.exe.manifest") Then
    Call URLDownloadToFile(0, "http://frogcreator.fr/update/Client.exe.manifest", App.Path & "\Client.exe.manifest", 0, 0)
    End If
    Call WriteINI("CONFIG", "Version", App.Minor & "." & App.Minor & "." & App.Revision, App.Path & "\Client\Config\Client.ini")
    Call InitXpStyle
    Rep_Theme = ReadINI("Themes", "Theme", App.Path & "\Themes.ini")
    dr = False
    frmsplash.Visible = True
    Call SetStatus("Vérification des dossiers...")
    DoEvents
        
    If FileExiste("GFX\curseur.cur") Then Call frmMainMenu.imgl.ListImages.Add(1, , LoadPNG(App.Path & "\GFX\curseur.png")): frmMainMenu.MouseIcon = frmMainMenu.imgl.ListImages(1).ExtractIcon: frmMainMenu.MousePointer = 99: frmMirage.MouseIcon = frmMainMenu.imgl.ListImages(1).ExtractIcon: frmMirage.MousePointer = 99: frmNewChar.MouseIcon = frmMainMenu.imgl.ListImages(1).ExtractIcon: frmNewChar.MousePointer = 99
    frmsplash.Shape1.Width = frmsplash.Shape1.Width + 200
    
  '  For i = 0 To 256
  '      If Not FileExiste("GFX\Tiles" & i & ".png") Then ExtraSheets = i - 1: Exit For
  '  Next i
    
    i = 0
    Do While FileExiste("GFX\Tiles" & i & ".png")
    ExtraSheets = i
    i = i + 1
    Loop
    i = 0
    
    If ExtraSheets < 0 Then ExtraSheets = 0
    
    ReDim DD_TileSurf(0 To ExtraSheets) As DirectDrawSurface7
    ReDim DDSD_Tile(0 To ExtraSheets) As DDSURFACEDESC2
    ReDim TileFile(0 To ExtraSheets) As Boolean
    
    MaxSprite = LoadMaxSprite()
    ReDim DD_SpriteSurf(0 To MaxSprite) As DirectDrawSurface7
    ReDim DDSD_Character(0 To MaxSprite) As DDSURFACEDESC2
    ReDim SpriteTimer(0 To MaxSprite) As Long
    ReDim SpriteUsed(0 To MaxSprite) As Boolean
    
    MaxPaperdoll = LoadMaxPaperdolls()
    ReDim DD_PaperDollSurf(0 To MaxPaperdoll) As DirectDrawSurface7
    ReDim DDSD_PaperDoll(0 To MaxPaperdoll) As DDSURFACEDESC2
    ReDim PaperDollTimer(0 To MaxPaperdoll) As Long
    ReDim PaperDollUsed(0 To MaxPaperdoll) As Boolean
    
    MaxSpell = LoadMaxSpells()
    ReDim DD_SpellAnim(0 To MaxSpell) As DirectDrawSurface7
    ReDim DDSD_SpellAnim(0 To MaxSpell) As DDSURFACEDESC2
    ReDim SpellTimer(0 To MaxSpell) As Long
    ReDim SpellUsed(0 To MaxSpell) As Boolean
    
    MaxBigSpell = LoadMaxBigSpells()
    ReDim DD_BigSpellAnim(0 To MaxBigSpell) As DirectDrawSurface7
    ReDim DDSD_BigSpellAnim(0 To MaxBigSpell) As DDSURFACEDESC2
    ReDim BigSpellTimer(0 To MaxBigSpell) As Long
    ReDim BigSpellUsed(0 To MaxBigSpell) As Boolean
    
    MaxPet = LoadMaxPet()
    ReDim DD_PetsSurf(0 To MaxPet) As DirectDrawSurface7
    ReDim DDSD_Pets(0 To MaxPet) As DDSURFACEDESC2
    ReDim PetTimer(0 To MaxPet) As Long
    ReDim PetUsed(0 To MaxPet) As Boolean
    frmsplash.Shape1.Width = frmsplash.Shape1.Width + 200
    ' Check if the maps directory is there, if its not make it
    If LCase$(Dir$(App.Path & "\Maps", vbDirectory)) <> "maps" Then Call MkDir$(App.Path & "\Maps")
    If UCase$(Dir$(App.Path & "\GFX", vbDirectory)) <> "GFX" Then Call MkDir$(App.Path & "\GFX")
    If UCase$(Dir$(App.Path & "\Music", vbDirectory)) <> "MUSIC" Then Call MkDir$(App.Path & "\Music")
    If UCase$(Dir$(App.Path & "\SFX", vbDirectory)) <> "SFX" Then Call MkDir$(App.Path & "\SFX")
    If UCase$(Dir$(App.Path & "\Flashs", vbDirectory)) <> "FLASHS" Then Call MkDir$(App.Path & "\Flashs")
    If UCase$(Dir$(App.Path & "\Videos", vbDirectory)) <> "VIDEOS" Then Call MkDir$(App.Path & "\Videos")
        
    Dim FileName As String
    FileName = App.Path & "\Config\Account.ini"
    If FileExiste("Config\Account.ini") Then
        With frmMirage
            .chkbubblebar.Value = ReadINI("CONFIG", "SpeechBubbles", FileName)
            .chknpcbar.Value = ReadINI("CONFIG", "NpcBar", FileName)
            .chknpcname.Value = ReadINI("CONFIG", "NPCName", FileName)
            .chkplayerbar.Value = ReadINI("CONFIG", "PlayerBar", FileName)
            .chkplayername.Value = ReadINI("CONFIG", "PlayerName", FileName)
            .chkplayerdamage.Value = ReadINI("CONFIG", "NPCDamage", FileName)
            .chknpcdamage.Value = ReadINI("CONFIG", "PlayerDamage", FileName)
            .chkmusic.Value = ReadINI("CONFIG", "Music", FileName)
            .chksound.Value = ReadINI("CONFIG", "Sound", FileName)
            .chkAutoScroll.Value = ReadINI("CONFIG", "AutoScroll", FileName)
            .chknobj.Value = Val(ReadINI("CONFIG", "NomObjet", FileName))
            .chkLowEffect.Value = Val(ReadINI("CONFIG", "LowEffect", FileName))
        End With
    Else
        WriteINI "INFO", "Account", "", App.Path & "\Config\Client.ini"
        WriteINI "INFO", "Password", "", App.Path & "\Config\Client.ini"
        WriteINI "CONFIG", "WebSite", "www.frogcreator.new.fr", App.Path & "\Config\Client.ini"
        WriteINI "CONFIG", "Version", "0.4", App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine1", "", App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine2", "", App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine3", "", App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine4", "", App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine5", "", App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine6", "", App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine7", "", App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine8", "", App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine9", "", App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine10", "", App.Path & "\Config\Client.ini"
        WriteINI "CREDIT", "CreditLine11", "", App.Path & "\Config\Client.ini"
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
    
    frmsplash.Shape1.Width = frmsplash.Shape1.Width + 200
    
    If Not FileExiste("Config\Ecriture.ini") Then
        WriteINI "POLICE", "Police", "MS Sans Serif", App.Path & "\Config\Ecriture.ini"
        WriteINI "POLICE", "PoliceSize", "8", App.Path & "\Config\Ecriture.ini"
        WriteINI "POLICE", "PoliceChat", "MS Sans Serif", App.Path & "\Config\Ecriture.ini"
        WriteINI "POLICE", "PoliceChatSize", "8", App.Path & "\Config\Ecriture.ini"
    
        WriteINI "CHATBOX", "R", 152, App.Path & "\Config\Ecriture.ini"
        WriteINI "CHATBOX", "G", 146, App.Path & "\Config\Ecriture.ini"
        WriteINI "CHATBOX", "B", 120, App.Path & "\Config\Ecriture.ini"
        
        WriteINI "CHATTEXTBOX", "R", 152, App.Path & "\Config\Ecriture.ini"
        WriteINI "CHATTEXTBOX", "G", 146, App.Path & "\Config\Ecriture.ini"
        WriteINI "CHATTEXTBOX", "B", 120, App.Path & "\Config\Ecriture.ini"
        
        WriteINI "BACKGROUND", "R", 152, App.Path & "\Config\Ecriture.ini"
        WriteINI "BACKGROUND", "G", 146, App.Path & "\Config\Ecriture.ini"
        WriteINI "BACKGROUND", "B", 120, App.Path & "\Config\Ecriture.ini"
        
        WriteINI "SPELLLIST", "R", 152, App.Path & "\Config\Ecriture.ini"
        WriteINI "SPELLLIST", "G", 146, App.Path & "\Config\Ecriture.ini"
        WriteINI "SPELLLIST", "B", 120, App.Path & "\Config\Ecriture.ini"

        WriteINI "WHOLIST", "R", 152, App.Path & "\Config\Ecriture.ini"
        WriteINI "WHOLIST", "G", 146, App.Path & "\Config\Ecriture.ini"
        WriteINI "WHOLIST", "B", 120, App.Path & "\Config\Ecriture.ini"
        
        WriteINI "NEWCHAR", "R", 152, App.Path & "\Config\Ecriture.ini"
        WriteINI "NEWCHAR", "G", 146, App.Path & "\Config\Ecriture.ini"
        WriteINI "NEWCHAR", "B", 120, App.Path & "\Config\Ecriture.ini"
        
        WriteINI "BARE", "R", 128, App.Path & "\Config\Ecriture.ini"
        WriteINI "BARE", "G", 128, App.Path & "\Config\Ecriture.ini"
        WriteINI "BARE", "B", 255, App.Path & "\Config\Ecriture.ini"
    End If
    
    Dim R1 As Long, G1 As Long, B1 As Long
    R1 = Val(ReadINI("CHATTEXTBOX", "R", App.Path & "\Config\Ecriture.ini"))
    G1 = Val(ReadINI("CHATTEXTBOX", "G", App.Path & "\Config\Ecriture.ini"))
    B1 = Val(ReadINI("CHATTEXTBOX", "B", App.Path & "\Config\Ecriture.ini"))
    frmMirage.txtMyTextBox.BackColor = RGB(R1, G1, B1)
       
    R1 = Val(ReadINI("FOND", "R", App.Path & Rep_Theme & "\Couleur.ini"))
    G1 = Val(ReadINI("FOND", "V", App.Path & Rep_Theme & "\Couleur.ini"))
    B1 = Val(ReadINI("FOND", "B", App.Path & Rep_Theme & "\Couleur.ini"))
    With frmMirage
        .Picture9.BackColor = RGB(R1, G1, B1)
        .Picture8.BackColor = RGB(R1, G1, B1)
        .Picture11.BackColor = RGB(R1, G1, B1)
        .Picture13.BackColor = RGB(R1, G1, B1)
        .picInv3.BackColor = RGB(R1, G1, B1)
        .itmDesc.BackColor = RGB(R1, G1, B1)
        .picWhosOnline.BackColor = RGB(R1, G1, B1)
        .picGuildAdmin.BackColor = RGB(R1, G1, B1)
        .picGuild.BackColor = RGB(R1, G1, B1)
        .picEquip.BackColor = RGB(R1, G1, B1)
        .picPlayerSpells.BackColor = RGB(R1, G1, B1)
        .picOptions.BackColor = RGB(R1, G1, B1)
        .pictTouche.BackColor = RGB(R1, G1, B1)
        .chkbubblebar.BackColor = RGB(R1, G1, B1)
        .chknpcbar.BackColor = RGB(R1, G1, B1)
        .chknpcname.BackColor = RGB(R1, G1, B1)
        .chkplayerbar.BackColor = RGB(R1, G1, B1)
        .chkplayername.BackColor = RGB(R1, G1, B1)
        .chkplayerdamage.BackColor = RGB(R1, G1, B1)
        .chknpcdamage.BackColor = RGB(R1, G1, B1)
        .chkmusic.BackColor = RGB(R1, G1, B1)
        .chksound.BackColor = RGB(R1, G1, B1)
        .chkAutoScroll.BackColor = RGB(R1, G1, B1)
        .chknobj.BackColor = RGB(R1, G1, B1)
        .chkLowEffect.BackColor = RGB(R1, G1, B1)
    End With
    
    frmsplash.Shape1.Width = frmsplash.Shape1.Width + 200
        
    R1 = Val(ReadINI("WHOLIST", "R", App.Path & "\Config\Ecriture.ini"))
    G1 = Val(ReadINI("WHOLIST", "G", App.Path & "\Config\Ecriture.ini"))
    B1 = Val(ReadINI("WHOLIST", "B", App.Path & "\Config\Ecriture.ini"))
    frmMirage.lstOnline.BackColor = RGB(R1, G1, B1)

    R1 = Val(ReadINI("NEWCHAR", "R", App.Path & "\Config\Ecriture.ini"))
    G1 = Val(ReadINI("NEWCHAR", "G", App.Path & "\Config\Ecriture.ini"))
    B1 = Val(ReadINI("NEWCHAR", "B", App.Path & "\Config\Ecriture.ini"))
    frmNewChar.optMale.BackColor = RGB(R1, G1, B1)
    frmNewChar.optFemale.BackColor = RGB(R1, G1, B1)
    
    Call SetStatus("Vérification du Statut...")
    frmsplash.Shape1.Width = frmsplash.Shape1.Width + 200
        
    If Not FileExiste("Config\Serveur.ini") Then
        WriteINI "SERVER0", "Name", "Server 0", App.Path & "\Config\Serveur.ini"
        WriteINI "SERVER0", "IP", "127.0.0.1", App.Path & "\Config\Serveur.ini"
        WriteINI "SERVER0", "Port", "4000", App.Path & "\Config\Serveur.ini"
    End If
    frmsplash.Visible = True
    
    Call SetStatus("Initialisation des mises à jours...")
    If Not FileExiste("Config\Updater.ini") Then
        WriteINI "UPDATER", "WebSite", "http://roonline.free.fr/patch/", App.Path & "\Config\Updater.ini"
        WriteINI "UPDATER", "WebNews", "http://roonline.free.fr/patch/patch.html", App.Path & "\Config\Updater.ini"
        WriteINI "VERSION", "Version", "0.1", App.Path & "\Config\info.ini"
    End If
    frmsplash.Shape1.Width = frmsplash.Shape1.Width + 200
    Call InitAccountOpt
    Call InitMirageVars
    'On initialise dès maintenant DirectX
    Call SetStatus("Initialisation de DirectX...")
    Call InitDirectX
    frmsplash.Shape1.Width = frmsplash.Shape1.Width + 200
    Call SetStatus("Initialisation du protocole TCP...")
        
    frmsplash.Shape1.Width = frmsplash.Shape1.Width + 400

    Call TcpInit
    
    If ReadINI("UPDATER", "actif", App.Path & "\Config\Updater.ini") = "1" And ReadINI("UPDATER", "Fin", App.Path & "\Config\Updater.ini") = "0" And ReadINI("UPDATER", "up", App.Path & "\Config\Updater.ini") = "0" Then
        WriteINI "UPDATER", "up", "1", App.Path & "\Config\Updater.ini"
        frmsplash.Visible = False
        Call Shell(App.Path & "\Updater.exe", vbNormalFocus)
        DoEvents
        frmsplash.Visible = False
        Call StopMidi
        Call GameDestroy
    Else
        WriteINI "UPDATER", "up", "0", App.Path & "\Config\Updater.ini"
        frmsplash.SetFocus
        frmServerChooser.Visible = True
        frmsplash.Visible = False
        'On initialise dès maintenant les surfaces
        Call InitSurfaces
        DoEvents
    End If
    
    ConOff = False
    frmsplash.Visible = False

Exit Sub
er:
Call MsgBox("Une erreur d'initialisation du logiciel c'est produite(Numéros de l'erreur : " & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "). Si le probléme pérsiste veulliez contacter un administrateur.", vbCritical, "Erreur")
Call GameDestroy
End Sub

Sub SetStatus(ByVal Caption As String)
    frmsplash.lblStatus.Caption = Caption
End Sub

Sub MenuState(ByVal State As Long)
    Connucted = True
    frmsplash.Visible = True
    frmsplash.Shape1.Width = 255
    Call SetStatus("Connection au Serveur...")
    Select Case State
        Case MENU_STATE_NEWACCOUNT
            frmMainMenu.fraNewAccount.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connecté, Envoie des informations du compte en cours..")
                Call SendNewAccount(frmMainMenu.txtname2.Text, frmMainMenu.txtpassword22.Text)
                Exit Sub
            End If
            
        'Case MENU_STATE_DELACCOUNT
            'frmDeleteAccount.Visible = False
         '   If ConnectToServer = True Then
          '      Call SetStatus("Connecté, Envoie de la requête d'éffacement du compte..")
                'Call SendDelAccount(frmDeleteAccount.txtName.Text, frmDeleteAccount.txtPassword.Text)
           ' End If
        
        Case MENU_STATE_LOGIN
            frmMainMenu.fraLogin.Visible = False
            If ConnectToServer = True Then Call SetStatus("Connecté, Envoie de la connexion au compte.."): Call SendLogin(frmMainMenu.txtName.Text, frmMainMenu.txtPassword.Text)
        
        Case MENU_STATE_NEWCHAR
            frmMainMenu.fraPers.Visible = False
            If ConnectToServer = True Then Call SetStatus("Connecté, Recherche des classes disponibles.."): Call SendGetClasses
            
        Case MENU_STATE_ADDCHAR
            frmNewChar.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connecté, envoie des information additionnel du personnages..")
                If frmNewChar.optMale.Value = True Then
                    Call SendAddChar(frmNewChar.txtName, 0, frmNewChar.cmbClass.ItemData(frmNewChar.cmbClass.ListIndex), frmMainMenu.lstChars.ListIndex + 1)
                Else
                    Call SendAddChar(frmNewChar.txtName, 1, frmNewChar.cmbClass.ItemData(frmNewChar.cmbClass.ListIndex), frmMainMenu.lstChars.ListIndex + 1)
                End If
            End If
        
        Case MENU_STATE_DELCHAR
            frmMainMenu.fraPers.Visible = False
            If ConnectToServer = True Then Call SetStatus("Connecté, envoie des information sur la requête d'éffacement du personnage..."): Call SendDelChar(frmMainMenu.lstChars.ListIndex + 1)
            
        Case MENU_STATE_USECHAR
            frmMainMenu.fraPers.Visible = False
            If ConnectToServer = True Then
                Call StopMidi
                Call SetStatus("Patience...")
                Call SendUseChar(frmMainMenu.lstChars.ListIndex + 1)
            End If
    End Select

    If Not IsConnected And Connucted = True Then
        frmMainMenu.Visible = True
        frmsplash.Visible = False
        Call MsgBox("Désolé, le serveur semble être indisponible, réessayer dans quelque minute ou visiter" & WEBSITE, vbOKOnly, GAME_NAME)
    End If
End Sub
Sub GameInit()
Dim i As Integer, x As Integer
    Call StopMidi
    
    If netbook Then
    frmMirage.Top = frmMainMenu.Top
    frmMirage.Left = frmMainMenu.Left
    End If
    
    frmMirage.Visible = True
    frmMainMenu.Visible = False
    frmsplash.Visible = False
    
    ' Initialize all surfaces
    'Call InitSurfaces
    
    frmMirage.picScreen.Visible = True
    Call initRac
    frmMirage.Show
End Sub

Sub initRac()
Dim i As Integer
    If LCase$(Dir$(App.Path & "\Config\Temps", vbDirectory)) <> "temps" Then Call MkDir$(App.Path & "\Config\Temps")
    For i = 0 To 13
        frmMirage.picRac(i).Picture = LoadPicture()
        rac(i, 0) = ReadINI("RAC_" & GetPlayerName(MyIndex), "rac" & i, App.Path & "\Config\Temps\" & GetPlayerName(MyIndex) & ".ini")
        rac(i, 1) = ReadINI("RAC_" & GetPlayerName(MyIndex), "type" & i, App.Path & "\Config\Temps\" & GetPlayerName(MyIndex) & ".ini")
    Next i
    frmMirage.Timer2.Enabled = True
End Sub
Sub affrac()
Dim i As Integer, Qq As Integer
    For i = 0 To 13
        If Val(rac(i, 0)) > 0 Then
            If Val(rac(i, 1)) = 1 Then
                Qq = Player(MyIndex).Spell(Val(rac(i, 0)))
            ElseIf Val(rac(i, 1)) = 2 Then
                Qq = Player(MyIndex).Inv(Val(rac(i, 0))).num
            End If
            
            If Qq = 0 Then
                frmMirage.picRac(i).Picture = LoadPicture()
            Else
                If Val(rac(i, 1)) = 1 Then
                    Call AffSurfPic(DD_ItemSurf, frmMirage.picRac(i), (Spell(Qq).SpellIco - (Spell(Qq).SpellIco \ 6) * 6) * PIC_X, (Spell(Qq).SpellIco \ 6) * PIC_Y)
                ElseIf Val(rac(i, 1)) = 2 Then
                    Call AffSurfPic(DD_ItemSurf, frmMirage.picRac(i), (Item(Qq).Pic - (Item(Qq).Pic \ 6) * 6) * PIC_X, (Item(Qq).Pic \ 6) * PIC_Y)
                Else
                    frmMirage.picRac(i).Picture = LoadPicture()
                End If
            End If
        End If
    Next i
End Sub

Sub saveRac()
Dim i As Integer
    For i = 0 To 13
        Call WriteINI("RAC_" & GetPlayerName(MyIndex), "rac" & i, rac(i, 0), App.Path & "\Config\Temps\" & GetPlayerName(MyIndex) & ".ini")
        Call WriteINI("RAC_" & GetPlayerName(MyIndex), "type" & i, rac(i, 1), App.Path & "\Config\Temps\" & GetPlayerName(MyIndex) & ".ini")
    Next i
End Sub

Sub useRac(Index As Integer)
Dim d As Byte
    If rac(Index, 0) <> "" And rac(Index, 0) <> "0" Then
        If rac(Index, 1) = "1" Then
            If Player(MyIndex).Spell(Val(rac(Index, 0))) > 0 Then
                If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
                    If Player(MyIndex).Moving = 0 Then
                        Call SendData("cast" & SEP_CHAR & Val(rac(Index, 0)) & END_CHAR)
                        Player(MyIndex).Attacking = 1
                        Player(MyIndex).AttackTimer = GetTickCount
                        Player(MyIndex).CastedSpell = YES
                    Else
                        Call AddText("Vous ne pouvez lancer un sort en marchant.", BrightRed)
                    End If
                End If
            Else
                Call AddText("Il n'y a aucun sort ici.", BrightRed)
            End If
        End If
        
        If rac(Index, 1) = "2" Then
            If Player(MyIndex).Inv(rac(Index, 0)).num <= 0 Or Player(MyIndex).Inv(rac(Index, 0)).num > MAX_ITEMS Then Exit Sub
    
            Call SendUseItem(rac(Index, 0))
            
            For d = 1 To MAX_INV
                If Player(MyIndex).Inv(d).num > 0 Then
                    If Item(GetPlayerInvItemNum(MyIndex, d)).Type = ITEM_TYPE_POTIONADDMP Or ITEM_TYPE_POTIONADDHP Or ITEM_TYPE_POTIONADDSP Or ITEM_TYPE_POTIONSUBHP Or ITEM_TYPE_POTIONSUBMP Or ITEM_TYPE_POTIONSUBSP Then frmMirage.picInv(d - 1).Picture = LoadPicture()
                End If
            Next d
            Call UpdateVisInv
        End If
    Else
        Call AddText("Il n'y a aucun raccourci ici.", BrightRed)
    End If
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
Dim screen_xg As Integer 'Nb de cases a gauche du "milieu" de picscreen
Dim screen_xd As Integer 'Nb de cases a droite du "milieu" de picscreen
Dim screen_yh As Integer 'Nb de cases en haut du "milieu" de picscreen
Dim screen_yb As Integer 'Nb de cases en bas du "milieu" de picscreen
Dim MaxDrawMapX As Long 'Calcul du maximum a dessiner en X
Dim MinDrawMapX As Long 'Calcul du minimum a dessiner en X
Dim MaxDrawMapY As Long 'Calcul du maximum a dessiner en Y
Dim MinDrawMapY As Long 'Calcul du minimum a dessiner en Y

'On Error GoTo er:
    If Not InGame Then Exit Sub
    
    If frmMirage.WindowState <> vbNormal Then Exit Sub
    
    ' Set the focus
    frmMirage.picScreen.SetFocus
    
    ' Modifier la police en jeu
    Call SetFont("Fixedsys", 20)
                
    ' Used for calculating fps
    TickFPS = 0
    TickMove = 0
    
    'Initialisation du RECT pour le backbuffer
    rec_back.Top = 0
    rec_back.Bottom = (MAX_MAPY + 1) * PIC_Y
    rec_back.Left = 0
    rec_back.Right = (MAX_MAPX + 1) * PIC_X
    
    'Initialisation des variables pour les limites de la "vue" du joueur
    screen_xg = (frmMirage.picScreen.Width \ 64) - 1
    screen_xd = (frmMirage.picScreen.Width \ 32) - screen_xg - 1
    screen_yh = (frmMirage.picScreen.height \ 64) - 1
    screen_yb = (frmMirage.picScreen.height \ 32) - screen_yh - 1
    
    Do While InGame
        Tick = GetTickCount
        
        ' Check to make sure they aren't trying to auto do anything
        'ne peux plus bouger si certaines frames sont visibles
        If frmMirage.txtMyTextBox.Locked = False Or frmTrade.Visible = True Or frmbank.Visible = True Or frmPlayerTrade.Visible = True Or frmFlash.Visible = True Or frmFixItem.Visible = True Or frmcraft.Visible = True Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
            ControlDown = False
            ShiftDown = False
        End If
        
        ' Check to make sure we are still connected
        InGame = IsConnected
        
        ' Check if we need to restore surfaces
        If NeedToRestoreSurfaces Then
rest:
            Do While NeedToRestoreSurfaces
                DoEvents
                Sleep 1
            Loop
            DD.RestoreAllSurfaces: Call InitBackBuffer
            DD.RestoreAllSurfaces: Call InitSurfaces
        End If
        
        If Not GettingMap Then
            sx = 32
            sy = 32
            
            'Calcul des variables pour l'affichage avec le scrolling
            If MAX_MAPX < screen_xg + screen_xd + 1 Then
                NewX = Player(MyIndex).x * PIC_X + Player(MyIndex).XOffset
                NewXOffset = 0
                NewPlayerX = 0
                sx = 0
            ElseIf Player(MyIndex).x <= screen_xg Then
                NewPlayerX = 0
                If Player(MyIndex).x = screen_xg And Player(MyIndex).Dir = DIR_LEFT Then
                    NewX = screen_xg * PIC_X
                    NewXOffset = Player(MyIndex).XOffset
                Else
                    NewX = Player(MyIndex).x * PIC_X + Player(MyIndex).XOffset
                    NewXOffset = 0
                End If
            ElseIf MAX_MAPX - Player(MyIndex).x <= screen_xd Then
                NewPlayerX = MAX_MAPX - screen_xd - screen_xg
                If MAX_MAPX - Player(MyIndex).x = screen_xd And Player(MyIndex).Dir = DIR_RIGHT Then
                    NewX = screen_xg * PIC_X
                    NewXOffset = Player(MyIndex).XOffset
                Else
                    NewX = (Player(MyIndex).x - MAX_MAPX + screen_xd + screen_xg) * PIC_X + Player(MyIndex).XOffset
                    NewXOffset = 0
                End If
            Else
                NewPlayerX = Player(MyIndex).x - screen_xg
                NewX = screen_xg * PIC_X
                NewXOffset = Player(MyIndex).XOffset
            End If
            
            If MAX_MAPY < screen_yh + screen_yb + 1 Then
                NewY = Player(MyIndex).y * PIC_Y + Player(MyIndex).YOffset
                NewYOffset = 0
                NewPlayerY = 0
                sy = 0
            ElseIf Player(MyIndex).y <= screen_yh Then
                NewPlayerY = 0
                If Player(MyIndex).y = screen_yh And Player(MyIndex).Dir = DIR_UP Then
                    NewY = screen_yh * PIC_Y
                    NewYOffset = Player(MyIndex).YOffset
                Else
                    NewY = Player(MyIndex).y * PIC_Y + Player(MyIndex).YOffset
                    NewYOffset = 0
                End If
            ElseIf MAX_MAPY - Player(MyIndex).y <= screen_yb Then
                NewPlayerY = MAX_MAPY - screen_yb - screen_yh
                If MAX_MAPY - Player(MyIndex).y = screen_yb And Player(MyIndex).Dir = DIR_DOWN Then
                    NewY = screen_yh * PIC_Y
                    NewYOffset = Player(MyIndex).YOffset
                Else
                    NewY = (Player(MyIndex).y - MAX_MAPY + screen_yb + screen_yh) * PIC_Y + Player(MyIndex).YOffset
                    NewYOffset = 0
                End If
            Else
                NewPlayerY = Player(MyIndex).y - screen_yh
                NewY = screen_yh * PIC_Y
                NewYOffset = Player(MyIndex).YOffset
            End If
            
            'Calcul des variables de scrolling restante
            NewPlayerPicX = NewPlayerX * PIC_X
            NewPlayerPicY = NewPlayerY * PIC_Y
            NewPlayerPOffsetX = NewPlayerPicX + NewXOffset
            NewPlayerPOffsetY = NewPlayerPicY + NewYOffset
            
            MaxDrawMapX = NewPlayerX + screen_xg + screen_xd + 1
            MinDrawMapX = NewPlayerX - 1
            MaxDrawMapY = NewPlayerY + screen_yh + screen_yb + 1
            MinDrawMapY = NewPlayerY - 1
            If MaxDrawMapX > MAX_MAPX Then MaxDrawMapX = MAX_MAPX
            If MaxDrawMapY > MAX_MAPY Then MaxDrawMapY = MAX_MAPY
            If MinDrawMapX < 0 Then MinDrawMapX = 0
            If MinDrawMapY < 0 Then MinDrawMapY = 0
            
            ' Blit out tiles layers ground/anim1/anim2
            For y = MinDrawMapY To MaxDrawMapY
                For x = MinDrawMapX To MaxDrawMapX
                    Call BltTile(x, y)
                Next x
            Next y
       
             For i = 1 To MAX_MAP_ITEMS
                 If MapItem(i).num > 0 Then Call BltItem(i)
             Next i
                             
             For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) And Player(i).PartyIndex = Player(MyIndex).PartyIndex Then
                    If Map(Player(MyIndex).Map).guildSoloView = 1 Then
                        If Player(MyIndex).Guild = Player(i).Guild Then
                            Call BltPlayerOmbre(i)
                            Call BltPlayerBar(i)
                        End If
                    Else
                        Call BltPlayerOmbre(i)
                            Call BltPlayerBar(i)
                    End If
                End If
            Next i
             If AccOpt.PlayBar And Player(MyIndex).PartyIndex > 0 Then
                 For i = 1 To MAX_PLAYERS
                     If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) And Player(i).PartyIndex = Player(MyIndex).PartyIndex Then
                        If Map(Player(MyIndex).Map).guildSoloView = 1 Then
                            If Player(MyIndex).Guild = Player(i).Guild Then
                                Call BltPlayerBar(i)
                            End If
                        Else
                            Call BltPlayerBar(i)
                        End If
                     End If
                 Next i
             ElseIf AccOpt.PlayBar Then
                    Call BltPlayerBar(MyIndex)
             End If
             
             ' Blit out the sprite change attribute
             For y = MinDrawMapY To MaxDrawMapY
                 For x = MinDrawMapX To MaxDrawMapX
                     If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_SPRITE_CHANGE Then
                         Call BltSpriteChange(x, y)
                         If PIC_PL > 1 Then Call BltSpriteChange2(x, y)
                     End If
                 Next x
             Next y
            
             ' Blit out the npcs
             For i = 1 To MAX_MAP_NPCS
                 If MapNpc(i).num > 0 And MapNpc(i).num < MAX_NPCS Then
                     If CLng(Npc(MapNpc(i).num).Vol) = 0 Then
                         Call BltNpc(i)
                         If AccOpt.NpcBar Then Call BltNpcBars(i)
                     End If
                 End If
             Next i
             
             ' Blit out players
             For i = 1 To MAX_PLAYERS
                 If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                     If Map(Player(MyIndex).Map).guildSoloView = 1 Then
                        If Player(MyIndex).Guild = Player(i).Guild Then
                            Call BltPlayer(i)
                        End If
                    Else
                        Call BltPlayer(i)
                    End If
                     Call BltArrow(i)
                     If Map(Player(MyIndex).Map).guildSoloView = 1 Then
                        If Player(MyIndex).Guild = Player(i).Guild Then
                            If Player(i).PetSlot <> 0 Then Call BltPlayerPet(i)
                        End If
                    Else
                        If Player(i).PetSlot <> 0 Then Call BltPlayerPet(i)
                    End If
                    
                 End If
             Next i

             ' Dessiner le haut des npc apres le bas des joueurs
             For i = 1 To MAX_MAP_NPCS
                 If MapNpc(i).num > 0 And MapNpc(i).num < MAX_NPCS Then If CLng(Npc(MapNpc(i).num).Vol) = 0 Then If PIC_PL > 1 Then Call BltNpcTop(i)
             Next i
             
             For i = 1 To MAX_PLAYERS
                 If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                     'Ajout du haut du personnage pour le 32*64
                     If PIC_PL > 1 Then
                        If Map(Player(MyIndex).Map).guildSoloView = 1 Then
                            If Player(MyIndex).Guild = Player(i).Guild Then
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
             '    If PIC_PL > 1 Then Call BltPlayerTop(MyIndex): Call BltEmoticons(MyIndex)
             '    Call BltPlayer(MyIndex)
             '    Call BltSpell(MyIndex)
             '    If Player(MyIndex).LevelUpT + 3000 > Tick Then Call BltPlayerLevelUp(MyIndex) Else Player(MyIndex).LevelUpT = 0
             'End If
        
            If Not GettingMap And AccOpt.PlayName Then
                'Verouiller le backbuffer pour pouvoir ecrire le nom des joueurs et de leur guildes
                TexthDC = DD_BackBuffer.GetDC
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        If Map(Player(MyIndex).Map).guildSoloView = 1 Then
                            If Player(MyIndex).Guild = Player(i).Guild Then
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
        End If
        
        If Not GettingMap Then
            'Dessiner les PNJs volant
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).num > 0 And MapNpc(i).num < MAX_NPCS Then
                    If CLng(Npc(MapNpc(i).num).Vol) <> 0 Then
                        Call BltNpc(i)
                        If AccOpt.NpcBar Then Call BltNpcBars(i)
                        If PIC_PL > 1 Then Call BltNpcTop(i)
                    End If
                End If
            Next i
        End If
        
        Call BltPlayerInt(MyIndex)
        
        If Not GettingMap Then If Map(GetPlayerMap(MyIndex)).Indoors = 0 Then Call BltWeather
    
        ' Lock the backbuffer so we can draw text and names
        TexthDC = DD_BackBuffer.GetDC
        If Not GettingMap Then
            If netbook = True Then
                cychat = 130
            Else
                cychat = 0
            End If
            If AccOpt.NpcDamage Then
                If NPCDmgDamage > 0 Then
                    If Not AccOpt.PlayName Then
                        If Tick < NPCDmgTime + 2000 Then Call DrawText(TexthDC, ((Len(NPCDmgDamage)) \ 2) * 3 + NewX + sx, NewY - 22 - cychat - ii + sx, NPCDmgDamage, QBColor(IIf(NPCDmgAddRem = 0, BrightRed, BrightGreen))) Else NPCDmgAddRem = 0
                    Else
                        If GetPlayerGuild(MyIndex) <> vbNullString Then
                            If Tick < NPCDmgTime + 2000 Then Call DrawText(TexthDC, ((Len(NPCDmgDamage)) \ 2) * 3 + NewX + sx, NewY - 42 - cychat - ii + sx, NPCDmgDamage, QBColor(IIf(NPCDmgAddRem = 0, BrightRed, BrightGreen))) Else NPCDmgAddRem = 0
                        Else
                            If Tick < NPCDmgTime + 2000 Then Call DrawText(TexthDC, ((Len(NPCDmgDamage)) \ 2) * 3 + NewX + sx, NewY - 22 - cychat - ii + sx, NPCDmgDamage, QBColor(IIf(NPCDmgAddRem = 0, BrightRed, BrightGreen))) Else NPCDmgAddRem = 0
                        End If
                    End If
                Else
                    If Not AccOpt.PlayName Then
                        If Tick < NPCDmgTime + 2000 Then Call DrawText(TexthDC, 6 + NewX + sx, NewY - 22 - cychat - ii + sx, "Raté", QBColor(BrightBlue)) Else NPCDmgAddRem = 0
                    Else
                        If GetPlayerGuild(MyIndex) <> vbNullString Then
                            If Tick < NPCDmgTime + 2000 Then Call DrawText(TexthDC, 6 + NewX + sx, NewY - 42 - cychat - ii + sx, "Raté", QBColor(BrightBlue)) Else NPCDmgAddRem = 0
                        Else
                            If Tick < NPCDmgTime + 2000 Then Call DrawText(TexthDC, 6 + NewX + sx, NewY - 22 - cychat - ii + sx, "Raté", QBColor(BrightBlue)) Else NPCDmgAddRem = 0
                        End If
                    End If
                End If
                ii = ii + 1
            End If
            
            If AccOpt.PlayDamage Then
                If NPCWho > 0 Then
                    If MapNpc(NPCWho).num > 0 Then
                        If DmgDamage > 0 Then
                            If Not AccOpt.NpcName Then
                                If Tick < DmgTime + 2000 Then Call DrawText(TexthDC, (MapNpc(NPCWho).x - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx - 20 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(IIf(DmgAddRem = 0, White, BrightGreen))) Else DmgAddRem = 0
                            Else
                                If Tick < DmgTime + 2000 Then Call DrawText(TexthDC, (MapNpc(NPCWho).x - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx - 30 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(IIf(DmgAddRem = 0, White, BrightGreen))) Else DmgAddRem = 0
                            End If
                        Else
                            If Not AccOpt.NpcName Then
                                If Tick < DmgTime + 2000 Then Call DrawText(TexthDC, (MapNpc(NPCWho).x - NewPlayerX) * PIC_X + sx + 6 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx - 20 + MapNpc(NPCWho).YOffset - NewYOffset - iii, "Raté", QBColor(BrightBlue)) Else DmgAddRem = 0
                            Else
                                If Tick < DmgTime + 2000 Then Call DrawText(TexthDC, (MapNpc(NPCWho).x - NewPlayerX) * PIC_X + sx + 6 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx - 30 + MapNpc(NPCWho).YOffset - NewYOffset - iii, "Raté", QBColor(BrightBlue)) Else DmgAddRem = 0
                            End If
                        End If
                        iii = iii + 1
                    End If
                End If
            End If
     
            ' speech bubble stuffs
            If AccOpt.SpeechBubbles Then
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        If Bubble(i).Text <> vbNullString Then Call BltPlayerText(i)
                        If Tick > Bubble(i).Created + DISPLAY_BUBBLE_TIME Then Bubble(i).Text = vbNullString
                    End If
                Next i
            End If
    
            'Draw NPC Names
            If AccOpt.NpcName Then
                For i = LBound(MapNpc) To UBound(MapNpc)
                    If MapNpc(i).num > 0 Then Call BltMapNPCName(i)
                Next i
            End If
                    
            ' Draw map name
            If Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_NONE Then
                ' Int((5) * PIC_X / 2) - (Len(Trim$(Map(GetPlayerMap(MyIndex)).name))) + sx
                Call DrawText(TexthDC, (frmMirage.picScreen.Width / 2) - (Len(Trim$(Map(GetPlayerMap(MyIndex)).name)) / 2), 5 + sx, Trim$(Map(GetPlayerMap(MyIndex)).name), QBColor(White))
            ElseIf Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_SAFE Then
                Call DrawText(TexthDC, (frmMirage.picScreen.Width / 2) - (Len(Trim$(Map(GetPlayerMap(MyIndex)).name)) / 2), 5 + sx, Trim$(Map(GetPlayerMap(MyIndex)).name), QBColor(White))
            ElseIf Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_NO_PENALTY Then
                Call DrawText(TexthDC, (frmMirage.picScreen.Width / 2) - (Len(Trim$(Map(GetPlayerMap(MyIndex)).name)) / 2), 5 + sx, Trim$(Map(GetPlayerMap(MyIndex)).name), QBColor(Black))
            End If
            
            For i = 1 To MAX_BLT_LINE
                If BattlePMsg(i).Index > 0 Then
                    If BattlePMsg(i).Color > 15 Then Coulor = BattlePMsg(i).Color Else Coulor = QBColor(BattlePMsg(i).Color)
                    If BattlePMsg(i).Time + 60000 > Tick Then Call DrawText(TexthDC, 1 + sx, BattlePMsg(i).y + PicScHeight - 80 - cychat + sx, Trim$(BattlePMsg(i).Msg), Coulor) Else BattlePMsg(i).Done = 0
                End If
                
                If BattleMMsg(i).Index > 0 Then
                    If BattleMMsg(i).Color > 15 Then Coulor = BattleMMsg(i).Color Else Coulor = QBColor(BattleMMsg(i).Color)
                    If BattleMMsg(i).Time + 60000 > Tick Then Call DrawText(TexthDC, (PicScWidth - (Len(BattleMMsg(i).Msg) * 8)) + sx, BattleMMsg(i).y + PicScHeight - 80 - cychat + sx, Trim$(BattleMMsg(i).Msg), Coulor) Else BattleMMsg(i).Done = 0
                End If
            Next i
        End If
        
        'Dessin de la nuit en "low effect"
        If GameTime = TIME_NIGHT And AccOpt.LowEffect And Map(GetPlayerMap(MyIndex)).Indoors = 0 Then Call Night(MinDrawMapX, MaxDrawMapX, MinDrawMapY, MaxDrawMapY)

        ' Check if we are getting a map, and if we are tell them so
        If GettingMap Then Call DrawText(TexthDC, 36, 70, "Chargement de la Carte en cours...", QBColor(BrightCyan))
                
        ' Release DC
        Call DD_BackBuffer.ReleaseDC(TexthDC)
        
        'Dessin du brouillard
        If Map(GetPlayerMap(MyIndex)).Fog <> 0 And Not AccOpt.LowEffect And GameTime <> TIME_NIGHT Then Call BltFog(MinDrawMapX, MaxDrawMapX, MinDrawMapY, MaxDrawMapY)
        
        
        
        'Dessin de la nuit en "hight"
        If GameTime = TIME_NIGHT And Not AccOpt.LowEffect And Map(GetPlayerMap(MyIndex)).Indoors = 0 Then Call Night(MinDrawMapX, MaxDrawMapX, MinDrawMapY, MaxDrawMapY)

        ' Get the rect to blit to
        Call dX.GetWindowRect(frmMirage.picScreen.hwnd, rec_pos)
        rec_pos.Bottom = rec_pos.Top - sx + ((MAX_MAPY + 1) * PIC_Y)
        rec_pos.Right = rec_pos.Left - sx + ((MAX_MAPX + 1) * PIC_X)
        rec_pos.Top = rec_pos.Bottom - ((MAX_MAPY + 1) * PIC_Y)
        rec_pos.Left = rec_pos.Right - ((MAX_MAPX + 1) * PIC_X)

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
            ' Thanks to kryzalid who told me about this "kind of lag"
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).num > 0 Then Call ProcessNpcMovement(i)
            Next i
            
            ' Change map animation every 250 milliseconds
            If Tick > MapAnimTimer + 250 Then
                If Not MapAnim Then MapAnim = True Else MapAnim = False
                MapAnimTimer = Tick
            End If
            
            Call MakeMidiLoop
            TickMove = Tick + 30
            
            'Calcul des FPS
            TickFPS = TickFPS + 1
            If TickFPS >= 33 Then TickFPS = 0: GameFPS = FPS: FPS = 0
        End If
        
        'Déchargement de textures en RAM
        UnloadTextures
        
        'Bloquer les FPS a 30 pour éviter de surcharger le processeur
        Do While GetTickCount < Tick + 30
            DoEvents
            Sleep 1
        Loop
  
        DoEvents
        'Sleep 2
        FPS = FPS + 1
    Loop
    If Not deco Then
        frmMirage.Visible = False
        frmsplash.Visible = True
        Call SetStatus("Destroying game data...")
        
        ' Shutdown the game
        Call GameDestroy
        
        ' Report disconnection if server disconnects
        If IsConnected = False Then Call MsgBox("Merci d'avoir joué à " & GAME_NAME & ".", vbOKOnly, GAME_NAME)
    Else
        deco = False
        Call MsgBox("Merci d'avoir joué à " & GAME_NAME & ".", vbOKOnly, GAME_NAME)
    End If
Exit Sub
er:
If Val(Mid(Err.Number, 1, 9)) = -200553208 Then GoTo rest:
Call MsgBox("Une erreur interne au logiciel c'est produite(Numéros de l'erreur : " & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "). Si le probléme pérsiste veulliez contacter un administrateur.", vbCritical, "Erreur")
Call GameDestroy
End Sub

Sub GameDestroy()
    If GettingMap = True Then Exit Sub
    On Error Resume Next
    DD.RestoreDisplayMode
    Call DestroyDirectX
    Call StopMidi
    Call WriteINI("UPDATER", "Fin", "0", App.Path & "\Config\Updater.ini")
    End
End Sub

Sub BltTile(ByVal x As Long, ByVal y As Long)
Dim Ground As Long
Dim Anim1 As Long
Dim Anim2 As Long
Dim Mask2 As Long
Dim M2Anim As Long
Dim Mask3 As Long
Dim M3Anim As Long
Dim GroundTileSet As Byte
Dim MaskTileSet As Byte
Dim AnimTileSet As Byte
Dim Mask2TileSet As Byte
Dim M2AnimTileSet As Byte
Dim Mask3TileSet As Byte
Dim M3AnimTileSet As Byte
Dim tx As Long
Dim ty As Long
    With Map(Player(MyIndex).Map).Tile(x, y)
        Ground = .Ground
        Anim1 = .Mask
        Anim2 = .Anim
        Mask2 = .Mask2
        M2Anim = .M2Anim
        Mask3 = .Mask3
        M3Anim = .M3Anim
        
        GroundTileSet = .GroundSet
        MaskTileSet = .MaskSet
        AnimTileSet = .AnimSet
        Mask2TileSet = .Mask2Set
        M2AnimTileSet = .M2AnimSet
        Mask3TileSet = .Mask3Set
        M3AnimTileSet = .M3AnimSet
    End With
    
    ' Only used if ever want to switch to blt rather then bltfast
    'With rec_pos
        '.Top = (y - NewPlayerY) * PIC_Y + sy - NewYOffset
        '.Bottom = .Top + PIC_Y
        '.Left = (x - NewPlayerX) * PIC_X + sy - NewXOffset
        '.Right = .Left + PIC_X
    'End With
    
    If GroundTileSet > ExtraSheets Then Exit Sub
    If Not TileFile(GroundTileSet) Then Exit Sub
    tx = (x - NewPlayerX) * PIC_X + sy - NewXOffset
    ty = (y - NewPlayerY) * PIC_Y + sy - NewYOffset
        
    rec.Top = (Ground \ TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Ground - (Ground \ TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
    'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(GroundTileSet), rec, DDBLT_WAIT)
    Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(GroundTileSet), rec, DDBLTFAST_WAIT)
   
    If (Not MapAnim) Or (Anim2 <= 0) Then
        ' Is there an animation tile to plot?
        If Anim1 > 0 And TempTile(x, y).DoorOpen = NO And MaskTileSet <= ExtraSheets Then
            If Not TileFile(MaskTileSet) Then Exit Sub
            rec.Top = (Anim1 \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Anim1 - (Anim1 \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(MaskTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
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
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
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
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(Mask2TileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
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
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(M2AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(M2AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If (Not MapAnim) Or (M3Anim <= 0) Then
        ' Is there an animation tile to plot?
        If Mask3 > 0 And Mask3TileSet <= ExtraSheets Then
            If Not TileFile(Mask3TileSet) Then Exit Sub
            rec.Top = (Mask3 \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Mask3 - (Mask3 \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(Mask3TileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
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
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(M3AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(M3AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    'Utiliser pour dessiner le panorama
    With rec_pos
        .Top = (y - NewPlayerY) * PIC_Y + sy - NewYOffset
        .Bottom = .Top + PIC_Y
        .Left = (x - NewPlayerX) * PIC_X + sx - NewXOffset
        .Right = .Left + PIC_X
    End With
    'Affichage du panorama inférieur si il y en à un
    If Trim$(Map(GetPlayerMap(MyIndex)).PanoInf) <> vbNullString Then
        rec.Top = y * PIC_Y
        If rec.Top + PIC_Y > DDSD_PanoInf.lHeight Then rec.Bottom = DDSD_PanoInf.lHeight: rec_pos.Bottom = rec_pos.Bottom - ((rec.Top + PIC_Y) - DDSD_PanoInf.lHeight) Else rec.Bottom = rec.Top + PIC_Y
        rec.Left = x * PIC_X
        If rec.Left + PIC_Y > DDSD_PanoInf.lWidth Then rec.Right = DDSD_PanoInf.lWidth: rec_pos.Right = rec_pos.Right - ((rec.Left + PIC_X) - DDSD_PanoInf.lWidth) Else rec.Right = rec.Left + PIC_X
        If Map(GetPlayerMap(MyIndex)).TranInf = 1 And TypeName(DD_PanoInfSurf) <> "Nothing" Then Call DD_BackBuffer.Blt(rec_pos, DD_PanoInfSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC) Else If TypeName(DD_PanoInfSurf) <> "Nothing" Then Call DD_BackBuffer.Blt(rec_pos, DD_PanoInfSurf, rec, DDBLT_WAIT)
    End If
End Sub

Sub BltItem(ByVal ItemNum As Long)
    ' Only used if ever want to switch to blt rather then bltfast
'    With rec_pos
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
    Call DD_BackBuffer.BltFast((MapItem(ItemNum).x - NewPlayerX) * PIC_X + sx - NewXOffset, (MapItem(ItemNum).y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
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
        .Top = -PIC_Y + (NewPlayerY * 32) + NewYOffset
        .Bottom = .Top + rec_pos.Bottom
        .Left = -PIC_X + (NewPlayerX * 32) + NewXOffset
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
Dim Fringe3 As Long
Dim F3Anim As Long
Dim FringeTileSet As Byte
Dim FAnimTileSet As Byte
Dim Fringe2TileSet As Byte
Dim F2AnimTileSet As Byte
Dim Fringe3TileSet As Byte
Dim F3AnimTileSet As Byte
Dim tx As Long
Dim ty As Long

    ' Only used if ever want to switch to blt rather then bltfast
'    With rec_pos
        '.Top = y * PIC_Y
        '.Bottom = .Top + PIC_Y
        '.Left = x * PIC_X
        '.Right = .Left + PIC_X
    'End With
    
    With Map(GetPlayerMap(MyIndex)).Tile(x, y)
        Fringe = .Fringe
        FAnim = .FAnim
        Fringe2 = .Fringe2
        F2Anim = .F2Anim
        Fringe3 = .Fringe3
        F3Anim = .F3Anim
        
        FringeTileSet = .FringeSet
        FAnimTileSet = .FAnimSet
        Fringe2TileSet = .Fringe2Set
        F2AnimTileSet = .F2AnimSet
        Fringe3TileSet = .Fringe3Set
        F3AnimTileSet = .F3AnimSet
    End With
    
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
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(Fringe2TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If F2Anim > 0 And F2AnimTileSet <= ExtraSheets Then
            If Not TileFile(F2AnimTileSet) Then Exit Sub
            rec.Top = (F2Anim \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (F2Anim - (F2Anim \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(F2AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If (Not MapAnim) Or (F3Anim <= 0) Then
        ' Is there an animation tile to plot?
        If Fringe3 > 0 And Fringe3TileSet <= ExtraSheets Then
            If Not TileFile(Fringe3TileSet) Then Exit Sub
            rec.Top = (Fringe3 \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Fringe3 - (Fringe3 \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(Fringe3TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If F3Anim > 0 And F3AnimTileSet <= ExtraSheets Then
            If Not TileFile(F3AnimTileSet) Then Exit Sub
            rec.Top = (F3Anim \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (F3Anim - (F3Anim \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(tx, ty, DD_TileSurf(F3AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    'Affichage du panorama supérieur si il y en à un
    If Trim$(Map(GetPlayerMap(MyIndex)).PanoSup) <> vbNullString Then
        rec.Top = y * PIC_Y
        If rec.Top + PIC_Y > DDSD_PanoSup.lHeight Then rec.Bottom = DDSD_PanoSup.lHeight: rec_pos.Bottom = rec_pos.Bottom - ((rec.Top + PIC_Y) - DDSD_PanoSup.lHeight) Else rec.Bottom = rec.Top + PIC_Y
        rec.Left = x * PIC_X
        If rec.Left + PIC_Y > DDSD_PanoSup.lWidth Then rec.Right = DDSD_PanoSup.lWidth: rec_pos.Right = rec_pos.Right - ((rec.Left + PIC_X) - DDSD_PanoSup.lWidth) Else rec.Right = rec.Left + PIC_X
        If Map(GetPlayerMap(MyIndex)).TranSup = 1 And TypeName(DD_PanoSupSurf) <> "Nothing" Then Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_PanoSupSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY) Else If TypeName(DD_PanoSupSurf) <> "Nothing" Then Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_PanoSupSurf, rec, DDBLTFAST_WAIT)
    End If
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
    
    num = Pets(Item(GetPlayerInvItemNum(Index, GetPlayerPetSlot(Index))).Data1).Sprite
    Call PreparePet(num)
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
    
    Call PrepareSprite(GetPlayerSprite(Index))
   
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
            ty = DDSD_PaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).paperdollPic).lHeight / 4
            tx = DDSD_PaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).paperdollPic).lWidth / 4
            
            rec.Top = GetPlayerDir(Index) * ty + (ty / 2)
            rec.Bottom = rec.Top + (ty / 2)
            rec.Left = Anim * tx + tx
            rec.Right = rec.Left + tx
        
            x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset - ((tx / 2) - 16)
            y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset
            
            Call PreparePaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).paperdollPic)
            If x < 0 Then rec.Left = rec.Left - x: rec.Right = rec.Left + (tx + x): x = 0
            If y < 0 Then rec.Top = rec.Top + (ty / 2): rec.Bottom = rec.Top: y = Player(Index).YOffset + sy
            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_PaperDollSurf(Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).paperdollPic), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If GetPlayerHelmetSlot(Index) > 0 Then
        If Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).paperdoll = 1 Then
            ty = DDSD_PaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).paperdollPic).lHeight / 4
            tx = DDSD_PaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).paperdollPic).lWidth / 4
            
            rec.Top = GetPlayerDir(Index) * ty + (ty / 2)
            rec.Bottom = rec.Top + (ty / 2)
            rec.Left = Anim * tx + tx
            rec.Right = rec.Left + tx
        
            x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset - ((tx / 2) - 16)
            y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset
            
            Call PreparePaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).paperdollPic)
            If x < 0 Then rec.Left = rec.Left - x: rec.Right = rec.Left + (tx + x): x = 0
            If y < 0 Then rec.Top = rec.Top + (ty / 2): rec.Bottom = rec.Top: y = Player(Index).YOffset + sy
            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_PaperDollSurf(Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).paperdollPic), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If GetPlayerWeaponSlot(Index) > 0 Then
        If Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).paperdoll = 1 Then
            ty = DDSD_PaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).paperdollPic).lHeight / 4
            tx = DDSD_PaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).paperdollPic).lWidth / 4
            
            rec.Top = GetPlayerDir(Index) * ty + (ty / 2)
            rec.Bottom = rec.Top + (ty / 2)
            rec.Left = Anim * tx + tx
            rec.Right = rec.Left + tx
        
            x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset - ((tx / 2) - 16)
            y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset
            
            Call PreparePaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).paperdollPic)
            If x < 0 Then rec.Left = rec.Left - x: rec.Right = rec.Left + (tx + x): x = 0
            If y < 0 Then rec.Top = rec.Top + (ty / 2): rec.Bottom = rec.Top: y = Player(Index).YOffset + sy
            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_PaperDollSurf(Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).paperdollPic), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If GetPlayerShieldSlot(Index) > 0 Then
        If Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).paperdoll = 1 Then
            ty = DDSD_PaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).paperdollPic).lHeight / 4
            tx = DDSD_PaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).paperdollPic).lWidth / 4
            
            rec.Top = GetPlayerDir(Index) * ty + (ty / 2)
            rec.Bottom = rec.Top + (ty / 2)
            rec.Left = Anim * tx + tx
            rec.Right = rec.Left + tx
        
            x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset - ((tx / 2) - 16)
            y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset
            
            Call PreparePaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).paperdollPic)
            If x < 0 Then rec.Left = rec.Left - x: rec.Right = rec.Left + (tx + x): x = 0
            If y < 0 Then rec.Top = rec.Top + (ty / 2): rec.Bottom = rec.Top: y = Player(Index).YOffset + sy
            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_PaperDollSurf(Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).paperdollPic), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    'FIN PAPERDOLL
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
   Call PrepareSprite(GetPlayerSprite(Index))
   
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
            ty = DDSD_PaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).paperdollPic).lHeight / 4
            tx = DDSD_PaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).paperdollPic).lWidth / 4
            
            rec.Top = GetPlayerDir(Index) * ty
            rec.Bottom = rec.Top + (ty / 2)
            rec.Left = Anim * tx + tx
            rec.Right = rec.Left + tx
            
            x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset - ((tx / 2) - 16)
            y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - (ty / 2)
            Call PreparePaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).paperdollPic)
            If x < 0 Then rec.Left = rec.Left - x: rec.Right = rec.Left + (tx + x): x = 0
            If y < 0 Then rec.Top = rec.Top + (ty / 2): rec.Bottom = rec.Top: y = Player(Index).YOffset + sy
            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_PaperDollSurf(Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).paperdollPic), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If GetPlayerHelmetSlot(Index) > 0 Then
        If Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).paperdoll = 1 Then
            ty = DDSD_PaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).paperdollPic).lHeight / 4
            tx = DDSD_PaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).paperdollPic).lWidth / 4
            
            rec.Top = GetPlayerDir(Index) * ty
            rec.Bottom = rec.Top + (ty / 2)
            rec.Left = Anim * tx + tx
            rec.Right = rec.Left + tx
            
            x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset - ((tx / 2) - 16)
            y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - (ty / 2)
            Call PreparePaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).paperdollPic)
            If x < 0 Then rec.Left = rec.Left - x: rec.Right = rec.Left + (tx + x): x = 0
            If y < 0 Then rec.Top = rec.Top + (ty / 2): rec.Bottom = rec.Top: y = Player(Index).YOffset + sy
            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_PaperDollSurf(Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).paperdollPic), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If GetPlayerWeaponSlot(Index) > 0 Then
        If Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).paperdoll = 1 Then
            ty = DDSD_PaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).paperdollPic).lHeight / 4
            tx = DDSD_PaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).paperdollPic).lWidth / 4
            
            rec.Top = GetPlayerDir(Index) * ty
            rec.Bottom = rec.Top + (ty / 2)
            rec.Left = Anim * tx + tx
            rec.Right = rec.Left + tx
            
            x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset - ((tx / 2) - 16)
            y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - (ty / 2)
            Call PreparePaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).paperdollPic)
            If x < 0 Then rec.Left = rec.Left - x: rec.Right = rec.Left + (tx + x): x = 0
            If y < 0 Then rec.Top = rec.Top + (ty / 2): rec.Bottom = rec.Top: y = Player(Index).YOffset + sy
            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_PaperDollSurf(Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).paperdollPic), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If GetPlayerShieldSlot(Index) > 0 Then
        If Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).paperdoll = 1 Then
            ty = DDSD_PaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).paperdollPic).lHeight / 4
            tx = DDSD_PaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).paperdollPic).lWidth / 4
            
            rec.Top = GetPlayerDir(Index) * ty
            rec.Bottom = rec.Top + (ty / 2)
            rec.Left = Anim * tx + tx
            rec.Right = rec.Left + tx
            
            x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset - ((tx / 2) - 16)
            y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - (ty / 2)
            Call PreparePaperDoll(Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).paperdollPic)
            If x < 0 Then rec.Left = rec.Left - x: rec.Right = rec.Left + (tx + x): x = 0
            If y < 0 Then rec.Top = rec.Top + (ty / 2): rec.Bottom = rec.Top: y = Player(Index).YOffset + sy
            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_PaperDollSurf(Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).paperdollPic), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    'FIN PAPERDOLL
End Sub

Sub BltMapNPCName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long

If Mid$(Trim$(Npc(MapNpc(Index).num).name), 1, 2) = "**" Then Exit Sub

With Npc(MapNpc(Index).num)
'Draw name
    TextX = MapNpc(Index).x * PIC_X + sx + MapNpc(Index).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.name)) / 2) * 8)
    If DDSD_Character(Npc(MapNpc(Index).num).Sprite).lHeight = 128 And DDSD_Character(Npc(MapNpc(Index).num).Sprite).lWidth = 128 Then
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
    
    Call PrepareSprite(Npc(MapNpc(MapNpcNum).num).Sprite)
    
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
    
    ty = DDSD_Character(Npc(MapNpc(MapNpcNum).num).Sprite).lHeight / 4
    tx = DDSD_Character(Npc(MapNpc(MapNpcNum).num).Sprite).lWidth / 4
    
    rec.Top = MapNpc(MapNpcNum).Dir * ty + (ty / 2)
    rec.Bottom = rec.Top + (ty / 2)
    rec.Left = Anim * tx + tx
    rec.Right = rec.Left + tx

    x = MapNpc(MapNpcNum).x * PIC_X + sx + MapNpc(MapNpcNum).XOffset - ((tx / 2) - 16) '(tx / 4) - ((tx / 4) / 2)
    y = MapNpc(MapNpcNum).y * PIC_Y + sx + MapNpc(MapNpcNum).YOffset
    
    If x < 0 Then rec.Left = rec.Left - x: rec.Right = rec.Left + (tx + x): x = 0
    If y < 0 Then rec.Top = rec.Top + (ty / 2): rec.Bottom = rec.Top: y = MapNpc(MapNpcNum).YOffset + sy
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_SpriteSurf(Npc(MapNpc(MapNpcNum).num).Sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltNpcTop(ByVal MapNpcNum As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim tx As Long, ty As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).num <= 0 Then Exit Sub
    Call PrepareSprite(Npc(MapNpc(MapNpcNum).num).Sprite)
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
    
    
    'rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * 64
    'rec.Bottom = rec.Top + PIC_Y
    'rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
    'rec.Right = rec.Left + PIC_X
    ty = DDSD_Character(Npc(MapNpc(MapNpcNum).num).Sprite).lHeight / 4
    tx = DDSD_Character(Npc(MapNpc(MapNpcNum).num).Sprite).lWidth / 4
    
    rec.Top = MapNpc(MapNpcNum).Dir * ty
    rec.Bottom = rec.Top + (ty / 2)
    rec.Left = Anim * tx + tx
    rec.Right = rec.Left + tx
    
    If tx > 32 Then
        x = MapNpc(MapNpcNum).x * PIC_X + sx + MapNpc(MapNpcNum).XOffset - ((tx / 2) - 16) '(tx / 4) - ((tx / 4) / 2)
    Else
        x = MapNpc(MapNpcNum).x * PIC_X + sx + MapNpc(MapNpcNum).XOffset
    End If
    y = MapNpc(MapNpcNum).y * PIC_Y + sx + MapNpc(MapNpcNum).YOffset - (ty / 2)
    
    If x < 0 Then rec.Left = rec.Left - x: rec.Right = rec.Left + (tx + x): x = 0
    If y < 0 Then rec.Top = rec.Top + (ty / 2): rec.Bottom = rec.Top: y = MapNpc(MapNpcNum).YOffset + sy
    
    Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_SpriteSurf(Npc(MapNpc(MapNpcNum).num).Sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
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
        Call DD_BackBuffer.BltFast(x - 32, y - 10 - Player(Index).LevelUp, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset
        y = GetPlayerY(Index) * PIC_Y + sy + Player(Index).YOffset
        Call DD_BackBuffer.BltFast(x - NewPlayerPicX - 32 - NewXOffset, y - NewPlayerPicY - 10 - Player(Index).LevelUp - NewYOffset, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
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
            Case 0: Color = QBColor(Brown)
            Case 1: Color = AccModo
            Case 2: Color = AccMapeur
            Case 3: Color = AccDevelopeur
            Case 4: Color = AccAdmin
        End Select
    Else
        Color = QBColor(BrightRed)
    End If
    
    ' Draw name
    TextX = Player(Index).x * PIC_X + sx + Player(Index).XOffset + (PIC_X \ 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
    If DDSD_Character(GetPlayerSprite(Index)).lHeight = 128 And DDSD_Character(GetPlayerSprite(Index)).lWidth = 128 Then
        TextY = Player(Index).y * PIC_Y + sx + Player(Index).YOffset - 40 - ((PIC_NPC1 - 1) * 10) + 16
    Else
        TextY = Player(Index).y * PIC_Y + sx + Player(Index).YOffset - 40 - ((PIC_NPC1 - 1) * 10)
    End If
    Call DrawText(TexthDC, TextX - NewPlayerPOffsetX, TextY - NewPlayerPOffsetY, GetPlayerName(Index), Color)
End Sub

Sub BltPlayerGuildName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long

    ' Check access level
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerGuildAccess(Index)
            Case 0: Color = QBColor(Red)
            Case 1: Color = QBColor(BrightCyan)
            Case 2: Color = QBColor(Yellow)
            Case 3: Color = QBColor(BrightGreen)
            Case 4: Color = QBColor(Yellow)
        End Select
    Else
        Color = QBColor(BrightRed)
    End If
    
    ' Draw name
    TextX = Player(Index).x * PIC_X + sx + Player(Index).XOffset + (PIC_X \ 2) - ((Len(GetPlayerGuild(Index)) / 2) * 8)
    TextY = Player(Index).y * PIC_Y + sx + Player(Index).YOffset - (PIC_Y \ 2) - 10 - ((PIC_NPC1 - 1) * 10)
    Call DrawText(TexthDC, TextX - NewPlayerPOffsetX, TextY - NewPlayerPOffsetY, GetPlayerGuild(Index), Color)
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
'MyText = frmMirage.txtMyTextBox.Text
If Len(frmMirage.txtMyTextBox.Text) > 200 Then
    MyText = Left(frmMirage.txtMyTextBox.Text, 200)
Else
    MyText = frmMirage.txtMyTextBox.Text
End If
' Handle when the player presses the return key
    
    If (KeyAscii = vbKeyReturn) Then
        If frmMirage.txtMyTextBox.Visible = True Then
            frmMirage.txtMyTextBox.Text = vbNullString
            frmMirage.txtMyTextBox.Locked = True
            frmMirage.txtMyTextBox.Visible = False
            frmMirage.Canal.Visible = False
            frmMirage.Canal.Locked = True
        Else
            frmMirage.txtMyTextBox.Locked = False
            frmMirage.txtMyTextBox.Text = vbNullString
            frmMirage.txtMyTextBox.Visible = True
            frmMirage.txtMyTextBox.SetFocus
            frmMirage.Canal.Visible = True
            frmMirage.Canal.Locked = False
            Exit Sub
        End If
    
        On Error Resume Next
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
        If Mid$(MyText, 1, 1) = "." Or Mid$(MyText, 1, 1) = "w" Then
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
                Call AddText("Utiliser: .nomjoueur msgici", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If
        
        If (Mid$(MyText, 1, 2) = "/w" Or Mid$(MyText, 1, 2) = "/W") And Mid$(MyText, 1, 7) <> "/warpto" Then
            ChatText = Mid$(MyText, 3, Len(MyText) - 2)
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
                Call AddText("Utiliser: /wnomjoueur msgici", AlertColor)
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
                
        'party request
        If LCase$(Mid$(MyText, 1, 7)) = "/invite" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 8 Then
                ChatText = Mid$(MyText, 9, Len(MyText) - 8)
                Call SendPartyRequest(ChatText)
            Else
                Call AddText("Utiliser : /invite nomdujoueur", AlertColor)
            End If
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
        If LCase$(Mid$(MyText, 1, 8)) = "/decline" Or LCase$(Mid$(MyText, 1, 5)) = "/refu" Then
            Call SendDeclineTrade
            MyText = vbNullString
            Exit Sub
        End If
                
        ' Party request
        If LCase$(Mid$(MyText, 1, 6)) = "/party" Or LCase$(Mid$(MyText, 1, 7)) = "/groupe" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 7 Then
                ChatText = Mid$(MyText, 8, Len(MyText) - 7)
                Call SendPartyRequest(ChatText)
            Else
                Call AddText("Utiliser : /group nomdujoueur", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Join party
        If LCase$(Mid$(MyText, 1, 5)) = "/join" Or LCase$(Mid$(MyText, 1, 7)) = "/rejoin" Then
            Call SendJoinParty
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Leave party
        If LCase$(Mid$(MyText, 1, 6)) = "/leave" Or LCase$(Mid$(MyText, 1, 7)) = "/quitte" Then
            Call SendLeaveParty
            MyText = vbNullString
            Exit Sub
        End If
        
        ' // Moniter Admin Commands //
        If GetPlayerAccess(MyIndex) > 0 Then
            ' day night command
            If LCase$(Mid$(MyText, 1, 9)) = "/daynight" Or LCase$(Mid$(MyText, 1, 9)) = "/journuit" Then
                If GameTime = TIME_DAY Then GameTime = TIME_NIGHT: Call InitNightAndFog(GetPlayerMap(MyIndex)) Else GameTime = TIME_DAY
                Call SendGameTime
                MyText = vbNullString
                Exit Sub
            End If
            
            ' weather command
            If LCase$(Mid$(MyText, 1, 8)) = "/weather" Or LCase$(Mid$(MyText, 1, 6)) = "/temps" Then
                If Len(MyText) > 8 Then
                    MyText = Mid$(MyText, 9, Len(MyText) - 8)
                    If IsNumeric(MyText) = True Then
                        Call SendData("weather" & SEP_CHAR & Val(MyText) & END_CHAR)
                    Else
                        If Trim$(LCase$(MyText)) = "none" Or Trim$(LCase$(MyText)) = "rien" Then i = 0
                        If Trim$(LCase$(MyText)) = "rain" Or Trim$(LCase$(MyText)) = "pluie" Then i = 1
                        If Trim$(LCase$(MyText)) = "snow" Or Trim$(LCase$(MyText)) = "neige" Then i = 2
                        If Trim$(LCase$(MyText)) = "thunder" Or Trim$(LCase$(MyText)) = "orage" Then i = 3
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
            
            ' Map report
            If LCase$(Mid$(MyText, 1, 10)) = "/mapreport" Then
                Call SendData("mapreport" & END_CHAR)
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
            
            ' Changement de nom
            If Mid$(MyText, 1, 5) = "/nom" Then
                If Len(MyText) > 6 Then
                    MyText = Mid$(MyText, 7, Len(MyText) - 6)
                    If Trim$(MyText) <> vbNullString Then
                        Call SendMOTDChange(MyText)
                    End If
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
                    If Trim$(MyText) <> vbNullString Then
                        Call SendMOTDChange(MyText)
                    End If
                End If
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Check the ban list
            If Mid$(MyText, 1, 8) = "/banlist" Then
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
        MyText = vbNullString
        KeyAscii = 0
    Exit Sub
    End If
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

    If GettingMap = False Then
        If KeyState = 1 Then
            If KeyCode = optTouche(CByte(Val(ReadINI("TJEU", "ramasser", App.Path & "\Config\Option.ini")))).Value Then
                If frmMirage.txtQ.Visible Then
                    frmMirage.txtQ.Visible = False
                Else
                    Call CheckMapGetItem
                End If
            End If
            If KeyCode = optTouche(CByte(Val(ReadINI("TJEU", "attaque", App.Path & "\Config\Option.ini")))).Value Then ControlDown = True
            If KeyCode = optTouche(CByte(Val(ReadINI("TJEU", "haut", App.Path & "\Config\Option.ini")))).Value Then
                DirUp = True
                DirDown = False
                DirLeft = False
                DirRight = False
            End If
            If KeyCode = optTouche(CByte(Val(ReadINI("TJEU", "bas", App.Path & "\Config\Option.ini")))).Value Then
                DirUp = False
                DirDown = True
                DirLeft = False
                DirRight = False
            End If
            If KeyCode = optTouche(CByte(Val(ReadINI("TJEU", "gauche", App.Path & "\Config\Option.ini")))).Value Then
                DirUp = False
                DirDown = False
                DirLeft = True
                DirRight = False
            End If
            If KeyCode = optTouche(CByte(Val(ReadINI("TJEU", "droite", App.Path & "\Config\Option.ini")))).Value Then
                DirUp = False
                DirDown = False
                DirLeft = False
                DirRight = True
            End If
            If KeyCode = optTouche(CByte(Val(ReadINI("TJEU", "courir", App.Path & "\Config\Option.ini")))).Value Then ShiftDown = True
        Else
            If KeyCode = optTouche(CByte(Val(ReadINI("TJEU", "haut", App.Path & "\Config\Option.ini")))).Value Then DirUp = False
            If KeyCode = optTouche(CByte(Val(ReadINI("TJEU", "bas", App.Path & "\Config\Option.ini")))).Value Then DirDown = False
            If KeyCode = optTouche(CByte(Val(ReadINI("TJEU", "gauche", App.Path & "\Config\Option.ini")))).Value Then DirLeft = False
            If KeyCode = optTouche(CByte(Val(ReadINI("TJEU", "droite", App.Path & "\Config\Option.ini")))).Value Then DirRight = False
            If KeyCode = optTouche(CByte(Val(ReadINI("TJEU", "courir", App.Path & "\Config\Option.ini")))).Value Then ShiftDown = False
            If KeyCode = optTouche(CByte(Val(ReadINI("TJEU", "attaque", App.Path & "\Config\Option.ini")))).Value Then ControlDown = False
        End If
    End If
End Sub

Public Sub InitMirageVars()
    PicScWidth = frmMirage.picScreen.Width
    PicScHeight = frmMirage.picScreen.height
End Sub

Function IsTryingToMove() As Boolean
    If (DirUp = True) Or (DirDown = True) Or (DirLeft = True) Or (DirRight = True) Then
        IsTryingToMove = True
    Else
        IsTryingToMove = False
    End If
End Function

Sub CaseChange(ByVal CX, ByVal CY)
Dim ONum As Long

If Val(ReadINI("CONFIG", "NomObjet", App.Path & "\Config\Account.ini")) = 0 Then frmMirage.ObjNm.Visible = False: Exit Sub

ONum = ObjetNumPos(CX, CY)

If ObjetPos(CX, CY) = True Then
    If Item(ONum).Type = ITEM_TYPE_CURRENCY Then
        frmMirage.OName.Caption = Trim$(Item(ONum).name) & "(" & ObjetValPos(CX, CY) & ")"
    Else
        frmMirage.OName.Caption = Trim$(Item(ONum).name) & "(1)"
    End If
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
            If Map(GetPlayerMap(MyIndex)).Up > 0 Then Call SendPlayerRequestNewMap: GettingMap = True
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
            If Map(GetPlayerMap(MyIndex)).Down > 0 Then Call SendPlayerRequestNewMap: GettingMap = True
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
            If Map(GetPlayerMap(MyIndex)).Left > 0 Then Call SendPlayerRequestNewMap: GettingMap = True
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
            If Map(GetPlayerMap(MyIndex)).Right > 0 Then Call SendPlayerRequestNewMap: GettingMap = True
            CanMove = False
            Exit Function
        End If
    End If
    If PX = 0 And PY = 0 Then CanMove = False: Exit Function
        ' Check to see if the map tile is blocked or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_SIGN Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_NIVEAUX Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_MONTURE Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_GUILDE Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_TOIT Then
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_MONTURE Then
                    If Player(MyIndex).ArmorSlot > 0 Then
                        If Item(Player(MyIndex).ArmorSlot).Type = ITEM_TYPE_MONTURE Then CanMove = False Else CanMove = True
                    Else
                        CanMove = True
                    End If
                ElseIf Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_NIVEAUX Then
                    If Player(MyIndex).level < Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Data1 Then CanMove = False Else CanMove = True
                ElseIf Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_GUILDE Then
                    If Trim$(Player(MyIndex).Guild) = Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).String1) Then CanMove = True Else CanMove = False
                Else
                    CanMove = False
                End If
                
                ' Set the new direction if they weren't facing that direction
                If d <> Dire Then Call Sendplayerdir
                Exit Function
            End If
            
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_CBLOCK Then
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Data1 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Data2 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Data3 = Player(MyIndex).Class Then Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> Dire Then Call Sendplayerdir
            End If
            
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_DIR Then
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Data1 = Player(MyIndex).Dir Then CanMove = True: Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Data2 = Player(MyIndex).Dir Then CanMove = True: Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Data3 = Player(MyIndex).Dir Then CanMove = True: Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> Dire Then Call Sendplayerdir
            End If
        ' verif atribut toit
        Call SuprTileToit(PY, PX)
                                                    
            ' Check to see if the key door is open or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_KEY Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_COFFRE Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_PORTE_CODE Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).DoorOpen = NO Then
                    CanMove = False
                    
                    ' Set the new direction if they weren't facing that direction
                    If d <> Dire Then
                        Call Sendplayerdir
                    End If
                Else
                    If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_COFFRE Then CanMove = False
                    Exit Function
                End If
            End If
                        
            ' Check to see if a player is already on that tile
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        If Map(Player(MyIndex).Map).guildSoloView = 1 Then
                            If Map(Player(MyIndex).Map).traversable = 0 Then
                                If Player(MyIndex).Guild = Player(i).Guild Then
                                    If (GetPlayerX(i) = GetPlayerX(MyIndex) + PX) And (GetPlayerY(i) = GetPlayerY(MyIndex) + PY) Then
                                        CanMove = False
                                    
                                        ' Set the new direction if they weren't facing that direction
                                        If d <> Dire Then Call Sendplayerdir
                                        Exit Function
                                    End If
                                End If
                            End If
                        Else
                            If Map(Player(MyIndex).Map).traversable = 0 Then
                                If (GetPlayerX(i) = GetPlayerX(MyIndex) + PX) And (GetPlayerY(i) = GetPlayerY(MyIndex) + PY) Then
                                    CanMove = False
                                
                                    ' Set the new direction if they weren't facing that direction
                                    If d <> Dire Then Call Sendplayerdir
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
                    If (MapNpc(i).x = GetPlayerX(MyIndex) + PX) And (MapNpc(i).y = GetPlayerY(MyIndex) + PY) And Npc(MapNpc(i).num).Vol = 0 Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> Dire Then Call Sendplayerdir
                        Exit Function
                    End If
                End If
            Next i
End Function

Sub SuprTileToit(ByVal dy As Long, ByVal dX As Long)
' verif atribut toit
On Error Resume Next
                
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, GetPlayerY(MyIndex) + dy).Type <> TILE_TYPE_WALKABLE And Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, GetPlayerY(MyIndex) + dy).Type <> TILE_TYPE_WALKABLE And Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_BLOCK_TOIT Then
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, GetPlayerY(MyIndex) + dy).Fringe > 0 Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, GetPlayerY(MyIndex) + dy).Fringe2 > 0 Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, GetPlayerY(MyIndex) + dy).F2Anim > 0 Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, GetPlayerY(MyIndex) + dy).FAnim > 0 Then
                Dim MX As Long
                Dim MY As Long
                Dim er As Long
                Dim i As Long
            
                
                If InToit = False Then
                
                For er = Player(MyIndex).y To MAX_MAPY
                If er < MAX_MAPY Then
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    For i = Player(MyIndex).x To MAX_MAPX
                    If i < MAX_MAPX Then
                        If Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).F3Anim = 0
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Type = TILE_TYPE_KEY Then
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).F3Anim = 0
                        End If
                    Else
                        If Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).F3Anim = 0
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Type = TILE_TYPE_KEY Then
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).F3Anim = 0
                        End If
                    End If
                    Next i
                        MX = Player(MyIndex).x
                    For i = 0 To Player(MyIndex).x
                        If Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).F3Anim = 0
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Type = TILE_TYPE_KEY Then
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Type = TILE_TYPE_BLOCKED Then Exit For
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).F3Anim = 0
                        End If
                        MX = MX - 1
                    Next i
                Else
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_KEY Then
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe2 = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe3 = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).FAnim = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).F2Anim = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).F3Anim = 0
                        Exit For
                    End If
                    If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe2 = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe3 = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).FAnim = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).F2Anim = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).F3Anim = 0
                End If
                Else
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).Type = TILE_TYPE_BLOCK_TOIT Then
                    For i = Player(MyIndex).x To MAX_MAPX
                    If i < MAX_MAPX Then
                        If Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).F3Anim = 0
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).Type = TILE_TYPE_KEY Then
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).F3Anim = 0
                            Exit For
                        End If
                            If Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er).F3Anim = 0
                        End If
                    Else
                        If Map(GetPlayerMap(MyIndex)).Tile(i, er).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(i, er).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(GetPlayerMap(MyIndex)).Tile(i, er).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er).F3Anim = 0
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(i, er).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(i, er).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(i, er).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(i, er).Type = TILE_TYPE_KEY Then
                            Map(GetPlayerMap(MyIndex)).Tile(i, er).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er).F3Anim = 0
                            Exit For
                        End If
                            If Map(GetPlayerMap(MyIndex)).Tile(i, er).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(GetPlayerMap(MyIndex)).Tile(i, er).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er).F3Anim = 0
                        End If
                    End If
                    Next i
                        MX = Player(MyIndex).x
                    For i = 0 To Player(MyIndex).x
                        If Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).F3Anim = 0
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).Type = TILE_TYPE_KEY Then
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).F3Anim = 0
                            Exit For
                        End If
                            If Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).Type = TILE_TYPE_BLOCKED Then Exit For
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er).F3Anim = 0
                        End If
                        MX = MX - 1
                    Next i
                Else
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).Type = TILE_TYPE_KEY Then
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).Fringe = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).Fringe2 = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).Fringe3 = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).FAnim = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).F2Anim = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).F3Anim = 0
                        Exit For
                    End If
                    If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).Fringe = 0
                    Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).Fringe2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).Fringe3 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).FAnim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).F2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er).F3Anim = 0
                End If
                End If
                Next er
                
                er = Player(MyIndex).y
                For MY = 0 To Player(MyIndex).y
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    For i = Player(MyIndex).x To MAX_MAPX
                    If i < MAX_MAPX Then
                        If Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).F3Anim = 0
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Type = TILE_TYPE_KEY Then
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i + dX, er + dy).F3Anim = 0
                        End If
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).F3Anim = 0
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Type = TILE_TYPE_KEY Then
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(i, er + dy).F3Anim = 0
                        End If
                    End If
                    Next i
                        MX = Player(MyIndex).x
                    For i = 0 To Player(MyIndex).x
                        If Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).F3Anim = 0
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Type = TILE_TYPE_KEY Then
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(MX + dX, er + dy).F3Anim = 0
                        End If
                        MX = MX - 1
                    Next i
                Else
                    If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_KEY Then
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe2 = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe3 = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).FAnim = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).F2Anim = 0
                        Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).F3Anim = 0
                        Exit For
                    End If
                    If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe = 0
                    Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe3 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).FAnim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).F2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + dX, er + dy).F3Anim = 0
                End If
                er = er - 1
                Next MY
                
                For er = Player(MyIndex).x To MAX_MAPX
                If er < MAX_MAPX Then
                If Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    For i = Player(MyIndex).y To MAX_MAPY
                    If i < MAX_MAPY Then
                        If Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).F3Anim = 0
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Type = TILE_TYPE_KEY Then
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).F3Anim = 0
                        End If
                    Else
                    If Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).F3Anim = 0
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Type = TILE_TYPE_KEY Then
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).F3Anim = 0
                            Exit For
                        End If
                            If Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).F3Anim = 0
                        End If
                    End If
                    Next i
                        MY = Player(MyIndex).y
                    For i = 0 To Player(MyIndex).y
                        If Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).F3Anim = 0
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Type = TILE_TYPE_KEY Then
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).F3Anim = 0
                        End If
                        MY = MY - 1
                    Next i
                Else
                    If Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_KEY Then
                        Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe = 0
                        Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe2 = 0
                        Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe3 = 0
                        Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).FAnim = 0
                        Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).F2Anim = 0
                        Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).F3Anim = 0
                        Exit For
                    End If
                    If Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe = 0
                    Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe3 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).FAnim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).F2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).F3Anim = 0
                End If
                Else
                If Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    For i = Player(MyIndex).y To MAX_MAPY
                    If i < MAX_MAPY Then
                        If Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).F3Anim = 0
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).Type = TILE_TYPE_KEY Then
                            Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i + dy).F3Anim = 0
                        End If
                    Else
                    If Map(GetPlayerMap(MyIndex)).Tile(er, i).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(er, i).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(GetPlayerMap(MyIndex)).Tile(er, i).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i).F3Anim = 0
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(er, i).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(er, i).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(er, i).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(er, i).Type = TILE_TYPE_KEY Then
                            Map(GetPlayerMap(MyIndex)).Tile(er, i).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i).F3Anim = 0
                            Exit For
                        End If
                            If Map(GetPlayerMap(MyIndex)).Tile(er, i).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(GetPlayerMap(MyIndex)).Tile(er, i).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, i).F3Anim = 0
                        End If
                    End If
                    Next i
                        MY = Player(MyIndex).y
                    For i = 0 To Player(MyIndex).y
                        If Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).F3Anim = 0
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).Type = TILE_TYPE_KEY Then
                            Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er, MY + dy).F3Anim = 0
                        End If
                        MY = MY - 1
                    Next i
                Else
                    If Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_KEY Then
                        Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).Fringe = 0
                        Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).Fringe2 = 0
                        Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).Fringe3 = 0
                        Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).FAnim = 0
                        Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).F2Anim = 0
                        Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).F3Anim = 0
                        Exit For
                    End If
                    If Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).Fringe = 0
                    Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).Fringe2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).Fringe3 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).FAnim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).F2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(er, GetPlayerY(MyIndex) + dy).F3Anim = 0
                End If
                End If
                Next er
                
                er = Player(MyIndex).x
                For MX = 0 To Player(MyIndex).x
                If Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    For i = Player(MyIndex).y To MAX_MAPY
                    If i < MAX_MAPY Then
                        If Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).F3Anim = 0
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Type = TILE_TYPE_KEY Then
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i + dy).F3Anim = 0
                        End If
                    Else
                        If Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).F3Anim = 0
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Type = TILE_TYPE_KEY Then
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).F3Anim = 0
                            Exit For
                        End If
                            If Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, i).F3Anim = 0
                        End If
                    End If
                    Next i
                        MY = Player(MyIndex).y
                    For i = 0 To Player(MyIndex).y
                        If Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Type = TILE_TYPE_TOIT Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).F3Anim = 0
                        Else
                        If Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Type = TILE_TYPE_KEY Then
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).F3Anim = 0
                            Exit For
                        End If
                            If Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Fringe = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Fringe2 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).Fringe3 = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).FAnim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).F2Anim = 0
                            Map(GetPlayerMap(MyIndex)).Tile(er + dX, MY + dy).F3Anim = 0
                        End If
                        MY = MY - 1
                    Next i
                Else
                    If Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_PORTE_CODE Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_WARP Or Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_KEY Then
                        Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe = 0
                        Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe2 = 0
                        Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe3 = 0
                        Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).FAnim = 0
                        Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).F2Anim = 0
                        Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).F3Anim = 0
                        Exit For
                    End If
                    If Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe = 0
                    Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe3 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).FAnim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).F2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(er + dX, GetPlayerY(MyIndex) + dy).F3Anim = 0
                End If
                er = er - 1
                Next MX
                InToit = True
                Else
                If InToit = True Then
                Call LoadMap(GetPlayerMap(MyIndex))
                End If
                InToit = False
                End If
            End If
            Else
                If InToit = True Then
                Call LoadMap(GetPlayerMap(MyIndex))
                End If
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
        
            Case DIR_DOWN
                Call SendPlayerMove
                Player(MyIndex).YOffset = PIC_Y * -1
                Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                                
            Case DIR_LEFT
                Call SendPlayerMove
                Player(MyIndex).XOffset = PIC_X
                Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                                
            Case DIR_RIGHT
                Call SendPlayerMove
                Player(MyIndex).XOffset = PIC_X * -1
                Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
        End Select
        If Player(MyIndex).Anim = 0 Then Player(MyIndex).Anim = 2 Else Player(MyIndex).Anim = 0
        
        ' Gotta check :)
        If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then GettingMap = True
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
        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(MyIndex, i) = ItemNum Then FindOpenInvSlot = i: Exit Function
        Next i
    End If
    
    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(MyIndex, i) <= 0 Then FindOpenInvSlot = i: Exit Function
    Next i
End Function

Public Sub UpdateTradeInventory()
Dim i As Long

    frmPlayerTrade.PlayerInv1.Clear
    
For i = 1 To MAX_INV
    If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
            frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerPetSlot(MyIndex) = i Then
                frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (worn)"
            Else
                frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name)
            End If
        End If
    Else
        frmPlayerTrade.PlayerInv1.AddItem "<Nothing>"
    End If
Next i
    
    frmPlayerTrade.PlayerInv1.ListIndex = 0
End Sub

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
    If MapItem(i).x = x And MapItem(i).y = y And MapItem(i).num > 0 Then ObjetValPos = MapItem(i).Value
Next i

End Function

Sub PlayerSearch(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1 As Long, y1 As Long

    x1 = (x \ PIC_X)
    y1 = (y \ PIC_Y)
    
    If (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
        Call SendData("search" & SEP_CHAR & x1 & SEP_CHAR & y1 & END_CHAR)
    End If
    MouseDownX = x1
    MouseDownY = y1
End Sub

Sub BltTile2(ByVal x As Long, ByVal y As Long, ByVal Tile As Long)
    rec.Top = (Tile \ TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Tile - (Tile \ TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
    Call DD_BackBuffer.BltFast(x - NewPlayerPicX + sx - NewXOffset, y - NewPlayerPicY + sx - NewYOffset, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
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
            
            If bytLineCount = MAX_LINES Then
                bytLineCount = bytLineCount - 1
                Exit For
            End If
            
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
            Call DrawText(TexthDC, TextX - NewPlayerPicX + sx - NewXOffset + (((DISPLAY_BUBBLE_WIDTH) * PIC_X) / 2) - ((Len(strLine(intLoop)) * 8) \ 2) - 7, TextY - NewPlayerPicY + sx - NewYOffset, strLine(intLoop), QBColor(DarkGrey))
            TextY = TextY + 16
        End If
    Next intLoop
End Sub
Sub BltPlayerBar(ByVal Index As Integer)
Dim x As Long, y As Long, ty As Long
    
    If Player(Index).HP <> 0 Then
        ty = (DDSD_Character(GetPlayerSprite(Index)).lHeight / 4) / 2
        x = (GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset) - NewPlayerPOffsetX
        y = (GetPlayerY(Index) * PIC_Y + sy + Player(Index).YOffset) - NewPlayerPOffsetY + ty + 3
        'draws the back bars
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(x, y + 2, x + 32, y - 2)
        ' Bar MP
        'If Player(Index).MaxMp > 0 Then
        '    Call DD_BackBuffer.SetFillColor(RGB(122, 10, 122))
        '    Call DD_BackBuffer.DrawBox(x, y + 2, x + 32, y + 6)
        'End If
    
        'draws HP
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        Call DD_BackBuffer.DrawBox(x, y + 2, x + (Player(Index).HP / Player(Index).MaxHp * 32), y - 2)
        ' Bar MP
        'If Player(Index).MaxMp > 0 Then
        '    Call DD_BackBuffer.SetFillColor(RGB(0, 0, 255))
        '    Call DD_BackBuffer.DrawBox(x, y + 2, x + (Player(Index).MP / Player(Index).MaxMp * 32), y + 6)
        'End If
    End If
End Sub
Sub BltNpcBars(ByVal Index As Long)
Dim x As Long, y As Long, ty As Long

If MapNpc(Index).HP = 0 Or MapNpc(Index).MaxHp <= 0 Or MapNpc(Index).num < 1 Then Exit Sub

    ty = (DDSD_Character(Npc(MapNpc(Index).num).Sprite).lHeight / 4) / 2
    x = (MapNpc(Index).x * PIC_X + sx + MapNpc(Index).XOffset) - NewPlayerPOffsetX
    y = (MapNpc(Index).y * PIC_Y + sy + MapNpc(Index).YOffset) - NewPlayerPOffsetY + ty + 3
    
    Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
    Call DD_BackBuffer.DrawBox(x, y, x + 32, y + 4)
    Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
    Call DD_BackBuffer.DrawBox(x, y, x + (MapNpc(Index).HP / MapNpc(Index).MaxHp * 32), y + 4)
    If MapNpc(Index).MaxMp > 0 Then
       Call DD_BackBuffer.SetFillColor(RGB(122, 10, 122))
       Call DD_BackBuffer.DrawBox(x, y + 4, x + 32, y + 4 + 4)
       Call DD_BackBuffer.SetFillColor(RGB(0, 0, 255))
       Call DD_BackBuffer.DrawBox(x, y + 4, x + (MapNpc(Index).MP / MapNpc(Index).MaxMp * 32), y + 4 + 4)
    End If
End Sub

Public Sub AffInv()
Dim Q As Long
Dim Qq As Long
    For Q = 0 To MAX_INV - 1
        Qq = Player(MyIndex).Inv(Q + 1).num
        If Qq = 0 Then frmMirage.picInv(Q).Picture = LoadPicture() Else Call AffSurfPic(DD_ItemSurf, frmMirage.picInv(Q), (Item(Qq).Pic - (Item(Qq).Pic \ 6) * 6) * PIC_X, (Item(Qq).Pic \ 6) * PIC_Y)
    Next Q
End Sub

Public Sub Affspell()
Dim Q As Long
Dim Qq As Long
    For Q = 0 To MAX_PLAYER_SPELLS - 1
        Qq = Player(MyIndex).Spell(Q + 1)
        If Qq = 0 Then
            frmMirage.picspell(Q).Picture = LoadPicture()
        Else
            Call AffSurfPic(DD_ItemSurf, frmMirage.picspell(Q), (Spell(Qq).SpellIco - (Spell(Qq).SpellIco \ 6) * 6) * PIC_X, (Spell(Qq).SpellIco \ 6) * PIC_Y)
        End If
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
        If GetPlayerWeaponSlot(MyIndex) = Index Then Call AffSurfPic(DD_ItemSurf, frmMirage.WeaponImage, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic \ 6) * 6) * PIC_X, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic \ 6) * PIC_Y)
        If GetPlayerHelmetSlot(MyIndex) = Index Then Call AffSurfPic(DD_ItemSurf, frmMirage.HelmetImage, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic \ 6) * 6) * PIC_X, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic \ 6) * PIC_Y)
        If GetPlayerArmorSlot(MyIndex) = Index Then Call AffSurfPic(DD_ItemSurf, frmMirage.ArmorImage, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic \ 6) * 6) * PIC_X, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic \ 6) * PIC_Y)
        If GetPlayerPetSlot(MyIndex) = Index Then Call AffSurfPic(DD_ItemSurf, frmMirage.PetImage, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic \ 6) * 6) * PIC_X, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic \ 6) * PIC_Y)
    Next Index
mont:
    frmMirage.EquipS(0).Visible = False
    frmMirage.EquipS(1).Visible = False
    frmMirage.EquipS(2).Visible = False
    frmMirage.EquipS(3).Visible = False
    frmMirage.EquipS(4).Visible = False
    
    For d = 0 To MAX_INV - 1
        If Player(MyIndex).Inv(d + 1).num > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type <> ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, d + 1)).Empilable = 0 Then
                'frmMirage.descName.Caption = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (" & GetPlayerInvItemValue(MyIndex, d + 1) & ")"
            'Else
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
                Else
                    'frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name)
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
PrepareSprite (GetPlayerSprite(MyIndex))
    ' Only used if ever want to switch to blt rather then bltfast
    'With rec_pos
        '.Top = y * PIC_Y
        '.Bottom = .Top + PIC_Y
        '.Left = x * PIC_X
        '.Right = .Left + PIC_X
    'End With
                                    
    rec.Top = Map(GetPlayerMap(MyIndex)).Tile(x, y).Data1 * (PIC_NPC1 * 32) + PIC_NPC2
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
    PrepareSprite (GetPlayerSprite(MyIndex))

    rec.Top = Map(GetPlayerMap(MyIndex)).Tile(x, y).Data1 * 64
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

    If index2 <= 0 Then Call clearItemSelected: Exit Sub

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

Sub clearItemSelected()
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

Sub AffSurfPic(DD_Surf As DirectDrawSurface7, ByVal PicBox As PictureBox, ByVal x As Long, ByVal y As Long)
Dim sRECT As RECT
Dim dRECT As RECT

    If Not (DD_Surf Is Nothing) Then
    If DD_Surf Is Nothing Then Exit Sub
    PicBox.Picture = LoadPicture()
    With dRECT
        .Top = 0
        .Bottom = PicBox.height
        .Left = 0
        .Right = PicBox.Width
    End With
    With sRECT
        .Top = y
        .Bottom = .Top + PicBox.height
        .Left = x
        .Right = .Left + PicBox.Width
    End With
    Call DD_Surf.BltToDC(PicBox.hDC, sRECT, dRECT)
    PicBox.Refresh
    End If
End Sub

Public Sub netbook_change()
Dim i As Byte
Dim Ending As String
    For i = 1 To 4
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        If i = 4 Then Ending = ".bmp"
        If netbook = True Then
            If FileExiste(Rep_Theme & "\Jeu\MiniInterface" & Ending) Then frmMirage.Interface.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\MiniInterface" & Ending)
        Else
            If FileExiste(Rep_Theme & "\Jeu\Interface" & Ending) Then frmMirage.Interface.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\Interface" & Ending)
        End If
    Next i
    For i = 0 To 13
        frmMirage.picRac(i).Visible = False
    Next i
    
    If netbook = True Then
        frmMirage.Interface.Width = 640
        frmMirage.picScreen.Width = 640
        frmMirage.picScreen.height = 416
        frmMirage.height = 7625
        frmMirage.Width = 9570
        For i = 0 To 8
            frmMirage.picRac(i).Visible = True
            frmMirage.picRac(i).Left = 7 + (i * 36)
            frmMirage.picRac(i).Top = 621 - 192
        Next i
        frmMirage.menu_inv.Left = 358
        frmMirage.menu_sort.Left = 382
        frmMirage.menu_equ.Left = 414
        frmMirage.menu_quete.Left = 446
        frmMirage.menu_guild.Left = 486
        frmMirage.menu_tchat.Left = 510
        frmMirage.menu_who.Left = 534
        frmMirage.menu_opt.Left = 574
        frmMirage.menu_quit.Left = 606
        
        frmMirage.menu_inv.Top = 616 - 188
        frmMirage.menu_sort.Top = 616 - 188
        frmMirage.menu_equ.Top = 616 - 188
        frmMirage.menu_quete.Top = 616 - 188
        frmMirage.menu_guild.Top = 616 - 188
        frmMirage.menu_tchat.Top = 616 - 188
        frmMirage.menu_who.Top = 616 - 188
        frmMirage.menu_opt.Top = 616 - 188
        frmMirage.menu_quit.Top = 616 - 188
    Else
        frmMirage.Interface.Width = 800
        frmMirage.picScreen.Width = 800
        frmMirage.picScreen.height = 608
        frmMirage.height = 10500
        frmMirage.Width = 12075
        For i = 0 To 13
            frmMirage.picRac(i).Visible = True
            frmMirage.picRac(i).Left = 7 + (i * 36)
            frmMirage.picRac(i).Top = 620
        Next i
        frmMirage.menu_inv.Left = 528
        frmMirage.menu_sort.Left = 552
        frmMirage.menu_equ.Left = 584
        frmMirage.menu_quete.Left = 616
        frmMirage.menu_guild.Left = 656
        frmMirage.menu_tchat.Left = 680
        frmMirage.menu_who.Left = 704
        frmMirage.menu_opt.Left = 744
        frmMirage.menu_quit.Left = 776
        
        frmMirage.menu_inv.Top = 616
        frmMirage.menu_sort.Top = 616
        frmMirage.menu_equ.Top = 616
        frmMirage.menu_quete.Top = 616
        frmMirage.menu_guild.Top = 616
        frmMirage.menu_tchat.Top = 616
        frmMirage.menu_who.Top = 616
        frmMirage.menu_opt.Top = 616
        frmMirage.menu_quit.Top = 616
    End If
    
    frmMirage.Interface.Top = frmMirage.picScreen.height
    frmMirage.txtQ.Top = frmMirage.picScreen.height - frmMirage.txtQ.height
    frmMirage.Canal.Top = frmMirage.picScreen.height - frmMirage.Canal.height
    frmMirage.txtMyTextBox.Top = frmMirage.picScreen.height - frmMirage.txtMyTextBox.height
    frmMirage.picParty.Top = frmMirage.picScreen.height - frmMirage.picParty.height
    frmMirage.fra_fenetre.Top = frmMirage.picScreen.height - frmMirage.fra_fenetre.height - 10
    frmMirage.fra_fenetre.Left = frmMirage.picScreen.Width - frmMirage.fra_fenetre.Width - 30
    If netbook = True Then
        frmMirage.itmDesc.Left = frmMirage.fra_fenetre.Left - frmMirage.itmDesc.Width
        frmMirage.itmDesc.Top = frmMirage.picScreen.height - frmMirage.itmDesc.height - 10
    Else
        frmMirage.itmDesc.Left = frmMirage.picScreen.Width - frmMirage.itmDesc.Width - 30
        frmMirage.itmDesc.Top = frmMirage.fra_fenetre.Top - frmMirage.itmDesc.height
    End If
End Sub
