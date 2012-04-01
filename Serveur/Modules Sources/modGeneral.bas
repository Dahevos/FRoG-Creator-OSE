Attribute VB_Name = "modGeneral"
Option Explicit

'***************************************************************************************************************************************************'
'ATTENTION : PENSER A NOTER LES MODIFICATIONS QUE VOUS APPORTER AU SOURCES POUR POUVOIR LES REFAIRE PLUS TARD SI VOUS DESIRER ACTUALISER LES SOURCES'
'***************************************************************************************************************************************************'
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Const CLIENT_MAJOR As String * 1 = "0"
Public Const CLIENT_MINOR As String * 1 = "6"
Public Const CLIENT_REVISION As String * 1 = "0"

'SCRIPTING
'Our dll cls
Global MyScript As clsSadScript
'Our hardcoded commands
Public clsScriptCommands As clsCommands
Public DetectScriptErr As Boolean

' Used for respawning items
Public SpawnSeconds As Long

' Used for weather effects
Public GameWeather As Long
Public WeatherSeconds As Long
Public GameTime As Long
Public TimeSeconds As Long
Public RainIntensity As Long
Public InDestroy As Boolean

' Used for closing key doors again
Public KeyTimer As Long

' Used for gradually giving back players and npcs hp
Public GiveHPTimer As Long
Public GiveNPCHPTimer As Long

' Used for logging
Public ServerLog As Boolean

'utiliser pour les cartes par FTP
Public CarteFTP As Boolean

Public Type LARGE_INTEGER
   LowPart As Long
   HighPart As Long
End Type

Sub InitServer()
Dim IPMask As String
Dim i As Long
Dim f As Long
    'On Error GoTo er:
    
    Randomize Timer
    
    CClasses = True
    InDestroy = False
    IBVisible = False
    
    Set HotelDeVente = New clsHdV
    Set Party = New clsParty
    
    CarteFTP = CBool(Val(GetVar(App.Path & "\Data.ini", "FTP", "ACTIF")))
    
    AdminMoMsg = False
    nid.cbSize = Len(nid)
    nid.hWnd = frmServer.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = frmServer.Icon
    nid.szTip = GAME_NAME & " Serveur" & vbNullChar
    
    ' Add to the sys tray
    Call Shell_NotifyIcon(NIM_ADD, nid)
    
    ' Init atmosphere
    GameWeather = WEATHER_NONE
    WeatherSeconds = 0
    GameTime = TIME_DAY
    TimeSeconds = 0
    RainIntensity = 25
    
    If LCase$(Dir(App.Path & "\maps", vbDirectory)) <> "maps" Then Call MkDir(App.Path & "\maps")
    If LCase$(Dir(App.Path & "\logs", vbDirectory)) <> "logs" Then Call MkDir(App.Path & "\Logs")
    If LCase$(Dir(App.Path & "\accounts", vbDirectory)) <> "accounts" Then Call MkDir(App.Path & "\accounts")
    If LCase$(Dir(App.Path & "\npcs", vbDirectory)) <> "npcs" Then Call MkDir(App.Path & "\Npcs")
    If LCase$(Dir(App.Path & "\items", vbDirectory)) <> "items" Then Call MkDir(App.Path & "\Items")
    If LCase$(Dir(App.Path & "\spells", vbDirectory)) <> "spells" Then Call MkDir(App.Path & "\Spells")
    If LCase$(Dir(App.Path & "\quetes", vbDirectory)) <> "quetes" Then Call MkDir(App.Path & "\Quetes")
    If LCase$(Dir(App.Path & "\shops", vbDirectory)) <> "shops" Then Call MkDir(App.Path & "\Shops")
    If LCase$(Dir(App.Path & "\classes", vbDirectory)) <> "classes" Then Call MkDir(App.Path & "\Classes")
    'If LCase$(Dir(App.Path & "\Pets", vbDirectory)) <> "Pets" Then Call MkDir(App.Path & "\Pets")
    If LCase$(Dir(App.Path & "\recettes", vbDirectory)) <> "recettes" Then Call MkDir(App.Path & "\recettes")
    'If LCase$(Dir(App.Path & "\Metiers", vbDirectory)) <> "Metiers" Then Call MkDir(App.Path & "\Metiers")
        
    SEP_CHAR = Chr$(0)
    END_CHAR = Chr$(237)
    
    ServerLog = True
    
    If Not FileExist("Data.ini") Then
        PutVar App.Path & "\Data.ini", "CONFIG", "GameName", ""
        PutVar App.Path & "\Data.ini", "CONFIG", "WebSite", ""
        PutVar App.Path & "\Data.ini", "CONFIG", "Port", 4000
        PutVar App.Path & "\Data.ini", "CONFIG", "HPRegen", 1
        PutVar App.Path & "\Data.ini", "CONFIG", "MPRegen", 1
        PutVar App.Path & "\Data.ini", "CONFIG", "SPRegen", 1
        PutVar App.Path & "\Data.ini", "CONFIG", "Scrolling", 1
        'PutVar App.Path & "\Data.ini", "CONFIG", "AutoTurn", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "Scripting", 1
        PutVar App.Path & "\Data.ini", "CONFIG", "ExpDynamique", 1
        PutVar App.Path & "\Data.ini", "MAX", "MAX_PLAYERS", 50
        PutVar App.Path & "\Data.ini", "MAX", "MAX_ITEMS", 100
        PutVar App.Path & "\Data.ini", "MAX", "MAX_NPCS", 100
        PutVar App.Path & "\Data.ini", "MAX", "MAX_SHOPS", 100
        PutVar App.Path & "\Data.ini", "MAX", "MAX_SPELLS", 100
        PutVar App.Path & "\Data.ini", "MAX", "MAX_MAPS", 255
        PutVar App.Path & "\Data.ini", "MAX", "MAX_MAP_ITEMS", 20
        PutVar App.Path & "\Data.ini", "MAX", "MAX_GUILDS", 20
        PutVar App.Path & "\Data.ini", "MAX", "MAX_GUILD_MEMBERS", 10
        PutVar App.Path & "\Data.ini", "MAX", "MAX_EMOTICONS", 10
        PutVar App.Path & "\Data.ini", "MAX", "MAX_LEVEL", 500
        PutVar App.Path & "\Data.ini", "MAX", "MAX_QUETES", 100
        PutVar App.Path & "\Data.ini", "MAX", "NOOB_LEVEL", 10
        PutVar App.Path & "\Data.ini", "MAX", "PK_LEVEL", 10
        PutVar App.Path & "\Data.ini", "MAX", "PIC_PL", 64
        PutVar App.Path & "\Data.ini", "MAX", "PIC_NPC1", 2
        PutVar App.Path & "\Data.ini", "MAX", "PIC_NPC2", 32
        PutVar App.Path & "\Data.ini", "MAX", "MAX_PETS", 10
        PutVar App.Path & "\Data.ini", "MAX", "MAX_METIER", 100
        PutVar App.Path & "\Data.ini", "MAX", "MAX_RECETTE", 200
        PutVar App.Path & "\Data.ini", "COULEURS", "AccAdmin", "16711935"
        PutVar App.Path & "\Data.ini", "COULEURS", "AccDevelopeur", "8388608"
        PutVar App.Path & "\Data.ini", "COULEURS", "AccModo", "8421504"
        PutVar App.Path & "\Data.ini", "COULEURS", "AccMapeur", "8421376"
        PutVar App.Path & "\Data.ini", "COULEURS", "MsgDiscu", "16777215"
        PutVar App.Path & "\Data.ini", "COULEURS", "MsgGlob", "32768"
        PutVar App.Path & "\Data.ini", "COULEURS", "MsgDist", "16777215"
        PutVar App.Path & "\Data.ini", "COULEURS", "MsgHurl", "16777215"
        PutVar App.Path & "\Data.ini", "COULEURS", "MsgEmot", "16777215"
        PutVar App.Path & "\Data.ini", "COULEURS", "MsgAdmin", "16776960"
        PutVar App.Path & "\Data.ini", "COULEURS", "MsgAide", "16777215"
        PutVar App.Path & "\Data.ini", "COULEURS", "MsgQui", "12632256"
        PutVar App.Path & "\Data.ini", "COULEURS", "MsgDep", "12632256"
        PutVar App.Path & "\Data.ini", "COULEURS", "MsgAlert", "16777215"
        PutVar App.Path & "\Data.ini", "COULEURS", "MsgGuilde", "65280"
        PutVar App.Path & "\Data.ini", "RATIO", "Exp_pvm", "1"
        PutVar App.Path & "\Data.ini", "RATIO", "Exp_pvp", "1"
    End If
    
    If Not FileExist("Stats.ini") Then
        PutVar App.Path & "\Stats.ini", "HP", "AddPerLevel", 10
        PutVar App.Path & "\Stats.ini", "HP", "AddPerStr", 10
        PutVar App.Path & "\Stats.ini", "HP", "AddPerDef", 0
        PutVar App.Path & "\Stats.ini", "HP", "AddPerMagi", 0
        PutVar App.Path & "\Stats.ini", "HP", "AddPerSpeed", 0
        PutVar App.Path & "\Stats.ini", "MP", "AddPerLevel", 10
        PutVar App.Path & "\Stats.ini", "MP", "AddPerStr", 0
        PutVar App.Path & "\Stats.ini", "MP", "AddPerDef", 0
        PutVar App.Path & "\Stats.ini", "MP", "AddPerMagi", 10
        PutVar App.Path & "\Stats.ini", "MP", "AddPerSpeed", 0
        PutVar App.Path & "\Stats.ini", "SP", "AddPerLevel", 10
        PutVar App.Path & "\Stats.ini", "SP", "AddPerStr", 0
        PutVar App.Path & "\Stats.ini", "SP", "AddPerDef", 0
        PutVar App.Path & "\Stats.ini", "SP", "AddPerMagi", 0
        PutVar App.Path & "\Stats.ini", "SP", "AddPerSpeed", 20
    End If

    Call SetStatus("Loading settings...")
    
    AddHP.Level = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerLevel"))
    AddHP.STR = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerStr"))
    AddHP.def = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerDef"))
    AddHP.magi = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerMagi"))
    AddHP.Speed = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerSpeed"))
    AddMP.Level = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerLevel"))
    AddMP.STR = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerStr"))
    AddMP.def = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerDef"))
    AddMP.magi = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerMagi"))
    AddMP.Speed = Val(GetVar(App.Path & "\Stats.ini", "MP", ""))
    AddSP.Level = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerLevel"))
    AddSP.STR = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerStr"))
    AddSP.def = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerDef"))
    AddSP.magi = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerMagi"))
    AddSP.Speed = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerSpeed"))
    
    GAME_NAME = Trim$(GetVar(App.Path & "\Data.ini", "CONFIG", "GameName"))
    MAX_PLAYERS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_PLAYERS"))
    MAX_ITEMS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_ITEMS"))
    MAX_NPCS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_NPCS"))
    MAX_SHOPS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_SHOPS"))
    MAX_SPELLS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_SPELLS"))
    MAX_MAPS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_MAPS"))
    MAX_MAP_ITEMS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_MAP_ITEMS"))
    MAX_GUILDS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_GUILDS"))
    MAX_GUILD_MEMBERS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_GUILD_MEMBERS"))
    MAX_EMOTICONS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_EMOTICONS"))
    MAX_LEVEL = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_LEVEL"))
    Scripting = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "Scripting"))
    MAX_QUETES = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_QUETES"))
    NOOB_LEVEL = Val(GetVar(App.Path & "\Data.ini", "MAX", "NOOB_LEVEL"))
    PK_LEVEL = Val(GetVar(App.Path & "\Data.ini", "MAX", "PK_LEVEL"))
    MAX_PETS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_PETS"))
    MAX_METIER = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_METIER"))
    MAX_RECETTE = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_RECETTE"))
    
    If MAX_PLAYERS <= 0 Then MAX_PLAYERS = 50
    If MAX_ITEMS <= 0 Then MAX_ITEMS = 100
    If MAX_NPCS <= 0 Then MAX_NPCS = 100
    If MAX_SHOPS <= 0 Then MAX_SHOPS = 100
    If MAX_SPELLS <= 0 Then MAX_SPELLS = 100
    If MAX_MAPS <= 0 Then MAX_MAPS = 255
    If MAX_MAP_ITEMS <= 0 Then MAX_MAP_ITEMS = 20
    If MAX_GUILDS <= 0 Then MAX_GUILDS = 20
    If MAX_GUILD_MEMBERS <= 0 Then MAX_GUILD_MEMBERS = 10
    If MAX_EMOTICONS <= 0 Then MAX_EMOTICONS = 100
    If MAX_LEVEL <= 0 Then MAX_LEVEL = 100
    If Scripting <= 0 Then Scripting = 1
    If MAX_QUETES <= 0 Then MAX_QUETES = 10
    If NOOB_LEVEL <= 0 Then NOOB_LEVEL = 10
    If MAX_PETS <= 0 Then MAX_PETS = 10
    If MAX_METIER <= 0 Then MAX_METIER = 100
    If MAX_RECETTE <= 0 Then MAX_RECETTE = 200
    If MAX_QUETES <= 0 Then MAX_QUETES = 100: Call PutVar(App.Path & "\Data.ini", "MAX", "MAX_QUETES", "100")

    MAX_MAPX = 30
    MAX_MAPY = 30
    
    If Val(GetVar(App.Path & "\Data.ini", "CONFIG", "Scrolling")) = 0 Then
        MAX_MAPX = 19
        MAX_MAPY = 14
    ElseIf Val(GetVar(App.Path & "\Data.ini", "CONFIG", "Scrolling")) = 1 Then
        MAX_MAPX = 30
        MAX_MAPY = 30
    End If
        
    ReDim quete(1 To MAX_QUETES) As QueteRec
    ReDim Map(1 To MAX_MAPS) As MapRec
    ReDim TempTile(1 To MAX_MAPS) As TempTileRec
    ReDim PlayersOnMap(1 To MAX_MAPS) As Long
    ReDim Player(1 To MAX_PLAYERS) As AccountRec
    ReDim item(0 To MAX_ITEMS) As ItemRec
    ReDim Npc(0 To MAX_NPCS) As NpcRec
    ReDim MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim MapNpc(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As MapNpcRec
    ReDim Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    ReDim Guild(1 To MAX_GUILDS) As GuildRec
    ReDim Emoticons(0 To MAX_EMOTICONS) As EmoRec
    ReDim Pets(1 To MAX_PETS) As PetsRec
    ReDim metier(1 To MAX_METIER) As MetierRec
    ReDim recette(1 To MAX_RECETTE) As RecetteRec
    
    For i = 1 To MAX_GUILDS
        ReDim Guild(i).Member(1 To MAX_GUILD_MEMBERS) As String * NAME_LENGTH
    Next i
    
    For i = 1 To MAX_PLAYERS
        For f = 1 To MAX_CHARS
            ReDim Player(i).Char(f).QueteStatut(1 To MAX_QUETES) As Integer
        Next f
    Next i
    
    For i = 1 To MAX_MAPS
        ReDim Map(i).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        ReDim TempTile(i).DoorOpen(0 To MAX_MAPX, 0 To MAX_MAPY) As Byte
    Next i
    
    ReDim experience(1 To MAX_LEVEL) As Long
    ReDim PnjMove(1 To MAX_MAP_NPCS, 1 To MAX_MAPS) As Boolean
    ReDim bouclier(1 To MAX_PLAYERS) As Boolean
    ReDim BouclierT(1 To MAX_PLAYERS) As Long
    ReDim Para(1 To MAX_PLAYERS) As Boolean
    ReDim ParaT(1 To MAX_PLAYERS) As Long
    ReDim Point(1 To MAX_PLAYERS) As Long
    ReDim PointT(1 To MAX_PLAYERS) As Long
    ReDim ParaN(1 To MAX_MAP_NPCS, 1 To MAX_MAPS) As Boolean
    ReDim ParaNT(1 To MAX_MAP_NPCS, 1 To MAX_MAPS) As Long
    
    START_MAP = 1
    START_X = MAX_MAPX / 2
    START_Y = MAX_MAPY / 2
        
    GAME_PORT = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "Port"))
    
    If Val(GetVar(App.Path & "\Data.ini", "MAX", "PIC_PL")) <= 0 Then
    PutVar App.Path & "\Data.ini", "MAX", "PIC_PL", 64
    PutVar App.Path & "\Data.ini", "MAX", "PIC_NPC1", 2
    PutVar App.Path & "\Data.ini", "MAX", "PIC_NPC2", 32
    End If
    
    'PIC_PL = Val(GetVar(App.Path & "\Data.ini", "MAX", "PIC_PL"))
    'PIC_NPC1 = Val(GetVar(App.Path & "\Data.ini", "MAX", "PIC_NPC1"))
    'PIC_NPC2 = Val(GetVar(App.Path & "\Data.ini", "MAX", "PIC_NPC2"))
    PIC_PL = 64
    PIC_NPC1 = 2
    PIC_NPC2 = 32
    
    'couleurs des accès
    If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "AccAdmin")) <> vbNullString Then AccAdmin = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "AccAdmin"))
    If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "AccDevelopeur")) <> vbNullString Then AccDevelopeur = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "AccDevelopeur"))
    If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "AccModo")) <> vbNullString Then AccModo = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "AccModo"))
    If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "AccMapeur")) <> vbNullString Then AccMapeur = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "AccMapeur"))
    
    'couleurs des messages
    If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgDiscu")) <> vbNullString Then SayColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgDiscu"))
    If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgGlob")) <> vbNullString Then GlobalColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgGlob"))
    If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgDist")) <> vbNullString Then TellColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgDist"))
    If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgHurl")) <> vbNullString Then BroadcastColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgHurl"))
    If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgEmot")) <> vbNullString Then EmoteColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgEmot"))
    If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgAdmin")) <> vbNullString Then AdminColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgAdmin"))
    If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgAide")) <> vbNullString Then HelpColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgAide"))
    If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgQui")) <> vbNullString Then WhoColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgQui"))
    If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgDep")) <> vbNullString Then JoinLeftColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgDep"))
    If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgAlert")) <> vbNullString Then AlertColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgAlert"))
    If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgGuilde")) <> vbNullString Then CouleurDesGuilde = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgGuilde"))
        
    If PIC_PL = 1 And PIC_NPC1 = 1 And PIC_NPC2 = 0 Then
    frmServer.petit.value = True
    frmServer.grand.value = False
    Else
    frmServer.grand.value = True
    frmServer.petit.value = False
    End If
    
    'Scripting
    If Scripting = 1 Then
        Call SetStatus("Chargement des scripts...")
        If FileExist("\Scripts\Main.txt") = False Then
            Call MsgBox("Le Main.txt est introuvable")
            Call DestroyServer
        End If
        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands
        MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
        frmServer.Bouclescript.Enabled = True
    End If
        
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = GAME_PORT
        
    ' Init all the player sockets
    For i = 1 To MAX_PLAYERS
        Call SetStatus("Initialisation des joueurs...")
        Call ClearPlayer(i)
        
        Load frmServer.Socket(i)
    Next i
    
    For i = 1 To MAX_PLAYERS
        Call ShowPLR(i)
    Next i
    
    If Not FileExist("CMessages.ini") Then
        For i = 1 To 6
            PutVar App.Path & "\CMessages.ini", "MESSAGES", "Title" & i, "Editer msg"
            PutVar App.Path & "\CMessages.ini", "MESSAGES", "Message" & i, ""
        Next i
    End If
    
    For i = 1 To 6
        CMessages(i).Title = GetVar(App.Path & "\CMessages.ini", "MESSAGES", "Title" & i)
        CMessages(i).Message = GetVar(App.Path & "\CMessages.ini", "MESSAGES", "Message" & i)
        frmServer.CustomMsg(i - 1).Caption = CMessages(i).Title
    Next i
    
    frmServer.lstTopics.Clear
    frmServer.lstTopics.AddItem "Introduction"
    frmServer.lstTopics.AddItem "Configurer le Serveur"
    frmServer.lstTopics.AddItem "Configurer le Client"
    frmServer.lstTopics.AddItem "Configurer l'Updater"
    frmServer.lstTopics.AddItem "Contrôle des joueurs"
    frmServer.lstTopics.AddItem "Commandes des joueurs"
    frmServer.lstTopics.AddItem "Discussions"
    frmServer.lstTopics.AddItem "Bugs/Erreurs"
    frmServer.lstTopics.AddItem "Convertisseur de cartes"
    frmServer.lstTopics.AddItem "Édition de cartes"
    frmServer.lstTopics.AddItem "Commande de scripts"
    frmServer.lstTopics.AddItem "Questions?"
    frmServer.lstTopics.AddItem "Nouveautés"
    frmServer.lstTopics.Selected(0) = True
    
    Call SetStatus("Nettoyage des tile temporaire...")
    Call ClearTempTile
    Call SetStatus("Nettoyage des cartes...")
    Call ClearMaps
    Call SetStatus("Nettoyage des objets des cartes...")
    Call ClearMapItems
    Call SetStatus("Nettoyage des PNJ des cartes...")
    Call ClearMapNpcs
    Call SetStatus("Nettoyage des PNJ...")
    Call ClearNpcs
    Call SetStatus("Nettoyage des objets...")
    Call ClearItems
    Call SetStatus("Nettoyage des magasins...")
    Call ClearShops
    Call SetStatus("Nettoyage des sorts...")
    Call ClearSpells
    Call SetStatus("Nettoyage de l'éxpérience...")
    Call ClearExps
    Call SetStatus("Nettoyage des émoticons...")
    Call ClearEmos
    Call SetStatus("Chargement des émoticons...")
    Call LoadEmos
    Call SetStatus("Nettoyage des flêches...")
    Call ClearArrows
    Call SetStatus("Chargement des flêches...")
    Call LoadArrows
    Call SetStatus("Chargement des émticons...")
    Call LoadExps
    Call SetStatus("Chargement des classes...")
    Call LoadClasses
    Call SetStatus("Nettoyage des familliers...")
    Call ClearPets
    Call SetStatus("Chargement des Familliers...")
    Call LoadPets
    Call SetStatus("Nettoyage des Metiers...")
    Call ClearMetiers
    Call SetStatus("Chargement des Metiers...")
    Call LoadMetiers
    Call SetStatus("Nettoyage des Recettes...")
    Call ClearRecettes
    Call SetStatus("Chargement des Recettes...")
    Call Loadrecettes
    'Call SetStatus("Loading first class advandcement...")
    'Call LoadClasses2
    'Call SetStatus("Loading second class advandcement...")
    'Call Loadclasses3
    Call SetStatus("Chargement des cartes...")
    Call LoadMaps
    Call SetStatus("Chargement des objets...")
    Call LoadItems
    Call SetStatus("Chargement des PNJ...")
    Call LoadNpcs
    Call SetStatus("Chargement des magasins...")
    Call LoadShops
    Call SetStatus("Chargement des sorts...")
    Call LoadSpells
    Call SetStatus("Chargement des quêtes...")
    Call LoadQuetes
    Call SetStatus("Placement des objets sur les cartes...")
    Call SpawnAllMapsItems
    Call SetStatus("Placement des PNJ sur les cartes...")
    Call SpawnAllMapNpcs
    Call VideIBMsg
    Call ChargIBOpt
    
    frmServer.MapList.Clear
        
    For i = 1 To MAX_MAPS
        frmServer.MapList.AddItem i & ": " & Map(i).Name
    Next i
    frmServer.MapList.Selected(0) = True
        
    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("accounts\charlist.txt") Then
        f = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Output As #f
        Close #f
    End If
    
    ' Start listening
    
    'On Error GoTo er:
    frmServer.Socket(0).Listen
    
    Call UpdateCaption
    
    frmLoad.Visible = False
    frmServer.Show
    
    SpawnSeconds = 0
    frmServer.tmrGameAI.Enabled = True
    
    Dim Repon As String
    If FileExist("\logs\admin.txt") Then
    If FileLen(App.Path & "\logs\admin.txt") > 5000000 Then
        Repon = MsgBox("Le fichier texte des logs d'administration a une taille supérieur à 5MO. Voulez-vous le suprimer?", vbYesNo, "Supression du Fichier")
        If Repon = vbYes Then Kill (App.Path & "\logs\admin.txt")
    End If
    End If
    
    If FileExist("\logs\player.txt") Then
    If FileLen(App.Path & "\logs\player.txt") > 5000000 Then
        Repon = MsgBox("Le fichier texte des logs des joueurs a une taille supérieur à 5MO. Voulez-vous le suprimer?", vbYesNo, "Supression du Fichier")
        If Repon = vbYes Then Kill (App.Path & "\logs\player.txt")
    End If
    End If
    
    If FileExist("\logs\InfoBulle.txt") Then
    If FileLen(App.Path & "\logs\InfoBulle.txt") > 5000000 Then
        Repon = MsgBox("Le fichier texte des logs de l'Info Bulle a une taille supérieur à 5MO. Voulez-vous le suprimer?", vbYesNo, "Supression du Fichier")
        If Repon = vbYes Then Kill (App.Path & "\logs\InfoBulle.txt")
    End If
    End If
    
    If FileExist("\logs\Err.txt") Then
    If FileLen(App.Path & "\logs\Err.txt") > 5000000 Then
        Repon = MsgBox("Le fichier texte des logs d'erreurs a une taille supérieur à 5MO. Voulez-vous le suprimer?", vbYesNo, "Supression du Fichier")
        If Repon = vbYes Then Kill (App.Path & "\logs\Err.txt")
    End If
    End If
    
    frmInfoBulle.Label2.Caption = "Serveur de " & GAME_NAME
    Call AffIBMsg(0, "Serveur chargé")

Exit Sub
er:
MsgBox "Erreur pendant l'initialisation du serveur, vérifiez que l'IP et le Port ne soit pas déjà utilisés(Détails :" & Err.Number & " " & Err.description & ")"
Call DestroyServer
End Sub

Sub DestroyServer()
Dim i As Long
    
    Close
    
    Call SetStatus("Fermeture en cours...")
    frmLoad.Visible = True
    frmServer.Visible = False
    DoEvents
    Call Shell_NotifyIcon(NIM_DELETE, nid)
    
    Call SetStatus("Sauvegarde des joueurs en ligne...")
    Call SaveAllPlayersOnline
    Call SetStatus("Nettoyage des cartes...")
    Call ClearMaps
    Call SetStatus("Nettoyage des objets sur les cartes...")
    Call ClearMapItems
    Call SetStatus("Nettoyage des PNJ sur les cartes...")
    Call ClearMapNpcs
    Call SetStatus("Nettoyage des NPCs...")
    Call ClearNpcs
    Call SetStatus("Nettoyage des Objets...")
    Call ClearItems
    Call SetStatus("Nettoyage des magasins...")
    Call ClearShops
    Call SetStatus("Fermeture du protocole TCP...")
    frmServer.tmrGameAI.Enabled = False
    
    On Error GoTo sock:
    For i = 1 To MAX_PLAYERS
        Call SetStatus("Fermeture du protocole TCP " & i & "/" & MAX_PLAYERS)
        DoEvents
        Unload frmServer.Socket(i)
    Next i
sock:
    
    If frmServer.chkChat.value = Checked Then
        Call SetStatus("Sauvegarde des logs de tchat...")
        Call SaveLogs
    End If
    InDestroy = True
    
    Set HotelDeVente = Nothing
    Set Party = Nothing

    Call Unload(frmEditor)
    Call Unload(frmLoad)
    Call Unload(frmServer)
    Call Unload(frmInfoBulle)
    Call Unload(frmOptCoul)
    Call Unload(frmOptInfoBulle)
    Call Unload(frmOptFTP)
    Call Unload(frmEnvFTP)
    Call Unload(frmclasseseditor)
    Call Unload(frmCoFTP)
    Call Shell_NotifyIcon(NIM_DELETE, nid)
    Call Unload(frmLoad)
End Sub

Sub SetStatus(ByVal Status As String)
    frmLoad.lblStatus.Caption = Status
    DoEvents
End Sub

Sub ServerLogic()
Dim i As Long

    ' Check for disconnections
    For i = 1 To MAX_PLAYERS
        If frmServer.Socket(i).State > 7 Then Call CloseSocket(i)
    Next i
        
    Call CheckIBMsg
    Call CheckGiveHP
    Call VerifEffetsJoueur
    Call GameAI
End Sub

Sub CheckSpawnMapItems()
Dim X As Long, Y As Long

    ' Used for map item respawning
    SpawnSeconds = SpawnSeconds + 1
    
    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    If SpawnSeconds >= 120 Then
        ' 2 minutes have passed
        For Y = 1 To MAX_MAPS
            ' Make sure no one is on the map when it respawns
            If PlayersOnMap(Y) = False Then
                ' Clear out unnecessary junk
                For X = 1 To MAX_MAP_ITEMS
                    Call ClearMapItem(X, Y)
                Next X
                    
                ' Spawn the items
                Call SpawnMapItems(Y)
                Call SendMapItemsToAll(Y)
            End If
            DoEvents
        Next Y
        
        SpawnSeconds = 0
    End If
End Sub

Sub GameAI()
Dim i As Long, X As Long, Y As Long, n As Long, x1 As Long, y1 As Long, TickCount As Long
Dim Damage As Long, DistanceX As Long, DistanceY As Long, npcnum As Long, Target As Long
Dim DidWalk As Boolean
Dim SpellSlot As Byte


            
    'WeatherSeconds = WeatherSeconds + 1
    'TimeSeconds = TimeSeconds + 1
    
    ' Lets change the weather if its time to
    If WeatherSeconds >= 60 Then
        i = Int(Rnd * 3)
        If i <> GameWeather Then
            GameWeather = i
            Call SendWeatherToAll
        End If
        WeatherSeconds = 0
    End If
    
    ' Check if we need to switch from day to night or night to day
    If TimeSeconds >= 60 Then
        If GameTime = TIME_DAY Then GameTime = TIME_NIGHT Else GameTime = TIME_DAY
        
        Call SendTimeToAll
        TimeSeconds = 0
    End If
            
    For Y = 1 To MAX_MAPS
        If PlayersOnMap(Y) = YES Then
            TickCount = GetTickCount
            
            ' ////////////////////////////////////
            ' // This is used for closing doors //
            ' ////////////////////////////////////
            If TickCount > TempTile(Y).DoorTimer + 5000 Then
                For y1 = 0 To MAX_MAPY
                    For x1 = 0 To MAX_MAPX
                        If Map(Y).Tile(x1, y1).type = TILE_TYPE_KEY And TempTile(Y).DoorOpen(x1, y1) = YES Then
                            TempTile(Y).DoorOpen(x1, y1) = NO
                            Call SendDataToMap(Y, "MAPKEY" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                        End If
                        
                        If Map(Y).Tile(x1, y1).type = TILE_TYPE_DOOR Or Map(Y).Tile(x1, y1).type = TILE_TYPE_COFFRE Or Map(Y).Tile(x1, y1).type = TILE_TYPE_PORTE_CODE And TempTile(Y).DoorOpen(x1, y1) = YES Then
                            TempTile(Y).DoorOpen(x1, y1) = NO
                            Call SendDataToMap(Y, "MAPKEY" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                        End If
                    Next x1
                Next y1
            End If
            
            For X = 1 To MAX_MAP_NPCS
                npcnum = MapNpc(Y, X).num
                
                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If MapNpc(Y, X).num > 0 And PnjMove(X, Y) = True Then
                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(npcnum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or Npc(npcnum).Behavior = NPC_BEHAVIOR_GUARD Then
                    
                        For i = 1 To MAX_PLAYERS
                            If IsPlaying(i) Then
                                
                                If GetPlayerMap(i) = Y And MapNpc(Y, X).Target = 0 And GetPlayerAccess(i) <= ADMIN_MONITER Then
                                    n = Npc(npcnum).Range
                                    
                                    DistanceX = MapNpc(Y, X).X - GetPlayerX(i)
                                    DistanceY = MapNpc(Y, X).Y - GetPlayerY(i)
                                    
                                    ' Make sure we get a positive value
                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                    If DistanceY < 0 Then DistanceY = DistanceY * -1
                                    
                                    ' Are they in range?  if so GET'M!
                                    If DistanceX <= n And DistanceY <= n Then
                                    
                                        If Npc(npcnum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                    
                                            If Trim$(Npc(npcnum).AttackSay) <> vbNullString Then Call QueteMsg(i, Trim$(Npc(npcnum).Name) & " : " & Trim$(Npc(npcnum).AttackSay) & "")
                                            
                                            MapNpc(Y, X).Target = i
                                        End If
                                    End If
                                End If
                                If GetPlayerAccess(i) >= ADMIN_MONITER And AdminMoMsg = False Then
                                    Call QueteMsg(i, "Les monstres ne vous attaquent pas car vous êtes un administrateur")
                                    AdminMoMsg = True
                                End If
                            End If
                        Next i
                        For i = 1 To MAX_MAP_NPCS
                            If MapNpc(Y, i).num > 0 And i <> X Then
                                If Npc(MapNpc(Y, i).num).Behavior = IIf(Npc(MapNpc(Y, X).num).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT, NPC_BEHAVIOR_GUARD, NPC_BEHAVIOR_ATTACKONSIGHT) Then
                                    DistanceX = Abs(MapNpc(Y, X).X - MapNpc(Y, i).X)
                                    DistanceY = Abs(MapNpc(Y, X).Y - MapNpc(Y, i).Y)
                                    
                                    If DistanceX <= Npc(MapNpc(Y, X).num).Range And DistanceY <= Npc(MapNpc(Y, X).num).Range Then
                                        MapNpc(Y, X).Target = i
                                        MapNpc(Y, X).TargetType = TARGET_TYPE_NPC
                                    End If
                                End If
                            End If
                        Next i
                    End If
                End If
                                                                        
                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If MapNpc(Y, X).num > 0 And PnjMove(X, Y) = True Then
                    Target = MapNpc(Y, X).Target
                    
                    ' Check to see if we are following a player or not
                    If Target > 0 Then
                        ' Check if the player is even playing, if so follow'm
                        If ValidTarget(Target, Y, MapNpc(Y, X).TargetType) Then
                            DidWalk = False
                            
                            i = Int(Rnd * 5)
                            
                            ' Lets move the npc
                            SelectMoveNpc i, Y, X, Target, MapNpc(Y, X).TargetType, DidWalk
                        Else
                            MapNpc(Y, X).Target = 0
                        End If
                    Else
                        If Map(Y).Npcs(X).Hasardm = 0 And Map(Y).Npcs(X).Imobile = 0 Then
                            If MapNpc(Y, X).X = Map(Y).Npcs(X).X And MapNpc(Y, X).Y = Map(Y).Npcs(X).Y Then
                                Map(Y).Npcs(X).Axy1 = False
                                Map(Y).Npcs(X).Axy = True
                                Map(Y).Npcs(X).Axy2 = False
                            ElseIf MapNpc(Y, X).X = Map(Y).Npcs(X).x1 And MapNpc(Y, X).Y = Map(Y).Npcs(X).y1 And Map(Y).Npcs(X).Axy = True Then
                                Map(Y).Npcs(X).Axy1 = True
                                Map(Y).Npcs(X).Axy = False
                                Map(Y).Npcs(X).Axy2 = False
                            ElseIf MapNpc(Y, X).X = Map(Y).Npcs(X).x2 And MapNpc(Y, X).Y = Map(Y).Npcs(X).y2 And Map(Y).Npcs(X).Axy1 = True Then
                                Map(Y).Npcs(X).Axy1 = False
                                Map(Y).Npcs(X).Axy = False
                                Map(Y).Npcs(X).Axy2 = True
                            End If
                        
                            ' mouvement a x1 et y1
                            If Map(Y).Npcs(X).Axy1 = True Then
                                If Map(Y).Npcs(X).x2 > 0 Or Map(Y).Npcs(X).y2 > 0 Then
                                    Call NpcMoveTo(Y, X, 2, MOVING_WALKING, Val(Map(Y).Npcs(X).x2), Val(Map(Y).Npcs(X).y2))
                                Else
                                    Call NpcMoveTo(Y, X, 2, MOVING_WALKING, Val(Map(Y).Npcs(X).X), Val(Map(Y).Npcs(X).Y))
                                End If
                            End If
                            
                            ' mouvement a X et Y
                            If Map(Y).Npcs(X).Axy = True Then Call NpcMoveTo(Y, X, 2, MOVING_WALKING, Val(Map(Y).Npcs(X).x1), Val(Map(Y).Npcs(X).y1))
                            
                            ' mouvement a x2 et y2
                            If Map(Y).Npcs(X).Axy2 = True Then Call NpcMoveTo(Y, X, 2, MOVING_WALKING, Val(Map(Y).Npcs(X).X), Val(Map(Y).Npcs(X).Y))
                            
                        Else
                            If Map(Y).Npcs(X).Imobile = 0 Then
                                i = Int(Rnd * 4)
                                If i = 1 Then
                                    i = Int(Rnd * 4)
                                    If CanNpcMove(Y, X, i) Then Call NpcMove(Y, X, i, MOVING_WALKING)
                                End If
                            End If
                        End If
                    End If
                End If
                
                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack players //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If MapNpc(Y, X).num > 0 And PnjMove(X, Y) = True Then
                    Target = MapNpc(Y, X).Target
                    
                    ' Check if the npc can attack the targeted player player
                    If Target > 0 Then
                        If MapNpc(Y, X).TargetType = TARGET_TYPE_PLAYER Then
                            ' Is the target playing and on the same map?
                            If IsPlaying(Target) And GetPlayerMap(Target) = Y Then
                                ' Verifie si le PNJ peut attaquer le joueur
                                If CanNpcAttackPlayer(X, Target) Then
                                    If Not CanPlayerBlockHit(Target) And Not CanPlayerEsquiveHit(Target) Then
                                        Damage = Npc(npcnum).STR - GetPlayerProtection(Target)
                                        If Damage > 0 Then
                                            Call NpcAttackPlayer(X, Target, Damage)
                                        Else
                                            Call BattleMsg(Target, Trim$(Npc(npcnum).Name) & " n'a pas pu vous blesser!", BrightBlue, 1)
                                        End If
                                    Else
                                        Call BattleMsg(Target, "Tu bloques/esquives le coup de " & Trim$(Npc(npcnum).Name), BrightCyan, 1)
                                    End If
                                ElseIf CanNpcAttackPlayerWithSpell(X, Target, SpellSlot) Then
                                    Call CastSpellTo(Target, Npc(MapNpc(Y, X).num).Spell(SpellSlot), X)
                                End If
                            Else
                                ' Player left map or game, set target to 0
                                MapNpc(Y, X).Target = 0
                            End If
                        ElseIf MapNpc(Y, X).TargetType = TARGET_TYPE_NPC Then
                            If MapNpc(Y, X).num > 0 Then
                                ' Can the npc attack the npc?
                                If CanNPCAttackNPC(Y, X, Target) Then
                                    Damage = Npc(npcnum).STR - Npc(MapNpc(Y, Target).num).def
                                    If Damage > 0 Then Call NPCAttackNPC(Y, Damage, X, Target)
                                End If
                            Else
                                ' Npc isn't on map, set target to 0
                                MapNpc(Y, X).Target = 0
                            End If
                        End If
                    End If
                End If
                
                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If MapNpc(Y, X).num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                    If MapNpc(Y, X).HP > 0 Then
                        MapNpc(Y, X).HP = MapNpc(Y, X).HP + GetNpcHPRegen(npcnum)
                        MapNpc(Y, X).MP = MapNpc(Y, X).MP + GetNpcMPRegen(npcnum) + IIf(MapNpc(Y, X).Amelio.Timer >= GetTickCount, MapNpc(Y, X).Amelio.Power / 3, 0)

                        ' Check if they have more then they should and if so just set it to max
                        If MapNpc(Y, X).HP > GetNpcMaxHP(npcnum) Then MapNpc(Y, X).HP = GetNpcMaxHP(npcnum)
                        If MapNpc(Y, X).MP > GetNpcMaxMP(npcnum) + IIf(MapNpc(Y, X).Amelio.Timer >= GetTickCount, MapNpc(Y, X).Amelio.Power * 2, 0) Then MapNpc(Y, X).MP = GetNpcMaxMP(npcnum) + IIf(MapNpc(Y, X).Amelio.Timer >= GetTickCount, MapNpc(Y, X).Amelio.Power * 2, 0)
                    End If
                End If
                
                ' ////////////////////////////////////////////////////
                ' // This is used for self-regenerating NPC's HP/MP //
                ' ////////////////////////////////////////////////////

                SpellSlot = 0
                If CanNpcRestoreHimself(X, Y, SpellSlot) Then
                    Call CastSpellOn(X, TARGET_TYPE_NPC, X, TARGET_TYPE_NPC, Y, SpellSlot)
                End If
                    
                ' ////////////////////////////////////////////////////////
                ' // This is used for checking if an NPC is dead or not //
                ' ////////////////////////////////////////////////////////
                ' Check if the npc is dead or not
                'If MapNpc(y, x).Num > 0 Then
                '    If MapNpc(y, x).HP <= 0 And Npc(MapNpc(y, x).Num).STR > 0 And Npc(MapNpc(y, x).Num).DEF > 0 Then
                '        MapNpc(y, x).Num = 0
                '        MapNpc(y, x).SpawnWait = TickCount
                '   End If
                'End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(Y, X).num = 0 And Map(Y).Npc(X) > 0 Then If TickCount > MapNpc(Y, X).SpawnWait + (Npc(Map(Y).Npc(X)).SpawnSecs * 1000) Then Call SpawnNpc(X, Y)
                If MapNpc(Y, X).num > 0 Then Call SendDataToMap(Y, "npchp" & SEP_CHAR & X & SEP_CHAR & MapNpc(Y, X).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(Y, X).num) & SEP_CHAR & END_CHAR)
                If MapNpc(Y, X).num > 0 Then Call SendDataToMap(Y, "npcmp" & SEP_CHAR & X & SEP_CHAR & MapNpc(Y, X).MP & SEP_CHAR & GetNpcMaxMP(MapNpc(Y, X).num) & SEP_CHAR & END_CHAR)
            Next X
            
        End If
        DoEvents
    Next Y
    
    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then GiveNPCHPTimer = GetTickCount

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then KeyTimer = GetTickCount
End Sub

Sub CheckGiveHP()
Dim i As Long, n As Long

    If GetTickCount > GiveHPTimer + 10000 Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                Call SetPlayerHP(i, GetPlayerHP(i) + GetPlayerHPRegen(i))
                Call SendHP(i)
                Call SetPlayerMP(i, GetPlayerMP(i) + GetPlayerMPRegen(i))
                Call SendMP(i)
                Call SetPlayerSP(i, GetPlayerSP(i) + GetPlayerSPRegen(i))
                Call SendSP(i)
            End If
            DoEvents
        Next i
        
        GiveHPTimer = GetTickCount
    End If
End Sub

Sub VerifEffetsJoueur()
Dim i As Long
Dim z As Long
For i = 1 To MAX_PLAYERS
    If bouclier(i) And GetTickCount >= BouclierT(i) Then bouclier(i) = False: BouclierT(i) = 0
    If Para(i) And GetTickCount >= ParaT(i) Then Call ContrOnOff(i): Para(i) = False: ParaT(i) = 0
    If Point(i) > 0 And Point(i) < MAX_SPELLS Then
    If Spell(Point(i)).type = SPELL_TYPE_AMELIO And GetTickCount >= PointT(i) Then
        Player(i).Char(Player(i).CharNum).def = Player(i).Char(Player(i).CharNum).def - Val(Spell(Point(i)).data3)
        Player(i).Char(Player(i).CharNum).magi = Player(i).Char(Player(i).CharNum).magi - Val(Spell(Point(i)).data3)
        Player(i).Char(Player(i).CharNum).STR = Player(i).Char(Player(i).CharNum).STR - Val(Spell(Point(i)).data3)
        Player(i).Char(Player(i).CharNum).Speed = Player(i).Char(Player(i).CharNum).Speed - Val(Spell(Point(i)).data3)
        Call SendStats(i)
        Point(i) = 0
        PointT(i) = 0
    ElseIf Spell(Point(i)).type = SPELL_TYPE_DECONC And GetTickCount >= PointT(i) Then
        Player(i).Char(Player(i).CharNum).def = Player(i).Char(Player(i).CharNum).def + Val(Spell(Point(i)).data3)
        Player(i).Char(Player(i).CharNum).magi = Player(i).Char(Player(i).CharNum).magi + Val(Spell(Point(i)).data3)
        Player(i).Char(Player(i).CharNum).STR = Player(i).Char(Player(i).CharNum).STR + Val(Spell(Point(i)).data3)
        Player(i).Char(Player(i).CharNum).Speed = Player(i).Char(Player(i).CharNum).Speed + Val(Spell(Point(i)).data3)
        Call SendStats(i)
        Point(i) = 0
        PointT(i) = 0
    End If
    End If
Next i
For i = 1 To MAX_MAPS
    For z = 1 To MAX_MAP_NPCS
        If ParaN(z, i) And GetTickCount >= ParaNT(z, i) Then Call PNJOnOff(z, i): ParaN(z, i) = False: ParaNT(z, i) = 0
    Next z
Next i
End Sub

Sub PlayerSaveTimer()
Static MinPassed As Long
Dim i As Long

MinPassed = MinPassed + 1
If MinPassed >= 60 Then
    If TotalOnlinePlayers > 0 Then
        PlayerI = 1
        frmServer.PlayerTimer.Enabled = True
        frmServer.tmrPlayerSave.Enabled = False
    End If

    MinPassed = 0
End If

End Sub

Sub ChargOptCoul()
If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "AccAdmin")) <> vbNullString Then AccAdmin = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "AccAdmin")): frmOptCoul.adm.BackColor = AccAdmin
If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "AccDevelopeur")) <> vbNullString Then AccDevelopeur = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "AccDevelopeur")): frmOptCoul.dev.BackColor = AccDevelopeur
If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "AccModo")) <> vbNullString Then AccModo = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "AccModo")): frmOptCoul.modo.BackColor = AccModo
If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "AccMapeur")) <> vbNullString Then AccMapeur = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "AccMapeur")): frmOptCoul.mapp.BackColor = AccMapeur
If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgDiscu")) <> vbNullString Then SayColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgDiscu")): frmOptCoul.MsgC(0).BackColor = SayColor
If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgGlob")) <> vbNullString Then GlobalColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgGlob")): frmOptCoul.MsgC(1).BackColor = GlobalColor
If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgDist")) <> vbNullString Then TellColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgDist")): frmOptCoul.MsgC(2).BackColor = TellColor
If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgHurl")) <> vbNullString Then BroadcastColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgHurl")): frmOptCoul.MsgC(3).BackColor = BroadcastColor
If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgEmot")) <> vbNullString Then EmoteColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgEmot")): frmOptCoul.MsgC(4).BackColor = EmoteColor
If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgAdmin")) <> vbNullString Then AdminColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgAdmin")): frmOptCoul.MsgC(5).BackColor = AdminColor
If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgAide")) <> vbNullString Then HelpColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgAide")): frmOptCoul.MsgC(6).BackColor = HelpColor
If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgQui")) <> vbNullString Then WhoColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgQui")): frmOptCoul.MsgC(7).BackColor = WhoColor
If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgDep")) <> vbNullString Then JoinLeftColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgDep")): frmOptCoul.MsgC(8).BackColor = JoinLeftColor
If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgAlert")) <> vbNullString Then AlertColor = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgAlert")): frmOptCoul.MsgC(9).BackColor = AlertColor
If Trim$(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgGuilde")) <> vbNullString Then CouleurDesGuilde = Val(GetVar(App.Path & "\Data.ini", "COULEURS", "MsgGuilde")): frmOptCoul.MsgC(10).BackColor = CouleurDesGuilde
frmOptCoul.Show vbModeless, frmServer
End Sub

Public Function ValidTarget(ByVal value As Long, ByVal MapNum As Long, ByVal TType As Byte) As Boolean
    Select Case TType
        Case TARGET_TYPE_PLAYER: If value > 0 And value <= MAX_PLAYERS Then If IsPlaying(value) And GetPlayerMap(value) = MapNum Then ValidTarget = True
        Case TARGET_TYPE_NPC: If value > 0 And value < MAX_MAP_NPCS Then If MapNpc(MapNum, value).num > 0 Then ValidTarget = True
        Case TARGET_TYPE_CASE: If value >= 0 And value <= (MAX_MAPX + 1) * (MAX_MAPY + 1) Then ValidTarget = True
    End Select
End Function

Public Function NpcBeside(ByVal Map As Long, ByVal MapNpc1 As Byte, ByVal MapNpc2 As Byte) As Boolean
On Error Resume Next

NpcBeside = False
If MapNpc1 < 1 Or MapNpc1 > MAX_MAP_NPCS Or MapNpc2 < 1 Or MapNpc2 > MAX_MAP_NPCS Then Exit Function

If MapNpc(Map, MapNpc1).X - 1 = MapNpc(Map, MapNpc2).X And MapNpc(Map, MapNpc1).Y = MapNpc(Map, MapNpc2).Y Then NpcBeside = True: Exit Function
If MapNpc(Map, MapNpc1).X = MapNpc(Map, MapNpc2).X And MapNpc(Map, MapNpc1).Y - 1 = MapNpc(Map, MapNpc2).Y Then NpcBeside = True: Exit Function
If MapNpc(Map, MapNpc1).X = MapNpc(Map, MapNpc2).X And MapNpc(Map, MapNpc1).Y + 1 = MapNpc(Map, MapNpc2).Y Then NpcBeside = True: Exit Function
If MapNpc(Map, MapNpc1).X + 1 = MapNpc(Map, MapNpc2).X And MapNpc(Map, MapNpc1).Y = MapNpc(Map, MapNpc2).Y Then NpcBeside = True: Exit Function
End Function

Public Sub SelectMoveNpc(ByVal value As Byte, ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal index As Long, ByVal IndexType As Long, DidWalk As Boolean)
Dim i As Byte, TmpX As Byte, TmpY As Byte
Select Case value
    Case 0
        If IndexType = TARGET_TYPE_PLAYER Then
            ' Up
            If MapNpc(MapNum, MapNpcNum).Y > GetPlayerY(index) And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then Call NpcMove(MapNum, MapNpcNum, DIR_UP, MOVING_WALKING): DidWalk = True
            ' Down
            ElseIf MapNpc(MapNum, MapNpcNum).Y < GetPlayerY(index) And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then Call NpcMove(MapNum, MapNpcNum, DIR_DOWN, MOVING_WALKING): DidWalk = True
            ' Left
            ElseIf MapNpc(MapNum, MapNpcNum).X > GetPlayerX(index) And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then Call NpcMove(MapNum, MapNpcNum, DIR_LEFT, MOVING_WALKING): DidWalk = True
            ' Right
            ElseIf MapNpc(MapNum, MapNpcNum).X < GetPlayerX(index) And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then Call NpcMove(MapNum, MapNpcNum, DIR_RIGHT, MOVING_WALKING): DidWalk = True
            End If
            'Débloquer
            If Not DidWalk And Not ACoter(MapNpcNum, index) Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then Call NpcMove(MapNum, MapNpcNum, DIR_DOWN, MOVING_WALKING): DidWalk = True
            End If
        ElseIf IndexType = TARGET_TYPE_NPC Then
            ' Up
            If MapNpc(MapNum, MapNpcNum).Y > MapNpc(MapNum, index).Y And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then Call NpcMove(MapNum, MapNpcNum, DIR_UP, MOVING_WALKING): DidWalk = True
            ' Down
            ElseIf MapNpc(MapNum, MapNpcNum).Y < MapNpc(MapNum, index).Y And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then Call NpcMove(MapNum, MapNpcNum, DIR_DOWN, MOVING_WALKING): DidWalk = True
            ' Left
            ElseIf MapNpc(MapNum, MapNpcNum).X > MapNpc(MapNum, index).X And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then Call NpcMove(MapNum, MapNpcNum, DIR_LEFT, MOVING_WALKING): DidWalk = True
            ' Right
            ElseIf MapNpc(MapNum, MapNpcNum).X < MapNpc(MapNum, index).X And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then Call NpcMove(MapNum, MapNpcNum, DIR_RIGHT, MOVING_WALKING): DidWalk = True
            End If
            'Débloquer
            If Not DidWalk And Not NpcBeside(MapNum, MapNpcNum, index) Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then Call NpcMove(MapNum, MapNpcNum, DIR_DOWN, MOVING_WALKING): DidWalk = True
            End If
        End If

    Case 1
        If IndexType = TARGET_TYPE_PLAYER Then
            ' Right
            If MapNpc(MapNum, MapNpcNum).X < GetPlayerX(index) And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then Call NpcMove(MapNum, MapNpcNum, DIR_RIGHT, MOVING_WALKING): DidWalk = True
            ' Left
            ElseIf MapNpc(MapNum, MapNpcNum).X > GetPlayerX(index) And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then Call NpcMove(MapNum, MapNpcNum, DIR_LEFT, MOVING_WALKING): DidWalk = True
            ' Down
            ElseIf MapNpc(MapNum, MapNpcNum).Y < GetPlayerY(index) And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then Call NpcMove(MapNum, MapNpcNum, DIR_DOWN, MOVING_WALKING): DidWalk = True
            ' Up
            ElseIf MapNpc(MapNum, MapNpcNum).Y > GetPlayerY(index) And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then Call NpcMove(MapNum, MapNpcNum, DIR_UP, MOVING_WALKING): DidWalk = True
            End If
            'Débloquer
            If Not DidWalk And Not ACoter(MapNpcNum, index) Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then Call NpcMove(MapNum, MapNpcNum, DIR_LEFT, MOVING_WALKING): DidWalk = True
            End If
        ElseIf IndexType = TARGET_TYPE_NPC Then
            ' Right
            If MapNpc(MapNum, MapNpcNum).X < MapNpc(MapNum, index).X And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then Call NpcMove(MapNum, MapNpcNum, DIR_RIGHT, MOVING_WALKING): DidWalk = True
            ' Left
            ElseIf MapNpc(MapNum, MapNpcNum).X > MapNpc(MapNum, index).X And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then Call NpcMove(MapNum, MapNpcNum, DIR_LEFT, MOVING_WALKING): DidWalk = True
            ' Down
            ElseIf MapNpc(MapNum, MapNpcNum).Y < MapNpc(MapNum, index).Y And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then Call NpcMove(MapNum, MapNpcNum, DIR_DOWN, MOVING_WALKING): DidWalk = True
            ' Up
            ElseIf MapNpc(MapNum, MapNpcNum).Y > MapNpc(MapNum, index).Y And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then Call NpcMove(MapNum, MapNpcNum, DIR_UP, MOVING_WALKING): DidWalk = True
            End If
            'Débloquer
            If Not DidWalk And Not NpcBeside(MapNum, MapNpcNum, index) Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then Call NpcMove(MapNum, MapNpcNum, DIR_LEFT, MOVING_WALKING): DidWalk = True
            End If
        End If

    Case 2
        If IndexType = TARGET_TYPE_PLAYER Then
            ' Down
            If MapNpc(MapNum, MapNpcNum).Y < GetPlayerY(index) And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then Call NpcMove(MapNum, MapNpcNum, DIR_DOWN, MOVING_WALKING): DidWalk = True
            ' Up
            ElseIf MapNpc(MapNum, MapNpcNum).Y > GetPlayerY(index) And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then Call NpcMove(MapNum, MapNpcNum, DIR_UP, MOVING_WALKING): DidWalk = True
            ' Right
            ElseIf MapNpc(MapNum, MapNpcNum).X < GetPlayerX(index) And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then Call NpcMove(MapNum, MapNpcNum, DIR_RIGHT, MOVING_WALKING): DidWalk = True
            ' Left
            ElseIf MapNpc(MapNum, MapNpcNum).X > GetPlayerX(index) And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then Call NpcMove(MapNum, MapNpcNum, DIR_LEFT, MOVING_WALKING): DidWalk = True
            End If
            'Débloquer
            If Not DidWalk And Not ACoter(MapNpcNum, index) Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then Call NpcMove(MapNum, MapNpcNum, DIR_UP, MOVING_WALKING): DidWalk = True
            End If
        ElseIf IndexType = TARGET_TYPE_NPC Then
            ' Down
            If MapNpc(MapNum, MapNpcNum).Y < MapNpc(MapNum, index).Y And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then Call NpcMove(MapNum, MapNpcNum, DIR_DOWN, MOVING_WALKING): DidWalk = True
            ' Up
            ElseIf MapNpc(MapNum, MapNpcNum).Y > MapNpc(MapNum, index).Y And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then Call NpcMove(MapNum, MapNpcNum, DIR_UP, MOVING_WALKING): DidWalk = True
            ' Right
            ElseIf MapNpc(MapNum, MapNpcNum).X < MapNpc(MapNum, index).X And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then Call NpcMove(MapNum, MapNpcNum, DIR_RIGHT, MOVING_WALKING): DidWalk = True
            ' Left
            ElseIf MapNpc(MapNum, MapNpcNum).X > MapNpc(MapNum, index).X And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then Call NpcMove(MapNum, MapNpcNum, DIR_LEFT, MOVING_WALKING): DidWalk = True
            End If
            'Débloquer
            If Not DidWalk And Not NpcBeside(MapNum, MapNpcNum, index) Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then Call NpcMove(MapNum, MapNpcNum, DIR_UP, MOVING_WALKING): DidWalk = True
            End If
        End If

    Case 3
        If IndexType = TARGET_TYPE_PLAYER Then
            ' Left
            If MapNpc(MapNum, MapNpcNum).X > GetPlayerX(index) And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then Call NpcMove(MapNum, MapNpcNum, DIR_LEFT, MOVING_WALKING): DidWalk = True
            ' Right
            ElseIf MapNpc(MapNum, MapNpcNum).X < GetPlayerX(index) And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then Call NpcMove(MapNum, MapNpcNum, DIR_RIGHT, MOVING_WALKING): DidWalk = True
            ' Up
            ElseIf MapNpc(MapNum, MapNpcNum).Y > GetPlayerY(index) And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then Call NpcMove(MapNum, MapNpcNum, DIR_UP, MOVING_WALKING): DidWalk = True
            ' Down
            ElseIf MapNpc(MapNum, MapNpcNum).Y < GetPlayerY(index) And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then Call NpcMove(MapNum, MapNpcNum, DIR_DOWN, MOVING_WALKING): DidWalk = True
            End If
            'Débloquer
            If Not DidWalk And Not ACoter(MapNpcNum, index) Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then Call NpcMove(MapNum, MapNpcNum, DIR_RIGHT, MOVING_WALKING): DidWalk = True
            End If
        ElseIf IndexType = TARGET_TYPE_NPC Then
            ' Left
            If MapNpc(MapNum, MapNpcNum).X > MapNpc(MapNum, index).X And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then Call NpcMove(MapNum, MapNpcNum, DIR_LEFT, MOVING_WALKING): DidWalk = True
            ' Right
            ElseIf MapNpc(MapNum, MapNpcNum).X < MapNpc(MapNum, index).X And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then Call NpcMove(MapNum, MapNpcNum, DIR_RIGHT, MOVING_WALKING): DidWalk = True
            ' Up
            ElseIf MapNpc(MapNum, MapNpcNum).Y > MapNpc(MapNum, index).Y And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then Call NpcMove(MapNum, MapNpcNum, DIR_UP, MOVING_WALKING): DidWalk = True
            ' Down
            ElseIf MapNpc(MapNum, MapNpcNum).Y < MapNpc(MapNum, index).Y And DidWalk = False Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then Call NpcMove(MapNum, MapNpcNum, DIR_DOWN, MOVING_WALKING): DidWalk = True
            End If
            'Débloquer
            If Not DidWalk And Not NpcBeside(MapNum, MapNpcNum, index) Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then Call NpcMove(MapNum, MapNpcNum, DIR_RIGHT, MOVING_WALKING): DidWalk = True
            End If
        End If
End Select
If Not DidWalk Then
    If IndexType = TARGET_TYPE_PLAYER Then
        If MapNpc(MapNum, MapNpcNum).X - 1 = GetPlayerX(index) And MapNpc(MapNum, MapNpcNum).Y = GetPlayerY(index) Then
            If MapNpc(MapNum, MapNpcNum).Dir <> DIR_LEFT Then Call NpcDir(MapNum, MapNpcNum, DIR_LEFT)
            DidWalk = True
        End If
        If MapNpc(MapNum, MapNpcNum).X + 1 = GetPlayerX(index) And MapNpc(MapNum, MapNpcNum).Y = GetPlayerY(index) Then
            If MapNpc(MapNum, MapNpcNum).Dir <> DIR_RIGHT Then Call NpcDir(MapNum, MapNpcNum, DIR_RIGHT)
            DidWalk = True
        End If
        If MapNpc(MapNum, MapNpcNum).X = GetPlayerX(index) And MapNpc(MapNum, MapNpcNum).Y - 1 = GetPlayerY(index) Then
            If MapNpc(MapNum, MapNpcNum).Dir <> DIR_UP Then Call NpcDir(MapNum, MapNpcNum, DIR_UP)
            DidWalk = True
        End If
        If MapNpc(MapNum, MapNpcNum).X = GetPlayerX(index) And MapNpc(MapNum, MapNpcNum).Y + 1 = GetPlayerY(index) Then
            If MapNpc(MapNum, MapNpcNum).Dir <> DIR_DOWN Then Call NpcDir(MapNum, MapNpcNum, DIR_DOWN)
            DidWalk = True
        End If
    ElseIf IndexType = TARGET_TYPE_NPC Then
        If MapNpc(MapNum, MapNpcNum).X - 1 = MapNpc(MapNum, index).X And MapNpc(MapNum, MapNpcNum).Y = MapNpc(MapNum, index).Y Then
            If MapNpc(MapNum, MapNpcNum).Dir <> DIR_LEFT Then Call NpcDir(MapNum, MapNpcNum, DIR_LEFT)
            DidWalk = True
        End If
        If MapNpc(MapNum, MapNpcNum).X + 1 = MapNpc(MapNum, index).X And MapNpc(MapNum, MapNpcNum).Y = MapNpc(MapNum, index).Y Then
            If MapNpc(MapNum, MapNpcNum).Dir <> DIR_RIGHT Then Call NpcDir(MapNum, MapNpcNum, DIR_RIGHT)
            DidWalk = True
        End If
        If MapNpc(MapNum, MapNpcNum).X = MapNpc(MapNum, index).X And MapNpc(MapNum, MapNpcNum).Y - 1 = MapNpc(MapNum, index).Y Then
            If MapNpc(MapNum, MapNpcNum).Dir <> DIR_UP Then Call NpcDir(MapNum, MapNpcNum, DIR_UP)
            DidWalk = True
        End If
        If MapNpc(MapNum, MapNpcNum).X = MapNpc(MapNum, index).X And MapNpc(MapNum, MapNpcNum).Y + 1 = MapNpc(MapNum, index).Y Then
            If MapNpc(MapNum, MapNpcNum).Dir <> DIR_DOWN Then Call NpcDir(MapNum, MapNpcNum, DIR_DOWN)
            DidWalk = True
        End If
    End If

    ' We could not move so player must be behind something, walk randomly.
    If Not DidWalk Then
        If Map(MapNum).Npcs(MapNpcNum).X <= -1 Then
            i = Int(Rnd * 2)
            If i = 1 Then
                i = Int(Rnd * 4)
                If CanNpcMove(MapNum, MapNpcNum, i) Then Call NpcMove(MapNum, MapNpcNum, i, MOVING_WALKING)
            End If
        End If
    End If
End If
End Sub

Function CanNPCAttackNPC(ByVal MapNum As Integer, ByVal MapNpcNumAtt As Byte, ByVal MapNpcNumDef As Byte) As Boolean
Dim npcnum As Integer
    
    CanNPCAttackNPC = False
    
    On Error GoTo er:
    
    ' Check for subscript out of range
    If MapNpcNumAtt <= 0 Or MapNpcNumAtt > MAX_MAP_NPCS Or MapNpcNumDef <= 0 Or MapNpcNumDef > MAX_MAP_NPCS Then Exit Function
        
    ' Check for subscript out of range
    If MapNpc(MapNum, MapNpcNumAtt).num <= 0 Or MapNpc(MapNum, MapNpcNumAtt).num > MAX_NPCS Or MapNpc(MapNum, MapNpcNumDef).num <= 0 Or MapNpc(MapNum, MapNpcNumDef).num > MAX_NPCS Then Exit Function
    
    npcnum = MapNpc(MapNum, MapNpcNumAtt).num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum, MapNpcNumAtt).HP <= 0 And CLng(Npc(npcnum).Inv) = 0 Or MapNpc(MapNum, MapNpcNumDef).HP <= 0 And CLng(Npc(MapNpc(MapNum, MapNpcNumDef).num).Inv) = 0 Then Exit Function
        
    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(MapNum, MapNpcNumAtt).AttackTimer + 1000 Then Exit Function
    
    MapNpc(MapNum, MapNpcNumAtt).AttackTimer = GetTickCount
            ' Check if at same coordinates
            If (MapNpc(MapNum, MapNpcNumDef).Y + 1 = MapNpc(MapNum, MapNpcNumAtt).Y) And (MapNpc(MapNum, MapNpcNumDef).X = MapNpc(MapNum, MapNpcNumAtt).X) Then
                CanNPCAttackNPC = True
            Else
                If (MapNpc(MapNum, MapNpcNumDef).Y - 1 = MapNpc(MapNum, MapNpcNumAtt).Y) And (MapNpc(MapNum, MapNpcNumDef).X = MapNpc(MapNum, MapNpcNumAtt).X) Then
                    CanNPCAttackNPC = True
                Else
                    If (MapNpc(MapNum, MapNpcNumDef).Y = MapNpc(MapNum, MapNpcNumAtt).Y) And (MapNpc(MapNum, MapNpcNumDef).X + 1 = MapNpc(MapNum, MapNpcNumAtt).X) Then
                        CanNPCAttackNPC = True
                    Else
                        If (MapNpc(MapNum, MapNpcNumDef).Y = MapNpc(MapNum, MapNpcNumAtt).Y) And (MapNpc(MapNum, MapNpcNumDef).X - 1 = MapNpc(MapNum, MapNpcNumAtt).X) Then
                            CanNPCAttackNPC = True
                        End If
                    End If
                End If
            End If

Exit Function
er:
CanNPCAttackNPC = False
On Error Resume Next
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur dans l'attaque d'un PNJ(" & MapNpc(MapNum, MapNpcNumDef).num & ")par un PNJ(" & npcnum & "). Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur dans l'attaque d'un PNJ(" & MapNpc(MapNum, MapNpcNumDef).num & ")par un PNJ(" & npcnum & ")", BrightRed, True)
End Function

Sub NPCAttackNPC(ByVal MapNum As Integer, ByVal Damage As Integer, ByVal MapNpcNumAtt As Byte, ByVal MapNpcNumDef As Byte)
Dim AttNpcNum As Integer, DefNpcNum As Integer

    On Error GoTo er
    
    ' Check for subscript out of range
    If MapNpcNumDef <= 0 Or MapNpcNumDef > MAX_MAP_NPCS Or MapNpcNumAtt <= 0 Or MapNpcNumAtt > MAX_MAP_NPCS Or Damage < 0 Then Exit Sub
    If MapNpc(MapNum, MapNpcNumDef).num <= 0 Or MapNpc(MapNum, MapNpcNumAtt).num <= 0 Then Exit Sub

    ' Send this packet so they can see the person attacking
    Call SendDataToMap(MapNum, "NPCATTACKNPC" & SEP_CHAR & MapNpcNumAtt & SEP_CHAR & END_CHAR)
    
    If Damage >= MapNpc(MapNum, MapNpcNumDef).HP Then
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum, MapNpcNumDef).num = 0
        MapNpc(MapNum, MapNpcNumDef).SpawnWait = GetTickCount
        MapNpc(MapNum, MapNpcNumDef).HP = 0
        Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNumDef & SEP_CHAR & END_CHAR)
        
        ' Set NPC target to 0
        MapNpc(MapNum, MapNpcNumDef).Target = 0
        MapNpc(MapNum, MapNpcNumDef).TargetType = 0
        MapNpc(MapNum, MapNpcNumAtt).Target = 0
        MapNpc(MapNum, MapNpcNumAtt).TargetType = 0
    Else
        MapNpc(MapNum, MapNpcNumDef).HP = MapNpc(MapNum, MapNpcNumDef).HP - Damage
        If MapNpc(MapNum, MapNpcNumDef).Target <> MapNpcNumAtt Then MapNpc(MapNum, MapNpcNumDef).Target = MapNpcNumAtt: MapNpc(MapNum, MapNpcNumDef).TargetType = TARGET_TYPE_NPC
    End If
    
    'Call SendDataTomap(mapnum, "NPCBLITNPCDMG" & SEP_CHAR & mapnpcnumdef & sep_char & Damage & SEP_CHAR & END_CHAR) ' --> <--
    'Call SendDataToMap(MapNum, "sound" & SEP_CHAR & "pain" & SEP_CHAR & END_CHAR)

Exit Sub
er:
On Error Resume Next

Call AddLog("le : " & Date & "     à : " & Time & "...Erreur dans l'attaque d'un PNJ(" & MapNpc(MapNum, MapNpcNumDef).num & ")par un PNJ(" & MapNpc(MapNum, MapNpcNumAtt).num & "). Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur dans l'attaque d'un PNJ(" & MapNpc(MapNum, MapNpcNumDef).num & ")par un PNJ(" & MapNpc(MapNum, MapNpcNumAtt).num & ")", BrightRed, True)
End Sub
'Script Hotel de ventes par Horace
Public Sub HdvCmd(ByVal index As Long, ByVal s As String)
Dim Parse() As String, Answer As Integer
    s = Mid$(s, 7)
    Parse = Split(s, " ")
    Select Case LCase$(Parse(0))
        Case "help", "?", "aide"
            Call PlayerMsg(index, " ---------- Aide Hotel de Ventes ---------- ", White)
            Call PlayerMsg(index, " Acheter -> /hdvs achat", White)
            Call PlayerMsg(index, " Vendre  -> /hdvs vente", White)
            Call PlayerMsg(index, " Annuler ->", White)
            Call PlayerMsg(index, " . Achat -> /hdvs sachat", White)
            Call PlayerMsg(index, " . Vente -> /hdvs svente", White)
        Case "achat", "vente"
            Parse(0) = LCase$(Parse(0))
            If UBound(Parse) >= 1 Then
                Parse(1) = LCase$(Parse(1))
                If Parse(1) = "help" Or Parse(1) = "?" Or Parse(1) = "aide" Then
                    PlayerMsg index, " --------- Aide " & Parse(0) & " --------- ", White
                    s = " /hdvs " & Parse(0) & " InvNum"
                    If Parse(0) = "achat" Then s = s & " ItemNum Val Dur"
                    PlayerMsg index, s, White
                ElseIf IsNumeric(Parse(1)) And Val(Parse(1)) > 0 And Val(Parse(1)) < MAX_INV Then
                    If Parse(0) = "achat" Then
                        If UBound(Parse) = 4 Then
                            If Val(Parse(2)) > 0 And Val(Parse(2)) <= MAX_ITEMS Then
                                Answer = HotelDeVente.AddAchat(index, Val(Parse(2)), Val(Parse(3)), Val(Parse(4)), True)
                                PlayerMsg index, "Votre achat a bien été effectué.", Green
                                PlayerMsg index, "Veuillez prendre note du numéro " & Answer, White
                                PlayerMsg index, "Il sera utile si vous souhaitez annuler votre achat.", White
                            Else
                                PlayerMsg index, "L'argument " & Parse(2) & " est invalide", 4
                            End If
                        Else
                            PlayerMsg index, "Le nombre d'arguments fournit à " & Parse(0) & " n'est pas bon.", 4
                        End If
                    Else
                        If Player(index).Char(Player(index).CharNum).Inv(Val(Parse(1))).num > 0 Then
                            Answer = HotelDeVente.AddVente(index, Player(index).Char(Player(index).CharNum).Inv(Val(Parse(1))).num, Player(index).Char(Player(index).CharNum).Inv(Val(Parse(1))).value, Player(index).Char(Player(index).CharNum).Inv(Val(Parse(1))).Dur, True)
                            PlayerMsg index, "Votre vente a bien été effectué.", Green
                            PlayerMsg index, "Veuillez prendre note du numéro " & Answer, White
                            PlayerMsg index, "Il sera utile si vous souhaitez annuler votre vente.", White
                        Else
                            PlayerMsg index, "Aucun objet n'est à cet emplacement", 4
                        End If
                    End If
                Else
                    PlayerMsg index, "L'argument fournit (" & Parse(1) & ") est invalide.", 4
                End If
            Else
                PlayerMsg index, "Veuillez fournir un argument à " & Parse(0), 4
            End If
        Case "sachat", "svente"
            Parse(0) = LCase$(Parse(0))
            If UBound(Parse) = 1 Then
                If Parse(0) = "svente" Then
                    HotelDeVente.CancelVente index, Val(Parse(1))
                Else
                    HotelDeVente.CancelAchat index, Val(Parse(1))
                End If
            Else
                PlayerMsg index, "Veuillez fournir un seul argument à " & Parse(0), 4
            End If
    End Select
End Sub
