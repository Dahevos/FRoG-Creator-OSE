Attribute VB_Name = "modTypes"
Option Explicit
Global PlayerI As Byte

' Winsock globals
Public GAME_PORT As Long

' General constants
Public GAME_NAME As String
Public MAX_PLAYERS As Long
Public MAX_SPELLS As Long
Public MAX_MAPS As Long
Public MAX_SHOPS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_MAP_ITEMS As Long
Public MAX_GUILDS As Long
Public MAX_GUILD_MEMBERS As Long
Public MAX_EMOTICONS As Long
Public MAX_LEVEL As Long
Public MAX_QUETES As Long
Public Scripting As Byte
Public NOOB_LEVEL As Long
Public PK_LEVEL As Long
Public MAX_PETS As Long
Public MAX_METIER As Long
Public MAX_RECETTE As Long

Public Const MAX_PARTY_MEMBERS As Byte = 20
Public Const MAX_PARTYS As Byte = 20
Public Const MAX_HDV_TRADES As Byte = 5
Public Const MAX_ARROWS = 100
Public Const MAX_INV = 26
Public Const MAX_MAP_NPCS = 15
Public Const MAX_PLAYER_SPELLS = 20
Public Const MAX_TRADES = 66
Public Const MAX_PLAYER_TRADES = 8
Public Const MAX_NPC_DROPS = 10
Public Const MAX_NPC_SPELLS = 10

Public Const NO = 0
Public Const YES = 1

' Account constants
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 3

' Basic Security Passwords, You cant connect without it
Public Const SEC_CODE1 = "jwehiehfojcvnvnsdinaoiwheoewyriusdyrflsdjncjkxzncisdughfusyfuapsipiuahfpaijnflkjnvjnuahguiryasbdlfkjblsahgfauygewuifaunfauf"
Public Const SEC_CODE2 = "ksisyshentwuegeguigdfjkldsnoksamdihuehfidsuhdushdsisjsyayejrioehdoisahdjlasndowijapdnaidhaioshnksfnifohaifhaoinfiwnfinsaihfas"
Public Const SEC_CODE3 = "saiugdapuigoihwbdpiaugsdcapvhvinbudhbpidusbnvduisysayaspiufhpijsanfioasnpuvnupashuasohdaiofhaosifnvnuvnuahiosaodiubasdi"
Public Const SEC_CODE4 = "88978465734619123425676749756722829121973794379467987945762347631462572792798792492416127957989742945642672"

' Sex constants
Public Const SEX_MALE = 0
Public Const SEX_FEMALE = 1

' Map constants
'Public Const MAX_MAPX = 30
'Public Const MAX_MAPY = 30
Public MAX_MAPX As Long
Public MAX_MAPY As Long
Public Const MAP_MORAL_NONE = 0
Public Const MAP_MORAL_SAFE = 1
Public Const MAP_MORAL_NO_PENALTY = 2

' Image constants
Public Const PIC_X = 32
Public Const PIC_Y = 32
Public PIC_PL As Byte
Public PIC_NPC1 As Byte
Public PIC_NPC2 As Byte

' Tile consants
Public Const TILE_TYPE_WALKABLE = 0
Public Const TILE_TYPE_BLOCKED = 1
Public Const TILE_TYPE_WARP = 2
Public Const TILE_TYPE_ITEM = 3
Public Const TILE_TYPE_NPCAVOID = 4
Public Const TILE_TYPE_KEY = 5
Public Const TILE_TYPE_KEYOPEN = 6
Public Const TILE_TYPE_HEAL = 7
Public Const TILE_TYPE_KILL = 8
Public Const TILE_TYPE_SHOP = 9
Public Const TILE_TYPE_CBLOCK = 10
Public Const TILE_TYPE_ARENA = 11
Public Const TILE_TYPE_SOUND = 12
Public Const TILE_TYPE_SPRITE_CHANGE = 13
Public Const TILE_TYPE_SIGN = 14
Public Const TILE_TYPE_DOOR = 15
Public Const TILE_TYPE_NOTICE = 16
Public Const TILE_TYPE_CHEST = 17
Public Const TILE_TYPE_CLASS_CHANGE = 18
Public Const TILE_TYPE_SCRIPTED = 19
Public Const TILE_TYPE_NPC_SPAWN = 20
Public Const TILE_TYPE_BANK = 21
Public Const TILE_TYPE_COFFRE = 22
Public Const TILE_TYPE_PORTE_CODE = 23
Public Const TILE_TYPE_BLOCK_MONTURE = 24
Public Const TILE_TYPE_BLOCK_NIVEAUX = 25
Public Const TILE_TYPE_TOIT = 26
Public Const TILE_TYPE_BLOCK_GUILDE = 27
Public Const TILE_TYPE_BLOCK_TOIT = 28
Public Const TILE_TYPE_BLOCK_DIR = 29
Public Const TILE_TYPE_CRAFT As Byte = 30
Public Const TILE_TYPE_METIER As Byte = 31

' quetes constant
Public Const QUETE_TYPE_AUCUN = 0
Public Const QUETE_TYPE_RECUP = 1
Public Const QUETE_TYPE_APORT = 2
Public Const QUETE_TYPE_PARLER = 3
Public Const QUETE_TYPE_TUER = 4
Public Const QUETE_TYPE_FINIR = 5
Public Const QUETE_TYPE_GAGNE_XP = 6
Public Const QUETE_TYPE_SCRIPT = 7
Public Const QUETE_TYPE_MINIQUETE = 8

' Item constants
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_WEAPON As Byte = 1
Public Const ITEM_TYPE_ARMOR As Byte = 2
Public Const ITEM_TYPE_HELMET As Byte = 3
Public Const ITEM_TYPE_SHIELD As Byte = 4
Public Const ITEM_TYPE_POTIONADDHP As Byte = 5
Public Const ITEM_TYPE_POTIONADDMP As Byte = 6
Public Const ITEM_TYPE_POTIONADDSP As Byte = 7
Public Const ITEM_TYPE_POTIONSUBHP As Byte = 8
Public Const ITEM_TYPE_POTIONSUBMP As Byte = 9
Public Const ITEM_TYPE_POTIONSUBSP As Byte = 10
Public Const ITEM_TYPE_KEY As Byte = 11
Public Const ITEM_TYPE_CURRENCY As Byte = 12
Public Const ITEM_TYPE_SPELL As Byte = 13
Public Const ITEM_TYPE_MONTURE As Byte = 14
Public Const ITEM_TYPE_SCRIPT As Byte = 15
Public Const ITEM_TYPE_PET As Byte = 16

Public Const ITEM_TYPEARME_NONE As Byte = 0
Public Const ITEM_TYPEARME_EPEES As Byte = 1
Public Const ITEM_TYPEARME_HACHES As Byte = 2
Public Const ITEM_TYPEARME_DAGUES As Byte = 3
Public Const ITEM_TYPEARME_FAUX As Byte = 4
Public Const ITEM_TYPEARME_MARTEAUX As Byte = 5
Public Const ITEM_TYPEARME_PIOCHES As Byte = 6
Public Const ITEM_TYPEARME_PELLES As Byte = 7
Public Const ITEM_TYPEARME_BATONS As Byte = 8
Public Const ITEM_TYPEARME_BAGUETTES As Byte = 9
Public Const ITEM_TYPEARME_OUTILLAGE As Byte = 10
Public Const ITEM_TYPEARME_ARC As Byte = 11

' Metier
Public Const METIER_CHASSEUR As Byte = 0
Public Const METIER_CRAFT As Byte = 1

' Direction constants
Public Const DIR_UP = 3
Public Const DIR_DOWN = 0
Public Const DIR_LEFT = 1
Public Const DIR_RIGHT = 2

' Constants for player movement
Public Const MOVING_WALKING = 1
Public Const MOVING_RUNNING = 2

' Weather constants
Public Const WEATHER_NONE = 0
Public Const WEATHER_RAINING = 1
Public Const WEATHER_SNOWING = 2
Public Const WEATHER_THUNDER = 3

' Time constants
Public Const TIME_DAY = 0
Public Const TIME_NIGHT = 1

' Admin constants
Public Const ADMIN_MONITER = 1
Public Const ADMIN_MAPPER = 2
Public Const ADMIN_DEVELOPER = 3
Public Const ADMIN_CREATOR = 4

' NPC constants
Public Const NPC_BEHAVIOR_ATTACKONSIGHT = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED = 1
Public Const NPC_BEHAVIOR_FRIENDLY = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER = 3
Public Const NPC_BEHAVIOR_GUARD = 4
Public Const NPC_BEHAVIOR_QUETEUR = 5
Public Const NPC_BEHAVIOR_SCRIPT = 6

' Spell constants
Public Const SPELL_TYPE_ADDHP = 0
Public Const SPELL_TYPE_ADDMP = 1
Public Const SPELL_TYPE_ADDSP = 2
Public Const SPELL_TYPE_SUBHP = 3
Public Const SPELL_TYPE_SUBMP = 4
Public Const SPELL_TYPE_SUBSP = 5
'Public Const SPELL_TYPE_GIVEITEM = 7
Public Const SPELL_TYPE_SCRIPT = 6
Public Const SPELL_TYPE_AMELIO = 7
Public Const SPELL_TYPE_DECONC = 8
Public Const SPELL_TYPE_PARALY = 9
Public Const SPELL_TYPE_DEFENC = 10
Public Const SPELL_TYPE_TELE = 11 'type ajouter à l'éditeur

' Target type constants
Public Const TARGET_TYPE_PLAYER = 0
Public Const TARGET_TYPE_NPC = 1
Public Const TARGET_TYPE_CASE = 2

Type IndRec
    data1 As Long
    data2 As Long
    data3 As Long
    String1 As String
End Type

Type PlayerInvRec
    num As Long
    value As Long
    Dur As Long
End Type

Type PlayerQueteRec
    temps As Long
    data1 As Long
    data2 As Long
    data3 As Long
    String1 As String
    indexe(1 To 15) As IndRec
End Type

Type PetPosRec
    X As Integer
    Y As Integer
    Dir As Byte
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Guild As String
    Guildaccess As Byte
    Sex As Byte
    Class As Long
    sprite As Long
    Level As Long
    Exp As Long
    Access As Byte
    PK As Byte

    ' Vitals
    HP As Long
    MP As Long
    SP As Long
    
    ' Stats
    STR As Long
    def As Long
    Speed As Long
    magi As Long
    POINTS As Long
    
    ' Worn equipment
    ArmorSlot As Long
    WeaponSlot As Long
    HelmetSlot As Long
    ShieldSlot As Long
    PetSlot As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    QueteStatut() As Integer
    pet As PetPosRec
    
    ' Position
    Map As Long
    X As Integer
    Y As Integer
    Dir As Byte
    
    QueteEnCour As Long
    Quetep As PlayerQueteRec
    
    'PAPERDOLL
    Casque As Long
    armure As Long
    arme As Long
    bouclier As Long
    
    'FIN PAPERDOLL
    
    vendeur As Long
    
    metier As Long
    MetierLvl As Long
    MetierExp As Long
    
End Type

Type PlayerTradeRec
    InvNum As Long
    InvName As String
    InvVal As Long
End Type
    
Type AccountRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String
       
    ' Characters (we use 0 to prevent a crash that still needs to be figured out)
    Char(0 To MAX_CHARS) As PlayerRec
    
    ' None saved local vars
    Buffer As String
    IncBuffer As String
    CharNum As Byte
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    
    PartyPlayer As Integer
    InParty As Byte
    TargetType As Byte
    Target As Long
    CastedSpell As Byte
    
    SpellTime As Long
    SpellVar As Long
    SpellDone As Long
    SpellNum As Long
    
    GettingMap As Byte
    InvitedBy As Byte
    
    Emoticon As Long

    InTrade As Byte
    TradePlayer As Long
    TradeOk As Byte
    TradeItemMax As Byte
    TradeItemMax2 As Byte
    Trading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
    
    InChat As Byte
    ChatPlayer As Long
    
    Mute As Boolean
    
    sync As Boolean
End Type

Type TileRec
    Ground As Long
    Mask As Long
    Anim As Long
    Mask2 As Long
    M2Anim As Long
    Mask3 As Long '<--
    M3Anim As Long '<--
    Fringe As Long
    FAnim As Long
    Fringe2 As Long
    F2Anim As Long
    Fringe3 As Long '<--
    F3Anim As Long '<--
    type As Byte
    data1 As Long
    data2 As Long
    data3 As Long
    String1 As String
    String2 As String
    String3 As String
    Light As Long
    GroundSet As Byte
    MaskSet As Byte
    AnimSet As Byte
    Mask2Set As Byte
    M2AnimSet As Byte
    Mask3Set As Byte '<--
    M3AnimSet As Byte '<--
    FringeSet As Byte
    FAnimSet As Byte
    Fringe2Set As Byte
    F2AnimSet As Byte
    Fringe3Set As Byte '<--
    F3AnimSet As Byte '<--
End Type

Type NpcMap
    X As Byte
    Y As Byte
    x1 As Byte
    y1 As Byte
    x2 As Byte
    y2 As Byte
    x3 As Byte
    y3 As Byte
    x4 As Byte
    y4 As Byte
    x5 As Byte
    y5 As Byte
    x6 As Byte
    y6 As Byte
    boucle As Byte
    Hasardm As Byte
    Hasardp As Byte
    Imobile As Byte
    Axy As Boolean
    Axy1 As Boolean
    Axy2 As Boolean
End Type

Type MapRec
    Name As String * 40
    Revision As Long
    Moral As Byte
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    Music As String
    BootMap As Long
    BootX As Byte
    BootY As Byte
    Shop As Long
    Indoors As Byte
    Tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Long
    Npcs(1 To MAX_MAP_NPCS) As NpcMap
    PanoInf As String * 50
    TranInf As Byte
    PanoSup As String * 50
    TranSup As Byte
    Fog As Integer
    FogAlpha As Byte
    guildSoloView As Byte
    petView As Byte
    traversable As Byte
End Type

Type RecompRec
    Exp As Long
    objn1 As Long
    objn2 As Long
    objn3 As Long
    objq1 As Long
    objq2 As Long
    objq3 As Long
End Type

Type QueteRec
    nom As String * 40
    type As Long
    description As String
    reponse As String
    temps As Long
    data1 As Long
    data2 As Long
    data3 As Long
    String1 As String
    Recompence As RecompRec
    indexe(1 To 15) As IndRec
    Case As Long
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    
    AdvanceFrom As Long
    LevelReq As Long
    type As Long
    Locked As Long
    
    MaleSprite As Long
    FemaleSprite As Long
    
    STR As Long
    def As Long
    Speed As Long
    magi As Long
    
    Map As Long
    X As Byte
    Y As Byte
End Type

Type ItemRec
    Name As String * NAME_LENGTH
    desc As String * 150
    
    Pic As Long
    type As Byte
    data1 As Long
    data2 As Long
    data3 As Long
    StrReq As Long
    DefReq As Long
    SpeedReq As Long
    ClassReq As Long
    AccessReq As Byte
    
    paperdoll As Byte
    paperdollPic As Long
    
    Empilable As Byte
    
    AddHP As Long
    AddMP As Long
    AddSP As Long
    AddStr As Long
    AddDef As Long
    AddMagi As Long
    AddSpeed As Long
    AddEXP As Long
    AttackSpeed As Long
    
    NCoul As Long
    
    Sex As Byte
    tArme As Long
End Type

Type MapItemRec
    num As Long
    value As Long
    Dur As Long
    
    X As Byte
    Y As Byte
End Type

Type NPCEditorRec
    ItemNum As Long
    ItemValue As Long
    chance As Long
End Type

Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String
    
    sprite As Long
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    
    STR  As Long
    def As Long
    Speed As Long
    magi As Long
    MaxHp As Long
    Exp As Long
    SpawnTime As Long
    
    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec
    QueteNum As Long
    Inv As Long
    Vol As Long
    Spell(1 To MAX_NPC_SPELLS) As Integer
End Type

Type AmelioRec
    Power As Integer
    Timer As Long
End Type

Type MapNpcRec
    num As Long
    
    Target As Long
    TargetType As Byte
    
    HP As Long
    MP As Long
    SP As Long
        
    X As Byte
    Y As Byte
    Dir As Integer
    
    Amelio As AmelioRec
    Immune As Long
    SpellTimer As Long
    
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
End Type

Type TradeItemRec
    GiveItem As Long
    GiveValue As Long
    GetItem As Long
    GetValue As Long
End Type

Type TradeItemsRec
    value(1 To MAX_TRADES) As TradeItemRec
End Type

Type ShopRec
    Name As String * NAME_LENGTH
    JoinSay As String * 100
    LeaveSay As String * 100
    FixesItems As Byte
    TradeItem(1 To 6) As TradeItemsRec
    FixObjet As Long
End Type
    
Type SpellRec
    Name As String * NAME_LENGTH
    ClassReq As Long
    LevelReq As Long
    MPCost As Long
    Sound As Long
    type As Long
    data1 As Long
    data2 As Long
    data3 As Long
    Range As Byte
    
    Big As Byte
    
    SpellAnim As Long
    SpellTime As Long
    SpellDone As Long
    
    SpellIco As Long
    
    AE As Long
End Type

Type TempTileRec
    DoorOpen()  As Byte
    DoorTimer As Long
End Type

Type GuildRec
    Name As String * NAME_LENGTH
    Founder As String * NAME_LENGTH
    Member() As String * NAME_LENGTH
End Type

Type EmoRec
    Pic As Long
    Command As String
End Type

Type CMRec
    Title As String
    Message As String
End Type

Type IBMsgs
    Texte As String
    Couleur As Long
    Gra As Boolean
End Type

Type PetsRec
    nom As String
    sprite As Long
    addForce As Byte
    addDefence As Byte
End Type

Type MetierRec
    nom As String
    type As Byte
    desc As String
    
    data(0 To MAX_DATA_METIER, 0 To 1) As Integer
End Type

Type RecetteRec
    nom As String
    InCraft(0 To 9, 0 To 1) As Integer
    craft(0 To 1) As Integer
End Type

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte
Public quete() As QueteRec
Public Party As clsParty
Public Map() As MapRec
Public TempTile() As TempTileRec
Public PlayersOnMap() As Long
Public Player() As AccountRec
Public Classe() As ClassRec
Public Class2() As ClassRec
Public Class3() As ClassRec
Public item() As ItemRec
Public Npc() As NpcRec
Public MapItem() As MapItemRec
Public MapNpc() As MapNpcRec
Public Shop() As ShopRec
Public Spell() As SpellRec
Public Guild() As GuildRec
Public Emoticons() As EmoRec
Public experience() As Long
Public CMessages(1 To 6) As CMRec
Public PnjMove() As Boolean
Public bouclier() As Boolean
Public BouclierT() As Long
Public Para() As Boolean
Public ParaT() As Long
Public ParaN() As Boolean
Public ParaNT() As Long
Public Point() As Long
Public PointT() As Long
Public IBMsgs(1 To 10) As IBMsgs
Public Pets() As PetsRec
Public metier() As MetierRec
Public recette() As RecetteRec

Type ArrowRec
    Name As String
    Pic As Long
    Range As Byte
End Type

Public Arrows(1 To MAX_ARROWS) As ArrowRec

Type StatRec
    Level As Long
    STR As Long
    def As Long
    magi As Long
    Speed As Long
End Type
Public AddHP As StatRec
Public AddMP As StatRec
Public AddSP As StatRec

'use for game ai
Public Axy1 As Boolean
Public Axy2 As Boolean
Public AdminMoMsg As Boolean

'utiliser pour le hacking
Public CClasses As Boolean

'utiliser pour les couleurs perso
Public AccModo As Long
Public AccMapeur As Long
Public AccDevelopeur As Long
Public AccAdmin As Long

Public HotelDeVente As clsHdV

Sub ClearTempTile()
Dim i As Long, Y As Long, X As Long

    For i = 1 To MAX_MAPS
        TempTile(i).DoorTimer = 0
        
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                TempTile(i).DoorOpen(X, Y) = NO
            Next X
        Next Y
    Next i
End Sub

Public Sub ContrOnOff(ByVal Index As Long)
Dim Packet As String

Packet = "CONOFF" & SEP_CHAR & END_CHAR

Call SendDataTo(Index, Packet)
End Sub

Public Sub PNJOnOff(ByVal Index As Long, ByVal Carte As Long)
If PnjMove(Index, Carte) = False Then PnjMove(Index, Carte) = True Else PnjMove(Index, Carte) = False
End Sub

Sub ClearClasses()
Dim i As Long

    For i = 0 To Max_Classes
        Classe(i).Name = vbNullString
        Classe(i).AdvanceFrom = 0
        Classe(i).LevelReq = 0
        Classe(i).type = 1
        Classe(i).STR = 0
        Classe(i).def = 0
        Classe(i).Speed = 0
        Classe(i).magi = 0
        Classe(i).FemaleSprite = 0
        Classe(i).MaleSprite = 0
        Classe(i).Map = 0
        Classe(i).X = 0
        Classe(i).Y = 0
    Next i
End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim i As Long
Dim n As Long
With Player(Index)
    .Login = vbNullString
    .Password = vbNullString
    
    For i = 1 To MAX_CHARS
        .Char(i).Name = vbNullString
        .Char(i).Class = 0
        .Char(i).Level = 0
        .Char(i).sprite = 0
        .Char(i).Exp = 0
        .Char(i).Access = 0
        .Char(i).PK = NO
        .Char(i).POINTS = 0
        .Char(i).Guild = vbNullString
        
        .Char(i).HP = 0
        .Char(i).MP = 0
        .Char(i).SP = 0
        
        .Char(i).STR = 0
        .Char(i).def = 0
        .Char(i).Speed = 0
        .Char(i).magi = 0
        
        For n = 1 To MAX_INV
            .Char(i).Inv(n).num = 0
            .Char(i).Inv(n).value = 0
            .Char(i).Inv(n).Dur = 0
        Next n
        
        For n = 1 To MAX_PLAYER_SPELLS
            .Char(i).Spell(n) = 0
        Next n
        
        .Char(i).ArmorSlot = 0
        .Char(i).WeaponSlot = 0
        .Char(i).HelmetSlot = 0
        .Char(i).ShieldSlot = 0
        .Char(i).PetSlot = 0
        
        .Char(i).Map = 0
        .Char(i).X = 0
        .Char(i).Y = 0
        .Char(i).Dir = 0
        
        .Char(i).pet.Dir = 0
        .Char(i).pet.X = 0
        .Char(i).pet.Y = 0
        
        .Char(i).vendeur = 0
        
        .Char(i).QueteEnCour = 0
        .Char(i).Quetep.data1 = 0
        .Char(i).Quetep.data2 = 0
        .Char(i).Quetep.data3 = 0
        .Char(i).Quetep.String1 = vbNullString
        
        .Char(i).metier = 0
        .Char(i).MetierLvl = 1
        .Char(i).MetierExp = 0
        
        For n = 1 To 15
        .Char(i).Quetep.indexe(n).data1 = 0
        .Char(i).Quetep.indexe(n).data2 = 0
        .Char(i).Quetep.indexe(n).data3 = 0
        .Char(i).Quetep.indexe(n).String1 = vbNullString
        Next n
        
        ' Temporary vars
        .Buffer = vbNullString
        .IncBuffer = vbNullString
        .CharNum = 0
        .InGame = False
        .AttackTimer = 0
        .DataTimer = 0
        .DataBytes = 0
        .DataPackets = 0
        .PartyPlayer = 0
        .InParty = 0
        .Target = -1
        .TargetType = 0
        .CastedSpell = NO
        .GettingMap = NO
        .Emoticon = -1
        .InTrade = 0
        .TradePlayer = 0
        .TradeOk = 0
        .TradeItemMax = 0
        .TradeItemMax2 = 0
        For n = 1 To MAX_PLAYER_TRADES
            .Trading(n).InvName = vbNullString
            .Trading(n).InvNum = 0
        Next n
        .ChatPlayer = 0
    Next i
End With
    
    bouclier(Index) = False
    BouclierT(Index) = 0
    Para(Index) = False
    ParaT(Index) = 0
    Point(Index) = 0
    PointT(Index) = 0
    
End Sub

Sub ClearChar(ByVal Index As Long, ByVal CharNum As Long)
Dim n As Long
With Player(Index)
    .Char(CharNum).Name = vbNullString
    .Char(CharNum).Class = 0
    .Char(CharNum).sprite = 0
    .Char(CharNum).Level = 0
    .Char(CharNum).Exp = 0
    .Char(CharNum).Access = 0
    .Char(CharNum).PK = NO
    .Char(CharNum).POINTS = 0
    .Char(CharNum).Guild = vbNullString
    
    .Char(CharNum).HP = 0
    .Char(CharNum).MP = 0
    .Char(CharNum).SP = 0
    
    .Char(CharNum).STR = 0
    .Char(CharNum).def = 0
    .Char(CharNum).Speed = 0
    .Char(CharNum).magi = 0
    
    For n = 1 To MAX_INV
        .Char(CharNum).Inv(n).num = 0
        .Char(CharNum).Inv(n).value = 0
        .Char(CharNum).Inv(n).Dur = 0
    Next n
    
    For n = 1 To MAX_PLAYER_SPELLS
        .Char(CharNum).Spell(n) = 0
    Next n
    
    For n = 1 To MAX_QUETES
        .Char(CharNum).QueteStatut(n) = 0
        
    Next
    .Char(CharNum).QueteEnCour = 0
    
    .Char(CharNum).ArmorSlot = 0
    .Char(CharNum).WeaponSlot = 0
    .Char(CharNum).HelmetSlot = 0
    .Char(CharNum).ShieldSlot = 0
    .Char(CharNum).PetSlot = 0
    
    .Char(CharNum).Map = 0
    .Char(CharNum).X = 0
    .Char(CharNum).Y = 0
    .Char(CharNum).Dir = 0
    
    .Char(CharNum).pet.Dir = 0
    .Char(CharNum).pet.X = 0
    .Char(CharNum).pet.Y = 0
End With
End Sub
    
Sub ClearItem(ByVal Index As Long)
With item(Index)
    .Name = vbNullString
    .desc = vbNullString
    
    .type = 0
    .data1 = 0
    .data2 = 0
    .data3 = 0
    .StrReq = 0
    .DefReq = 0
    .SpeedReq = 0
    .ClassReq = -1
    .AccessReq = 0
    
    .paperdoll = 0
    .paperdollPic = 0
    
    .Empilable = 0
    
    .AddHP = 0
    .AddMP = 0
    .AddSP = 0
    .AddStr = 0
    .AddDef = 0
    .AddMagi = 0
    .AddSpeed = 0
    .AddEXP = 0
    .AttackSpeed = 1000
    
    .NCoul = 0
    .tArme = 0
End With
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
End Sub

Sub ClearNpc(ByVal Index As Long)
Dim i As Long
With Npc(Index)
    .Name = vbNullString
    .AttackSay = vbNullString
    .sprite = 0
    .SpawnSecs = 0
    .Behavior = 0
    .Range = 0
    .STR = 0
    .def = 0
    .Speed = 0
    .magi = 0
    .MaxHp = 0
    .Exp = 0
    .SpawnTime = 0
    .QueteNum = 0
    .Inv = 0
    .Vol = 0
    For i = 1 To MAX_NPC_DROPS
        .ItemNPC(i).chance = 0
        .ItemNPC(i).ItemNum = 0
        .ItemNPC(i).ItemValue = 0
    Next i
    For i = 1 To MAX_NPC_SPELLS
        .Spell(i) = 0
    Next
End With
End Sub

Sub ClearNpcs()
Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next i
End Sub

Sub ClearPet(ByVal Index As Long)
With Pets(Index)
    .nom = ""
    .sprite = 0
    .addForce = 0
    .addDefence = 0
End With
End Sub

Sub ClearPets()
Dim i As Long

    For i = 1 To MAX_PETS
        Call ClearPet(i)
    Next i
End Sub

Sub ClearMetier(ByVal Index As Long)
Dim i As Long
With metier(Index)
    .nom = ""
    .type = 0
    .desc = ""
    For i = 0 To MAX_DATA_METIER
        .data(i, 0) = 0
        .data(i, 1) = 1
    Next i
End With
End Sub

Sub ClearMetiers()
Dim i As Long

    For i = 1 To MAX_METIER
        Call ClearMetier(i)
    Next i
End Sub

Sub ClearRecette(ByVal Index As Long)
Dim i As Long, z As Long
With recette(Index)
    .nom = ""
    For i = 0 To 9
        .InCraft(i, 0) = 0
        .InCraft(i, 1) = 0
    Next i
    For z = 0 To 1
        .craft(z) = 0
    Next z
End With
End Sub

Sub ClearRecettes()
Dim i As Long

    For i = 1 To MAX_RECETTE
        Call ClearRecette(i)
    Next i
End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Long)
    MapItem(MapNum, Index).num = 0
    MapItem(MapNum, Index).value = 0
    MapItem(MapNum, Index).Dur = 0
    MapItem(MapNum, Index).X = 0
    MapItem(MapNum, Index).Y = 0
End Sub

Sub ClearMapItems()
Dim X As Long
Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(X, Y)
        Next X
    Next Y
End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal MapNum As Long)
With MapNpc(MapNum, Index)
    .num = 0
    .Target = 0
    .TargetType = 0
    .Immune = 0
    .SpellTimer = 0
    .Amelio.Power = 0
    .Amelio.Timer = 0
    .HP = 0
    .MP = 0
    .SP = 0
    .X = 0
    .Y = 0
    .Dir = 0
    PnjMove(Index, MapNum) = True
    
    ' Server use only
    .SpawnWait = 0
    .AttackTimer = 0
End With
End Sub

Sub ClearMapNpcs()
Dim X As Long
Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(X, Y)
        Next X
    Next Y
End Sub

Sub ClearMap(ByVal MapNum As Long)
Dim i As Long
Dim X As Long
Dim Y As Long

With Map(MapNum)
    .Name = vbNullString
    .Revision = 0
    .Moral = 0
    .Up = 0
    .Down = 0
    .Left = 0
    .Right = 0
    .Indoors = 0
        
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            .Tile(X, Y).Ground = 0
            .Tile(X, Y).Mask = 0
            .Tile(X, Y).Anim = 0
            .Tile(X, Y).Mask2 = 0
            .Tile(X, Y).M2Anim = 0
            .Tile(X, Y).Fringe = 0
            .Tile(X, Y).FAnim = 0
            .Tile(X, Y).Fringe2 = 0
            .Tile(X, Y).F2Anim = 0
            .Tile(X, Y).type = 0
            .Tile(X, Y).data1 = 0
            .Tile(X, Y).data2 = 0
            .Tile(X, Y).data3 = 0
            .Tile(X, Y).String1 = vbNullString
            .Tile(X, Y).String2 = vbNullString
            .Tile(X, Y).String3 = vbNullString
            .Tile(X, Y).Light = 0
            .Tile(X, Y).GroundSet = 0
            .Tile(X, Y).MaskSet = 0
            .Tile(X, Y).AnimSet = 0
            .Tile(X, Y).Mask2Set = 0
            .Tile(X, Y).M2AnimSet = 0
            .Tile(X, Y).FringeSet = 0
            .Tile(X, Y).FAnimSet = 0
            .Tile(X, Y).Fringe2Set = 0
            .Tile(X, Y).F2AnimSet = 0
        Next X
    Next Y
    
    For i = 1 To MAX_MAP_NPCS
    .Npc(i) = 0
    .Npcs(i).Axy = False
    .Npcs(i).Axy1 = False
    .Npcs(i).Axy2 = False
    .Npcs(i).boucle = 0
    .Npcs(i).Hasardm = 1
    .Npcs(i).Hasardp = 1
    .Npcs(i).Imobile = 0
    .Npcs(i).X = 0
    .Npcs(i).x1 = 0
    .Npcs(i).x2 = 0
    .Npcs(i).x3 = 0
    .Npcs(i).x4 = 0
    .Npcs(i).x5 = 0
    .Npcs(i).x6 = 0
    .Npcs(i).Y = 0
    .Npcs(i).y2 = 0
    .Npcs(i).y3 = 0
    .Npcs(i).y4 = 0
    .Npcs(i).y5 = 0
    .Npcs(i).y6 = 0
    Next i
    .PanoInf = vbNullString
    .TranInf = 0
    .PanoSup = vbNullString
    .TranSup = 0
    .Fog = 0
    .FogAlpha = 0
    .guildSoloView = 0
    .petView = 0
    .traversable = 0
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
End With
End Sub

Sub ClearQuete(ByVal Index As Long)
Dim i As Long
With quete(Index)
    .nom = vbNullString
    .data1 = 0
    .data2 = 0
    .data2 = 0
    .description = vbNullString
    .reponse = vbNullString
    .String1 = vbNullString
    .temps = 0
    .type = 0
    
    For i = 1 To 15
        .indexe(i).data1 = 1
        .indexe(i).data2 = 0
        .indexe(i).data3 = 0
        .indexe(i).String1 = vbNullString
    Next i
    
    .Recompence.Exp = 0
    .Recompence.objn1 = 1
    .Recompence.objn2 = 1
    .Recompence.objn3 = 1
    .Recompence.objq1 = 0
    .Recompence.objq2 = 0
    .Recompence.objq3 = 0
    .Case = 0
End With
End Sub

Sub ClearPlayerQuete(ByVal Index As Long)
Dim i As Long
With Player(Index).Char(Player(Index).CharNum)
    .QueteEnCour = 0
    .Quetep.data1 = 0
    .Quetep.data2 = 0
    .Quetep.data3 = 0
    .Quetep.String1 = vbNullString
            
    For i = 1 To 15
        .Quetep.indexe(i).data1 = 0
        .Quetep.indexe(i).data2 = 0
        .Quetep.indexe(i).data3 = 0
        .Quetep.indexe(i).String1 = 0
    Next i
End With
End Sub

Sub ClearMaps()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next i
End Sub

Sub ClearQuetes()
Dim i As Long

    For i = 1 To MAX_QUETES
        Call ClearQuete(i)
    Next i
End Sub

Sub ClearShop(ByVal Index As Long)
Dim i As Long
Dim z As Long

    Shop(Index).Name = vbNullString
    Shop(Index).JoinSay = vbNullString
    Shop(Index).LeaveSay = vbNullString
    Shop(Index).FixesItems = 0
    Shop(Index).FixObjet = -1
    
    For z = 1 To 6
        For i = 1 To MAX_TRADES
            Shop(Index).TradeItem(z).value(i).GiveItem = 0
            Shop(Index).TradeItem(z).value(i).GiveValue = 0
            Shop(Index).TradeItem(z).value(i).GetItem = 0
            Shop(Index).TradeItem(z).value(i).GetValue = 0
        Next i
    Next z
End Sub

Sub ClearShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next i
End Sub

Sub ClearSpell(ByVal Index As Long)
With Spell(Index)
    .Name = vbNullString
    .ClassReq = 0
    .LevelReq = 0
    .type = 0
    .data1 = 0
    .data2 = 0
    .data3 = 0
    .MPCost = 0
    .Sound = 0
    .Range = 0
    
    .Big = 0
    
    .SpellAnim = 0
    .SpellTime = 40
    .SpellDone = 1
    
    .SpellIco = 0
    
    .AE = 0
End With
End Sub

Sub ClearSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next i
End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////

Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function

Sub SetPlayerLogin(ByVal Index As Long, ByVal Login As String)
    Player(Index).Login = Login
End Sub

Function GetPlayerPassword(ByVal Index As Long) As String
    GetPlayerPassword = Trim$(Player(Index).Password)
End Function

Sub SetPlayerPassword(ByVal Index As Long, ByVal Password As String)
    Player(Index).Password = Password
End Sub

Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).Char(Player(Index).CharNum).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Char(Player(Index).CharNum).Name = Name
End Sub

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).Char(Player(Index).CharNum).Guild)
End Function

Sub SetPlayerGuild(ByVal Index As Long, ByVal Guild As String)
    Player(Index).Char(Player(Index).CharNum).Guild = Guild
End Sub

Function GetPlayerGuildAccess(ByVal Index As Long) As Long
    GetPlayerGuildAccess = Player(Index).Char(Player(Index).CharNum).Guildaccess
End Function

Sub SetPlayerGuildAccess(ByVal Index As Long, ByVal Guildaccess As Long)
    Player(Index).Char(Player(Index).CharNum).Guildaccess = Guildaccess
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Char(Player(Index).CharNum).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Char(Player(Index).CharNum).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Char(Player(Index).CharNum).sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal sprite As Long)
    Player(Index).Char(Player(Index).CharNum).sprite = sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Char(Player(Index).CharNum).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    If GetPlayerLevel(Index) > MAX_LEVEL Then Exit Sub
    Player(Index).Char(Player(Index).CharNum).Level = Level
End Sub

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    If GetPlayerLevel(Index) > MAX_LEVEL Then Exit Function
    GetPlayerNextLevel = experience(Val(GetPlayerLevel(Index)))
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Char(Player(Index).CharNum).Exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
Dim Queten As Long
Queten = Val(Player(Index).Char(Player(Index).CharNum).QueteEnCour)
    If Queten > 0 Then If quete(Queten).type = QUETE_TYPE_GAGNE_XP Then Call PlayerQueteTypeXp(Index, Queten, Exp)
    Player(Index).Char(Player(Index).CharNum).Exp = Exp
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Char(Player(Index).CharNum).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Char(Player(Index).CharNum).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).Char(Player(Index).CharNum).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).Char(Player(Index).CharNum).PK = PK
End Sub

Function GetPlayerHP(ByVal Index As Long) As Long
    GetPlayerHP = Player(Index).Char(Player(Index).CharNum).HP
End Function

Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)
    Player(Index).Char(Player(Index).CharNum).HP = HP
    
    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then Player(Index).Char(Player(Index).CharNum).HP = GetPlayerMaxHP(Index)
    If GetPlayerHP(Index) < 0 Then Player(Index).Char(Player(Index).CharNum).HP = 0
    Call SendStats(Index)
End Sub

Function GetPlayerMP(ByVal Index As Long) As Long
    GetPlayerMP = Player(Index).Char(Player(Index).CharNum).MP
End Function

Sub SetPlayerMP(ByVal Index As Long, ByVal MP As Long)
    Player(Index).Char(Player(Index).CharNum).MP = MP

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then Player(Index).Char(Player(Index).CharNum).MP = GetPlayerMaxMP(Index)
    If GetPlayerMP(Index) < 0 Then Player(Index).Char(Player(Index).CharNum).MP = 0
End Sub

Function GetPlayerSP(ByVal Index As Long) As Long
    GetPlayerSP = Player(Index).Char(Player(Index).CharNum).SP
End Function

Sub SetPlayerSP(ByVal Index As Long, ByVal SP As Long)
    Player(Index).Char(Player(Index).CharNum).SP = SP

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then Player(Index).Char(Player(Index).CharNum).SP = GetPlayerMaxSP(Index)
    If GetPlayerSP(Index) < 0 Then Player(Index).Char(Player(Index).CharNum).SP = 0
End Sub

Function GetPlayerMaxHP(ByVal Index As Long) As Long
Dim CharNum As Long
Dim i As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then add = item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddHP
    If GetPlayerArmorSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddHP
    If GetPlayerShieldSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddHP
    If GetPlayerHelmetSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddHP
    
    CharNum = Player(Index).CharNum
    'GetPlayerMaxHP = ((Player(index).Char(CharNum).Level + Int(GetPlayerstr(index) / 2) + ClassE(Player(index).Char(CharNum).Class).STR) * 2) + add
    GetPlayerMaxHP = (GetPlayerLevel(Index) * AddHP.Level) + (GetPlayerStr(Index) * AddHP.STR) + (GetPlayerDEF(Index) * AddHP.def) + (GetPlayerMAGI(Index) * AddHP.magi) + (GetPlayerSPEED(Index) * AddHP.Speed) + add
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
Dim CharNum As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then add = item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddMP
    If GetPlayerArmorSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddMP
    If GetPlayerShieldSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddMP
    If GetPlayerHelmetSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddMP
    
    CharNum = Player(Index).CharNum
    'GetPlayerMaxMP = ((Player(index).Char(CharNum).Level + Int(GetPlayerMAGI(index) / 2) + Class(Player(index).Char(CharNum).Class).MAGI) * 2) + add
    GetPlayerMaxMP = (GetPlayerLevel(Index) * AddMP.Level) + (GetPlayerStr(Index) * AddMP.STR) + (GetPlayerDEF(Index) * AddMP.def) + (GetPlayerMAGI(Index) * AddMP.magi) + (GetPlayerSPEED(Index) * AddMP.Speed) + add
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
Dim CharNum As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then add = item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddSP
    If GetPlayerArmorSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddSP
    If GetPlayerShieldSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddSP
    If GetPlayerHelmetSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddSP
    
    CharNum = Player(Index).CharNum
    'GetPlayerMaxSP = ((Player(index).Char(CharNum).Level + Int(GetPlayerSPEED(index) / 2) + Class(Player(index).Char(CharNum).Class).SPEED) * 2) + add
    GetPlayerMaxSP = (GetPlayerLevel(Index) * AddSP.Level) + (GetPlayerStr(Index) * AddSP.STR) + (GetPlayerDEF(Index) * AddSP.def) + (GetPlayerMAGI(Index) * AddSP.magi) + (GetPlayerSPEED(Index) * AddSP.Speed) + add
End Function

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Classe(ClassNum).Name)
End Function

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
    GetClassMaxHP = (1 + Int(Classe(ClassNum).STR / 2) + Classe(ClassNum).STR) * 2
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
    GetClassMaxMP = (1 + Int(Classe(ClassNum).magi / 2) + Classe(ClassNum).magi) * 2
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
    GetClassMaxSP = (1 + Int(Classe(ClassNum).Speed / 2) + Classe(ClassNum).Speed) * 2
End Function

Function GetClassStr(ByVal ClassNum As Long) As Long
    GetClassStr = Classe(ClassNum).STR
End Function

Function GetClassDEF(ByVal ClassNum As Long) As Long
    GetClassDEF = Classe(ClassNum).def
End Function

Function GetClassSPEED(ByVal ClassNum As Long) As Long
    GetClassSPEED = Classe(ClassNum).Speed
End Function

Function GetClassMAGI(ByVal ClassNum As Long) As Long
    GetClassMAGI = Classe(ClassNum).magi
End Function

Function GetPlayerStr(ByVal Index As Long) As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then add = item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddStr
    If GetPlayerArmorSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddStr
    If GetPlayerShieldSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddStr
    If GetPlayerHelmetSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddStr
    
    GetPlayerStr = Player(Index).Char(Player(Index).CharNum).STR + add
End Function

Sub SetPlayerStr(ByVal Index As Long, ByVal STR As Long)
    Player(Index).Char(Player(Index).CharNum).STR = STR
End Sub

Function GetPlayerDEF(ByVal Index As Long) As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then add = item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddDef
    If GetPlayerArmorSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddDef
    If GetPlayerShieldSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddDef
    If GetPlayerHelmetSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddDef
    
    GetPlayerDEF = Player(Index).Char(Player(Index).CharNum).def + add
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal def As Long)
    Player(Index).Char(Player(Index).CharNum).def = def
End Sub

Function GetPlayerSPEED(ByVal Index As Long) As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then add = item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddSpeed
    If GetPlayerArmorSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddSpeed
    If GetPlayerShieldSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddSpeed
    If GetPlayerHelmetSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddSpeed
    
    GetPlayerSPEED = Player(Index).Char(Player(Index).CharNum).Speed + add
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal Speed As Long)
    Player(Index).Char(Player(Index).CharNum).Speed = Speed
End Sub

Function GetPlayerMAGI(ByVal Index As Long) As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then add = item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddMagi
    If GetPlayerArmorSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddMagi
    If GetPlayerShieldSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddMagi
    If GetPlayerHelmetSlot(Index) > 0 Then add = add + item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddMagi
    
    GetPlayerMAGI = Player(Index).Char(Player(Index).CharNum).magi + add
End Function

Sub SetPlayerMAGI(ByVal Index As Long, ByVal magi As Long)
    Player(Index).Char(Player(Index).CharNum).magi = magi
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).Char(Player(Index).CharNum).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).Char(Player(Index).CharNum).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = Player(Index).Char(Player(Index).CharNum).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    If MapNum > 0 And MapNum <= MAX_MAPS Then Player(Index).Char(Player(Index).CharNum).Map = MapNum
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).Char(Player(Index).CharNum).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    Player(Index).Char(Player(Index).CharNum).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Char(Player(Index).CharNum).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Char(Player(Index).CharNum).Y = Y
End Sub

Function GetPlayerSex(ByVal Index As Long) As Byte
    GetPlayerSex = Player(Index).Char(Player(Index).CharNum).Sex
End Function

Sub SetPlayerSex(ByVal Index As Long, ByVal Sex As Byte)
    Player(Index).Char(Player(Index).CharNum).Sex = Sex
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Char(Player(Index).CharNum).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Char(Player(Index).CharNum).Dir = Dir
End Sub

Function GetPlayerIP(ByVal Index As Long) As String
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot)
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
    Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot) = SpellNum
End Sub

Function GetPlayerArmorSlot(ByVal Index As Long) As Long
    GetPlayerArmorSlot = Player(Index).Char(Player(Index).CharNum).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal Index As Long) As Long
    GetPlayerWeaponSlot = Player(Index).Char(Player(Index).CharNum).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal Index As Long) As Long
    GetPlayerHelmetSlot = Player(Index).Char(Player(Index).CharNum).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal Index As Long) As Long
    GetPlayerShieldSlot = Player(Index).Char(Player(Index).CharNum).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).ShieldSlot = InvNum
End Sub

Function GetPlayerPetSlot(ByVal Index As Long) As Long
    GetPlayerPetSlot = Player(Index).Char(Player(Index).CharNum).PetSlot
End Function

Sub SetPlayerPetSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).PetSlot = InvNum
End Sub

Sub BattleMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Long, ByVal Side As Byte)
    Call SendDataTo(Index, "damagedisplay" & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR)
End Sub

Public Sub Attendre(ByVal temps As Long)
Dim lngEndingTime As Long
Dim Seconde As Long
     
     Seconde = temps * 1000
     lngEndingTime = GetTickCount() + (Seconde)
     
     Do While GetTickCount() < lngEndingTime
         DoEvents
     Loop
End Sub

Function Rand(ByVal High As Long, ByVal Low As Long)
Randomize
High = High + 1

Do Until Rand >= Low
    Rand = Int(Rnd * High)
Loop
End Function

Function Anne() As Integer
Anne = Year(Date)
End Function

Function Mois() As Byte
Mois = Month(Date)
End Function

Function JMois() As Byte
JMois = Day(Date)
End Function

Function JSemaine() As Byte
JSemaine = Weekday(Date, vbMonday)
End Function

Function Heure() As Byte
Heure = Hour(Time)
End Function

Function Minutes() As Byte
Minutes = Minute(Time)
End Function

Function Seconde() As Byte
Seconde = Second(Time)
End Function

