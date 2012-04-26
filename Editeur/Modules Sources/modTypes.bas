Attribute VB_Name = "modTypes"
'Option Explicit

' General constants
Public GAME_NAME As String
Public WEBSITE As String
Public MAX_PLAYERS As Long
Public MAX_SPELLS As Long
Public MAX_MAPS As Long
Public MAX_SHOPS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_MAP_ITEMS As Long
Public MAX_EMOTICONS As Long
Public MAX_SPELL_ANIM As Long
Public MAX_BLT_LINE As Long
Public HORS_LIGNE As Byte
Public MAX_LEVEL As Long
Public MAX_QUETES As Long
Public MAX_NPC_SPELLS As Long
Public MAX_DX_SPRITE As Long
Public MAX_DX_PAPERDOLL As Long
Public MAX_DX_SPELLS As Long
Public MAX_DX_BIGSPELLS As Long
Public MAX_DX_PETS As Long
Public MAX_PETS As Long

Public Const MAX_ARROWS As Byte = 100
Public Const MAX_PLAYER_ARROWS As Byte = 100

Public MAX_INV As Integer
Public Const MAX_MAP_NPCS As Byte = 15
Public Const MAX_PLAYER_SPELLS As Byte = 20
Public Const MAX_TRADES As Byte = 66
Public Const MAX_PLAYER_TRADES As Byte = 8
Public Const MAX_NPC_DROPS As Byte = 10
Public Const MAX_DATA_METIER = 100

Public Const NO As Byte = 0
Public Const YES As Byte = 1

Public RecetteSelect As Integer

' Account constants
Public Const NAME_LENGTH As Byte = 20
Public Const MAX_CHARS As Byte = 3

' Basic Security Passwords, You cant connect without it
Public Const SEC_CODE1 As String = "jwehiehfojcvnvnsdinaoiwheoewyriusdyrflsdjncjkxzncisdughfusyfuapsipiuahfpaijnflkjnvjnuahguiryasbdlfkjblsahgfauygewuifaunfauf"
Public Const SEC_CODE2 As String = "ksisyshentwuegeguigdfjkldsnoksamdihuehfidsuhdushdsisjsyayejrioehdoisahdjlasndowijapdnaidhaioshnksfnifohaifhaoinfiwnfinsaihfas"
Public Const SEC_CODE3 As String = "saiugdapuigoihwbdpiaugsdcapvhvinbudhbpidusbnvduisysayaspiufhpijsanfioasnpuvnupashuasohdaiofhaosifnvnuvnuahiosaodiubasdi"
Public Const SEC_CODE4 As String = "88978465734619123425676749756722829121973794379467987945762347631462572792798792492416127957989742945642672"

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Map constants
'Public Const MAX_MAPX = 30
'Public Const MAX_MAPY = 30
Public MAX_MAPX As Long
Public MAX_MAPY As Long
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
Public Const MAP_MORAL_NO_PENALTY As Byte = 2

' Image constants/inconstants
Public Const PIC_X = 32
Public Const PIC_Y = 32
Public PIC_PL As Byte
Public PIC_NPC1 As Byte
Public PIC_NPC2 As Byte

' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEY As Byte = 5
Public Const TILE_TYPE_KEYOPEN As Byte = 6
Public Const TILE_TYPE_HEAL As Byte = 7
Public Const TILE_TYPE_KILL As Byte = 8
Public Const TILE_TYPE_SHOP As Byte = 9
Public Const TILE_TYPE_CBLOCK As Byte = 10
Public Const TILE_TYPE_ARENA As Byte = 11
Public Const TILE_TYPE_SOUND As Byte = 12
Public Const TILE_TYPE_SPRITE_CHANGE As Byte = 13
Public Const TILE_TYPE_SIGN As Byte = 14
Public Const TILE_TYPE_DOOR As Byte = 15
Public Const TILE_TYPE_NOTICE As Byte = 16
'Public Const TILE_TYPE_CHEST As Byte = 17
Public Const TILE_TYPE_CLASS_CHANGE As Byte = 18
Public Const TILE_TYPE_SCRIPTED As Byte = 19
Public Const TILE_TYPE_NPC_SPAWN As Byte = 20
Public Const TILE_TYPE_BANK As Byte = 21
Public Const TILE_TYPE_COFFRE As Byte = 22
Public Const TILE_TYPE_PORTE_CODE As Byte = 23
Public Const TILE_TYPE_BLOCK_MONTURE As Byte = 24
Public Const TILE_TYPE_BLOCK_NIVEAUX As Byte = 25
Public Const TILE_TYPE_TOIT As Byte = 26
Public Const TILE_TYPE_BLOCK_GUILDE As Byte = 27
Public Const TILE_TYPE_BLOCK_TOIT As Byte = 28
Public Const TILE_TYPE_BLOCK_DIR As Byte = 29
Public Const TILE_TYPE_CRAFT As Byte = 30
Public Const TILE_TYPE_METIER As Byte = 31

' quetes constant
Public Const QUETE_TYPE_AUCUN As Byte = 0
Public Const QUETE_TYPE_RECUP As Byte = 1
Public Const QUETE_TYPE_APORT As Byte = 2
Public Const QUETE_TYPE_PARLER As Byte = 3
Public Const QUETE_TYPE_TUER As Byte = 4
Public Const QUETE_TYPE_FINIR As Byte = 5
Public Const QUETE_TYPE_GAGNE_XP As Byte = 6
Public Const QUETE_TYPE_SCRIPT As Byte = 7
Public Const QUETE_TYPE_MINIQUETE As Byte = 8

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
Public Const DIR_UP As Byte = 3
Public Const DIR_DOWN As Byte = 0
Public Const DIR_LEFT As Byte = 1
Public Const DIR_RIGHT As Byte = 2

' Constants for player movement
Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2

' Weather constants
Public Const WEATHER_NONE As Byte = 0
Public Const WEATHER_RAINING As Byte = 1
Public Const WEATHER_SNOWING As Byte = 2
Public Const WEATHER_THUNDER As Byte = 3

' Time constants
Public Const TIME_DAY As Byte = 0
Public Const TIME_NIGHT As Byte = 1

' Admin constants
Public Const ADMIN_MONITER As Byte = 1
Public Const ADMIN_MAPPER As Byte = 2
Public Const ADMIN_DEVELOPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4

' NPC constants
Public Const NPC_BEHAVIOR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOR_FRIENDLY As Byte = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER As Byte = 3
Public Const NPC_BEHAVIOR_GUARD As Byte = 4
Public Const NPC_BEHAVIOR_QUETEUR As Byte = 5
Public Const NPC_BEHAVIOR_SCRIPT As Byte = 6


' Speach bubble constants
Public Const DISPLAY_BUBBLE_TIME As Integer = 2000 ' In milliseconds.
Public DISPLAY_BUBBLE_WIDTH As Byte
Public Const MAX_BUBBLE_WIDTH As Byte = 6 ' In tiles. Includes corners.
Public Const MAX_LINE_LENGTH As Byte = 23 ' In characters.
Public Const MAX_LINES As Byte = 3

' Spell constants
Public Const SPELL_TYPE_ADDHP As Byte = 0
Public Const SPELL_TYPE_ADDMP As Byte = 1
Public Const SPELL_TYPE_ADDSP As Byte = 2
Public Const SPELL_TYPE_SUBHP As Byte = 3
Public Const SPELL_TYPE_SUBMP As Byte = 4
Public Const SPELL_TYPE_SUBSP As Byte = 5
Public Const SPELL_TYPE_SCRIPT As Byte = 6
Public Const SPELL_TYPE_AMELIO As Byte = 7
Public Const SPELL_TYPE_DECONC As Byte = 8
Public Const SPELL_TYPE_PARALY As Byte = 9
Public Const SPELL_TYPE_DEFENC As Byte = 10

Type IndRec
    Data1 As Long
    Data2 As Long
    Data3 As Long
    String1 As String
End Type

Type ChatBubble
    Text As String
    Created As Long
End Type

Type PlayerInvRec
    num As Long
    value As Long
    dur As Long
End Type

Type CoffreTempRec 'coffre
    Numeros As Long
    Valeur As Long
    Durabiliter As Long
End Type

Type SpellAnimRec
    CastedSpell As Byte
    
    SpellTime As Long
    SpellVar As Long
    SpellDone As Long
    
    Target As Long
    TargetType As Long
End Type

Type PlayerArrowRec
    Arrow As Byte
    ArrowNum As Long
    ArrowAnim As Long
    ArrowTime As Long
    ArrowVarX As Long
    ArrowVarY As Long
    ArrowX As Long
    ArrowY As Long
    ArrowPosition As Byte
End Type

Type PlayerQueteRec
    Temps As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    String1 As String
    indexe(1 To 15) As IndRec
End Type

Type PetPosRec
    x As Integer
    y As Integer
    Dir As Byte
    XOffset As Integer
    YOffset As Integer
    Anim As Byte
End Type

Type PlayerRec

    ' General
    name As String * NAME_LENGTH
    guild As String
    Guildaccess As Byte
    Class As Long
    sprite As Long
    Level As Long
    exp As Long
    Access As Byte
    PK As Byte
    
    ' Vitals
    HP As Long
    MP As Long
    SP As Long
    
    ' Stats
    STR As Long
    def As Long
    speed As Long
    magi As Long
    POINTS As Long
    
    ' Worn equipment
    ArmorSlot As Long
    WeaponSlot As Long
    HelmetSlot As Long
    ShieldSlot As Long
    PetSlot As Long
    
    ' Inventory
    inv() As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    pet As PetPosRec
    
    ' Position
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
    
    ' Client use only
    MaxHp As Long
    MaxMP As Long
    MaxSP As Long
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    CastedSpell As Byte
    
    SpellNum As Long
    SpellAnim() As SpellAnimRec
    BloodAnim As SpellAnimRec

    EmoticonNum As Long
    EmoticonTime As Long
    EmoticonVar As Long
    
    LevelUp As Long
    LevelUpT As Long

    Arrow(1 To MAX_PLAYER_ARROWS) As PlayerArrowRec
    
    QueteEnCour As Long
    Quetep As PlayerQueteRec
    
    Anim As Byte
    
   'Paperdoll
   Casque As Long
   armure As Long
   arme As Long
   bouclier As Long
   'Fin paperdoll
   
   Metier As Long
   MetierLvl As Long
   MetierExp As Long
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
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
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

Type NpcMapRec
    x As Byte
    y As Byte
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
    name As String * 40
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
    tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Long
    Npcs(1 To MAX_MAP_NPCS) As NpcMapRec
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
    exp As Long
    objn1 As Long
    objn2 As Long
    objn3 As Long
    objq1 As Long
    objq2 As Long
    objq3 As Long
End Type

Type QueteRec
    nom As String * 40
    Type As Long
    description As String
    reponse As String
    Temps As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    String1 As String
    Recompence As RecompRec
    indexe(1 To 15) As IndRec
    Case As Long
End Type

Type ClassRec
    name As String * NAME_LENGTH
    MaleSprite As Long
    FemaleSprite As Long
    
    Locked As Long
    
    STR As Long
    def As Long
    speed As Long
    magi As Long
    
    ' For client use
    HP As Long
    MP As Long
    SP As Long
End Type

Type ItemRec
    name As String * NAME_LENGTH
    desc As String * 150
    
    Pic As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
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
    dur As Long
    
    x As Byte
    y As Byte
End Type

Type NPCEditorRec
    ItemNum As Long
    ItemValue As Long
    Chance As Long
End Type

Type NpcRec
    name As String * NAME_LENGTH
    AttackSay As String
    
    sprite As Long
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
        
    STR  As Long
    def As Long
    speed As Long
    magi As Long
    MaxHp As Long
    exp As Long
    SpawnTime As Long
    
    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec
    quetenum As Long
    inv As Long
    vol As Long
    
    Spell() As Integer
End Type

Type MapNpcRec
    num As Long
    
    Target As Long
    
    HP As Long
    MaxHp As Long
    MP As Long
    SP As Long
    
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
    
    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
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
    name As String * NAME_LENGTH
    JoinSay As String * 100
    LeaveSay As String * 100
    FixesItems As Byte
    TradeItem(1 To 6) As TradeItemsRec
    FixObjet As Long
End Type

Type SpellRec
    name As String * NAME_LENGTH
    ClassReq As Long
    LevelReq As Long
    Sound As Long
    MPCost As Long
    Type As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Range As Byte
    
    Big As Byte
    
    SpellAnim As Long
    SpellTime As Long
    SpellDone As Long
    
    SpellIco As Long
    
    AE As Long
End Type

Type TempTileRec
    DoorOpen As Byte
End Type

Type PlayerTradeRec
    InvNum As Long
    InvName As String
    InvVal As Long
End Type
Public Trading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
Public Trading2(1 To MAX_PLAYER_TRADES) As PlayerTradeRec

Type EmoRec
    Pic As Long
    Command As String
End Type

Type DropRainRec
    x As Long
    y As Long
    Randomized As Boolean
    speed As Byte
End Type

Type PetsRec
    nom As String
    sprite As Long
    addForce As Byte
    addDefence As Byte
End Type

'pour patchs
Type Fichiers
    nom As String
    version As String
    Chemins As String
End Type

Type MetierRec
    nom As String
    Type As Byte
    desc As String
    
    Data(0 To MAX_DATA_METIER, 0 To 1) As Integer
End Type

Type RecetteRec
    nom As String
    InCraft(0 To 9, 0 To 1) As Integer
    craft(0 To 1) As Integer
End Type


' Bubble thing
Public Bubble() As ChatBubble

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte
Public quete() As QueteRec
Public Map() As MapRec
Public TempMap(0 To 5) As MapRec
Public TempTile() As TempTileRec
Public Player() As PlayerRec
Public PlayerAnim() As Long
Public Class() As ClassRec
Public Item() As ItemRec
Public Npc() As NpcRec
Public MapItem() As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop() As ShopRec
Public Spell() As SpellRec
Public Emoticons() As EmoRec
Public MapReport() As MapRec
Public Experience() As Long
Public Pets() As PetsRec
Public CoffreTmp(1 To 30) As CoffreTempRec 'coffre
Public Metier() As MetierRec
Public MAX_METIER As Long
Public recette() As RecetteRec
Public MAX_RECETTE As Long
Public MAX_RAINDROPS As Long
Public BLT_RAIN_DROPS As Long
Public DropRain() As DropRainRec

Public BLT_SNOW_DROPS As Long
Public DropSnow() As DropRainRec

Type ItemTradeRec
    ItemGetNum As Long
    ItemGiveNum As Long
    ItemGetVal As Long
    ItemGiveVal As Long
End Type
Type TradeRec
    Items(1 To MAX_TRADES) As ItemTradeRec
    Selected As Long
    SelectedItem As Long
End Type
Public Trade(1 To 6) As TradeRec

Type ArrowRec
    name As String
    Pic As Long
    Range As Byte
End Type
Public Arrows(1 To MAX_ARROWS) As ArrowRec

Type BattleMsgRec
    Msg As String
    Index As Byte
    Color As Long
    Time As Long
    Done As Byte
    y As Long
End Type
Public BattlePMsg() As BattleMsgRec
Public BattleMMsg() As BattleMsgRec

Type ItemDurRec
    Item As Long
    dur As Long
    Done As Byte
End Type
Public ItemDur(1 To 4) As ItemDurRec

Public Inventory As Long

Public Minu As Long
Public Seco As Long

'Type pour stocker le contenu de Account.ini
Type TpAccOpt
    InfName As String
    InfPass As String
    SpeechBubbles As Boolean
    NpcBar As Boolean
    NpcName As Boolean
    NpcDamage As Boolean
    PlayBar As Boolean
    PlayName As Boolean
    PlayDamage As Boolean
    MapGrid As Boolean
    Music As Boolean
    Sound As Boolean
    Autoscroll As Boolean
    NomObjet As Boolean
    CPreVisu As Boolean
    LowEffect As Boolean
End Type
Public AccOpt As TpAccOpt


Sub ClearSpell(ByVal Index As Long)
    Spell(Index).name = vbNullString
    Spell(Index).ClassReq = 0
    Spell(Index).LevelReq = 0
    Spell(Index).Type = 0
    Spell(Index).Data1 = 0
    Spell(Index).Data2 = 0
    Spell(Index).Data3 = 0
    Spell(Index).MPCost = 0
    Spell(Index).Sound = 0
    Spell(Index).Range = 0
    
    Spell(Index).Big = 0
    
    Spell(Index).SpellAnim = 0
    Spell(Index).SpellTime = 40
    Spell(Index).SpellDone = 1
    
    Spell(Index).SpellIco = 0
    
    Spell(Index).AE = 0
End Sub

Sub ClearShop(ByVal Index As Long)
Dim i As Long
Dim z As Long

    Shop(Index).name = vbNullString
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

Sub ClearNpc(ByVal Index As Long)
Dim i As Long
With Npc(Index)
    .name = vbNullString
    .AttackSay = vbNullString
    .sprite = 0
    .SpawnSecs = 0
    .Behavior = 0
    .Range = 0
    .STR = 0
    .def = 0
    .speed = 0
    .magi = 0
    .MaxHp = 0
    .exp = 0
    .SpawnTime = 0
    .quetenum = 0
    .inv = 0
    .vol = 0
    For i = 1 To MAX_NPC_DROPS
        .ItemNPC(i).Chance = 0
        .ItemNPC(i).ItemNum = 0
        .ItemNPC(i).ItemValue = 0
    Next i
End With
End Sub

Sub ClearQuete(ByVal Index As Long)
    quete(Index).nom = vbNullString
    quete(Index).Data1 = 0
    quete(Index).Data2 = 0
    quete(Index).Data2 = 0
    quete(Index).description = vbNullString
    quete(Index).reponse = vbNullString
    quete(Index).String1 = vbNullString
    quete(Index).Temps = 0
    quete(Index).Type = 0
    Dim i As Long
    For i = 1 To 15
        quete(Index).indexe(i).Data1 = 1
        quete(Index).indexe(i).Data2 = 0
        quete(Index).indexe(i).Data3 = 0
        quete(Index).indexe(i).String1 = vbNullString
    Next i
    quete(Index).Recompence.exp = 0
    quete(Index).Recompence.objn1 = 1
    quete(Index).Recompence.objn2 = 1
    quete(Index).Recompence.objn3 = 1
    quete(Index).Recompence.objq1 = 0
    quete(Index).Recompence.objq2 = 0
    quete(Index).Recompence.objq3 = 0
    quete(Index).Case = 0
End Sub

Sub ClearTempTile()
Dim x As Long, y As Long

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            TempTile(x, y).DoorOpen = NO
        Next x
    Next y
End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim i As Long
Dim n As Long

    Player(Index).name = vbNullString
    Player(Index).guild = vbNullString
    Player(Index).Guildaccess = 0
    Player(Index).Class = 0
    Player(Index).Level = 0
    Player(Index).sprite = 0
    Player(Index).exp = 0
    Player(Index).Access = 0
    Player(Index).PK = NO
        
    Player(Index).HP = 0
    Player(Index).MP = 0
    Player(Index).SP = 0
        
    Player(Index).STR = 0
    Player(Index).def = 0
    Player(Index).speed = 0
    Player(Index).magi = 0
        
    For n = 1 To MAX_INV
        Player(Index).inv(n).num = 0
        Player(Index).inv(n).value = 0
        Player(Index).inv(n).dur = 0
    Next n
        
    Player(Index).ArmorSlot = 0
    Player(Index).WeaponSlot = 0
    Player(Index).HelmetSlot = 0
    Player(Index).ShieldSlot = 0
    Player(Index).PetSlot = 0
    
    Player(Index).pet.Dir = 0
    Player(Index).pet.x = 0
    Player(Index).pet.y = 0
    
    Player(Index).Map = 0
    Player(Index).x = 0
    Player(Index).y = 0
    Player(Index).Dir = 0
    
    ' Client use only
    Player(Index).MaxHp = 0
    Player(Index).MaxMP = 0
    Player(Index).MaxSP = 0
    Player(Index).XOffset = 0
    Player(Index).YOffset = 0
    Player(Index).Moving = 0
    Player(Index).Attacking = 0
    Player(Index).AttackTimer = 0
    Player(Index).MapGetTimer = 0
    Player(Index).CastedSpell = NO
    Player(Index).EmoticonNum = -1
    Player(Index).EmoticonTime = 0
    Player(Index).EmoticonVar = 0
    
    For i = 1 To MAX_SPELL_ANIM
        Player(Index).SpellAnim(i).CastedSpell = NO
        Player(Index).SpellAnim(i).SpellTime = 0
        Player(Index).SpellAnim(i).SpellVar = 0
        Player(Index).SpellAnim(i).SpellDone = 0
        
        Player(Index).SpellAnim(i).Target = 0
        Player(Index).SpellAnim(i).TargetType = 0
    Next i
    
    Player(Index).SpellNum = 0
    
    For i = 1 To MAX_BLT_LINE
        BattlePMsg(i).Index = 1
        BattlePMsg(i).Time = i
        BattleMMsg(i).Index = 1
        BattleMMsg(i).Time = i
    Next i
    
    Player(Index).QueteEnCour = 0
    Player(Index).Quetep.Data1 = 0
    Player(Index).Quetep.Data2 = 0
    Player(Index).Quetep.Data3 = 0
    Player(Index).Quetep.String1 = vbNullString
      
    For n = 1 To 15
        Player(Index).Quetep.indexe(n).Data1 = 0
        Player(Index).Quetep.indexe(n).Data2 = 0
        Player(Index).Quetep.indexe(n).Data3 = 0
        Player(Index).Quetep.indexe(n).String1 = vbNullString
    Next n
    
    Inventory = 1
End Sub

Sub ClearPlayerQuete(ByVal Index As Long)
Dim i As Long
        Player(MyIndex).QueteEnCour = 0
        Player(MyIndex).Quetep.Data1 = 0
        Player(MyIndex).Quetep.Data2 = 0
        Player(MyIndex).Quetep.Data3 = 0
        Player(MyIndex).Quetep.String1 = vbNullString
        Accepter = False
        
        For i = 1 To 15
            Player(MyIndex).Quetep.indexe(i).Data1 = 0
            Player(MyIndex).Quetep.indexe(i).Data2 = 0
            Player(MyIndex).Quetep.indexe(i).Data3 = 0
            Player(MyIndex).Quetep.indexe(i).String1 = 0
        Next i
End Sub

Sub ClearPet(ByVal Index As Long)
    Pets(Index).nom = ""
    Pets(Index).sprite = 0
    Pets(Index).addForce = 0
    Pets(Index).addDefence = 0
End Sub

Sub ClearRecette(ByVal Index As Long)
    recette(Index).nom = vbNullString
    For i = 0 To 9
        recette(Index).InCraft(i, 0) = 0
        recette(Index).InCraft(i, 1) = 0
    Next i
    recette(Index).craft(0) = 0
    recette(Index).craft(1) = 0
End Sub

Sub ClearItem(ByVal Index As Long)
    Item(Index).name = vbNullString
    Item(Index).desc = vbNullString
    
    Item(Index).Type = 0
    Item(Index).Data1 = 0
    Item(Index).Data2 = 0
    Item(Index).Data3 = 0
    Item(Index).StrReq = 0
    Item(Index).DefReq = 0
    Item(Index).SpeedReq = 0
    Item(Index).ClassReq = -1
    Item(Index).AccessReq = 0
    
    Item(Index).paperdoll = 0
    Item(Index).paperdollPic = 0
    
    Item(Index).Empilable = 0
    
    Item(Index).AddHP = 0
    Item(Index).AddMP = 0
    Item(Index).AddSP = 0
    Item(Index).AddStr = 0
    Item(Index).AddDef = 0
    Item(Index).AddMagi = 0
    Item(Index).AddSpeed = 0
    Item(Index).AddEXP = 0
    Item(Index).AttackSpeed = 1000
    
    Item(Index).NCoul = 0
    Item(Index).tArme = 0
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
End Sub

Sub ClearMapItem(ByVal Index As Long)
Call ZeroMemory(MapItem(Index), Len(MapItem(Index)))
End Sub
Sub ClearMap(i As Integer)
Dim x As Long
Dim y As Long

    Map(i).name = vbNullString
    Map(i).Revision = 0
    Map(i).Moral = 0
    Map(i).Up = 0
    Map(i).Down = 0
    Map(i).Left = 0
    Map(i).Right = 0
    Map(i).Indoors = 0
        
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            Map(i).tile(x, y).Ground = 0
            Map(i).tile(x, y).Mask = 0
            Map(i).tile(x, y).Anim = 0
            Map(i).tile(x, y).Mask2 = 0
            Map(i).tile(x, y).M2Anim = 0
            Map(i).tile(x, y).Mask3 = 0 '<--
            Map(i).tile(x, y).M3Anim = 0 '<--
            Map(i).tile(x, y).Fringe = 0
            Map(i).tile(x, y).FAnim = 0
            Map(i).tile(x, y).Fringe2 = 0
            Map(i).tile(x, y).F2Anim = 0
            Map(i).tile(x, y).Fringe3 = 0 '<--
            Map(i).tile(x, y).F3Anim = 0 '<--
            Map(i).tile(x, y).Type = 0
            Map(i).tile(x, y).Data1 = 0
            Map(i).tile(x, y).Data2 = 0
            Map(i).tile(x, y).Data3 = 0
            Map(i).tile(x, y).String1 = vbNullString
            Map(i).tile(x, y).String2 = vbNullString
            Map(i).tile(x, y).String3 = vbNullString
            Map(i).tile(x, y).Light = 0
            Map(i).tile(x, y).GroundSet = 0
            Map(i).tile(x, y).MaskSet = 0
            Map(i).tile(x, y).AnimSet = 0
            Map(i).tile(x, y).Mask2Set = 0
            Map(i).tile(x, y).M2AnimSet = 0
            Map(i).tile(x, y).Mask3Set = 0 '<--
            Map(i).tile(x, y).M3AnimSet = 0 '<--
            Map(i).tile(x, y).FringeSet = 0
            Map(i).tile(x, y).FAnimSet = 0
            Map(i).tile(x, y).Fringe2Set = 0
            Map(i).tile(x, y).F2AnimSet = 0
            Map(i).tile(x, y).Fringe3Set = 0 '<--
            Map(i).tile(x, y).F3AnimSet = 0 '<--
        Next x
    Next y
    For x = 1 To MAX_MAP_NPCS
        Map(i).Npc(x) = 0
        Map(i).Npcs(x).Axy = False
        Map(i).Npcs(x).Axy1 = False
        Map(i).Npcs(x).Axy2 = False
        Map(i).Npcs(x).boucle = 0
        Map(i).Npcs(x).Hasardm = 1
        Map(i).Npcs(x).Hasardp = 1
        Map(i).Npcs(x).Imobile = 0
        Map(i).Npcs(x).x = 0
        Map(i).Npcs(x).x1 = 0
        Map(i).Npcs(x).x2 = 0
        Map(i).Npcs(x).x3 = 0
        Map(i).Npcs(x).x4 = 0
        Map(i).Npcs(x).x5 = 0
        Map(i).Npcs(x).x6 = 0
        Map(i).Npcs(x).y = 0
        Map(i).Npcs(x).y2 = 0
        Map(i).Npcs(x).y3 = 0
        Map(i).Npcs(x).y4 = 0
        Map(i).Npcs(x).y5 = 0
        Map(i).Npcs(x).y6 = 0
    Next x
    Map(i).PanoInf = vbNullString
    Map(i).TranInf = 0
    Map(i).PanoSup = vbNullString
    Map(i).TranSup = 0


End Sub

Sub ClearTempMaps()
Dim i As Integer
For i = 0 To 5
    TempMap(i).name = vbNullString
    TempMap(i).Revision = -1
    TempMap(i).Moral = 0
    TempMap(i).Up = 0
    TempMap(i).Down = 0
    TempMap(i).Left = 0
    TempMap(i).Right = 0
    TempMap(i).Indoors = 0
        
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            TempMap(i).tile(x, y).Ground = 0
            TempMap(i).tile(x, y).Mask = 0
            TempMap(i).tile(x, y).Anim = 0
            TempMap(i).tile(x, y).Mask2 = 0
            TempMap(i).tile(x, y).M2Anim = 0
            TempMap(i).tile(x, y).Mask3 = 0 '<--
            TempMap(i).tile(x, y).M3Anim = 0 '<--
            TempMap(i).tile(x, y).Fringe = 0
            TempMap(i).tile(x, y).FAnim = 0
            TempMap(i).tile(x, y).Fringe2 = 0
            TempMap(i).tile(x, y).F2Anim = 0
            TempMap(i).tile(x, y).Fringe3 = 0 '<--
            TempMap(i).tile(x, y).F3Anim = 0 '<--
            TempMap(i).tile(x, y).Type = 0
            TempMap(i).tile(x, y).Data1 = 0
            TempMap(i).tile(x, y).Data2 = 0
            TempMap(i).tile(x, y).Data3 = 0
            TempMap(i).tile(x, y).String1 = vbNullString
            TempMap(i).tile(x, y).String2 = vbNullString
            TempMap(i).tile(x, y).String3 = vbNullString
            TempMap(i).tile(x, y).Light = 0
            TempMap(i).tile(x, y).GroundSet = 0
            TempMap(i).tile(x, y).MaskSet = 0
            TempMap(i).tile(x, y).AnimSet = 0
            TempMap(i).tile(x, y).Mask2Set = 0
            TempMap(i).tile(x, y).M2AnimSet = 0
            TempMap(i).tile(x, y).Mask3Set = 0 '<--
            TempMap(i).tile(x, y).M3AnimSet = 0 '<--
            TempMap(i).tile(x, y).FringeSet = 0
            TempMap(i).tile(x, y).FAnimSet = 0
            TempMap(i).tile(x, y).Fringe2Set = 0
            TempMap(i).tile(x, y).F2AnimSet = 0
            TempMap(i).tile(x, y).Fringe3Set = 0 '<--
            TempMap(i).tile(x, y).F3AnimSet = 0 '<--
        Next x
    Next y
    For x = 1 To MAX_MAP_NPCS
        TempMap(i).Npc(x) = 0
        TempMap(i).Npcs(x).Axy = False
        TempMap(i).Npcs(x).Axy1 = False
        TempMap(i).Npcs(x).Axy2 = False
        TempMap(i).Npcs(x).boucle = 0
        TempMap(i).Npcs(x).Hasardm = 1
        TempMap(i).Npcs(x).Hasardp = 1
        TempMap(i).Npcs(x).Imobile = 0
        TempMap(i).Npcs(x).x = 0
        TempMap(i).Npcs(x).x1 = 0
        TempMap(i).Npcs(x).x2 = 0
        TempMap(i).Npcs(x).x3 = 0
        TempMap(i).Npcs(x).x4 = 0
        TempMap(i).Npcs(x).x5 = 0
        TempMap(i).Npcs(x).x6 = 0
        TempMap(i).Npcs(x).y = 0
        TempMap(i).Npcs(x).y2 = 0
        TempMap(i).Npcs(x).y3 = 0
        TempMap(i).Npcs(x).y4 = 0
        TempMap(i).Npcs(x).y5 = 0
        TempMap(i).Npcs(x).y6 = 0
    Next x
    TempMap(i).PanoInf = vbNullString
    TempMap(i).TranInf = 0
    TempMap(i).PanoSup = vbNullString
    TempMap(i).TranSup = 0
    TempMap(i).Fog = 0
    TempMap(i).FogAlpha = 0
    TempMap(i).guildSoloView = 0
    TempMap(i).petView = 0
    TempMap(i).traversable = 0
Next
End Sub

Sub NetQueteType(ByVal Index As Integer)
Dim i As Long
    Call ZeroMemory(quete(i), Len(quete(i)))
    For i = 1 To 15
        quete(Index).indexe(i).Data1 = 1
    Next i
End Sub

Sub NetTempMap(ByVal Index As Byte)
Dim x As Long
   Call ZeroMemory(TempMap(Index), Len(TempMap(Index)))
    TempMap(Index).Revision = -1
    For x = 1 To MAX_MAP_NPCS
        TempMap(Index).Npcs(x).Hasardm = 1
        TempMap(Index).Npcs(x).Hasardp = 1
    Next x
End Sub

Sub VidercttMap(ByVal MapNum As Long)

Call ZeroMemory(Map(MapNum), Len(Map(MapNum)))
Call ClearMapItems
Call ClearMapNpcs
End Sub

Sub ViderTMap(ByVal MapNum As Long)
Call ZeroMemory(Map(MapNum), Len(Map(MapNum)))
End Sub

Sub ClearMapItems()
Dim x As Long

    For x = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(x)
    Next x
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    Call ZeroMemory(MapNpc(Index), Len(MapNpc(Index)))
    PNJAnim(Index) = 1
End Sub

Sub ClearMapNpcs()
Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next i
End Sub

Function GetPlayerName(ByVal Index As Long) As String
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal name As String)
    Player(Index).name = name
End Sub

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).guild)
End Function

Sub SetPlayerGuild(ByVal Index As Long, ByVal guild As String)
    Player(Index).guild = guild
End Sub

Function GetPlayerGuildAccess(ByVal Index As Long) As Long
    GetPlayerGuildAccess = Player(Index).Guildaccess
End Function

Sub SetPlayerGuildAccess(ByVal Index As Long, ByVal Guildaccess As Long)
    Player(Index).Guildaccess = Guildaccess
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal sprite As Long)
    Player(Index).sprite = sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Level = Level
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal exp As Long)
    Player(Index).exp = exp
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).PK = PK
End Sub

Function GetPlayerHP(ByVal Index As Long) As Long
    GetPlayerHP = Player(Index).HP
End Function

Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)
    Player(Index).HP = HP
    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then Player(Index).HP = GetPlayerMaxHP(Index)
End Sub

Function GetPlayerMP(ByVal Index As Long) As Long
    GetPlayerMP = Player(Index).MP
End Function

Sub SetPlayerMP(ByVal Index As Long, ByVal MP As Long)
    Player(Index).MP = MP

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then Player(Index).MP = GetPlayerMaxMP(Index)
End Sub

Function GetPlayerSP(ByVal Index As Long) As Long
    GetPlayerSP = Player(Index).SP
End Function

Sub SetPlayerSP(ByVal Index As Long, ByVal SP As Long)
    Player(Index).SP = SP

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then Player(Index).SP = GetPlayerMaxSP(Index)
End Sub

Function GetPlayerMaxHP(ByVal Index As Long) As Long
    GetPlayerMaxHP = Player(Index).MaxHp
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
    GetPlayerMaxMP = Player(Index).MaxMP
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
    GetPlayerMaxSP = Player(Index).MaxSP
End Function

Function Getplayerstr(ByVal Index As Long) As Long
    Getplayerstr = Player(Index).STR
End Function

Sub Setplayerstr(ByVal Index As Long, ByVal STR As Long)
    Player(Index).STR = STR
End Sub

Function GetPlayerDEF(ByVal Index As Long) As Long
    GetPlayerDEF = Player(Index).def
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal def As Long)
    Player(Index).def = def
End Sub

Function GetPlayerSPEED(ByVal Index As Long) As Long
    GetPlayerSPEED = Player(Index).speed
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal speed As Long)
    Player(Index).speed = speed
End Sub

Function GetPlayerMAGI(ByVal Index As Long) As Long
    GetPlayerMAGI = Player(Index).magi
End Function

Sub SetPlayerMAGI(ByVal Index As Long, ByVal magi As Long)
    Player(Index).magi = magi
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).POINTS = POINTS
End Sub

'Function GetPlayerMap(ByVal Index As Long) As Long
'If Index <= 0 Then Exit Function
'    GetPlayerMap = Player(Index).Map
'End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    Player(Index).Map = MapNum
    WriteINI "CONFIG", "ERR", Val(MapNum), App.Path & "\Config.ini"
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).x
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)
    Player(Index).x = x
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
    Player(Index).y = y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Dir = Dir
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).inv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).inv(InvSlot).num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).inv(InvSlot).value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).inv(InvSlot).value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).inv(InvSlot).dur
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).inv(InvSlot).dur = ItemDur
End Sub

Function GetPlayerArmorSlot(ByVal Index As Long) As Long
    GetPlayerArmorSlot = Player(Index).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal Index As Long) As Long
    GetPlayerWeaponSlot = Player(Index).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal Index As Long) As Long
    GetPlayerHelmetSlot = Player(Index).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal Index As Long) As Long
    GetPlayerShieldSlot = Player(Index).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).ShieldSlot = InvNum
End Sub

Function GetPlayerPetSlot(ByVal Index As Long) As Long
    GetPlayerPetSlot = Player(Index).PetSlot
End Function

Sub SetPlayerPetSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).PetSlot = InvNum
End Sub


