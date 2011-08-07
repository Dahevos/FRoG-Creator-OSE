Attribute VB_Name = "modTypes"
Option Explicit

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
Public MAX_LEVEL As Long
Public MAX_QUETES As Long
Public MAX_DX_PETS As Long
Public MAX_PETS As Long
Public Const MAX_ARROWS As Byte = 100
Public Const MAX_PLAYER_ARROWS As Byte = 100

Public MAX_INV As Integer
Public Const MAX_PARTY_MEMBERS As Byte = 20
Public Const MAX_MAP_NPCS As Byte = 15
Public Const MAX_PLAYER_SPELLS As Byte = 20
Public Const MAX_TRADES As Byte = 66
Public Const MAX_PLAYER_TRADES As Byte = 8
Public Const MAX_NPC_DROPS As Byte = 10

Public Const NO As Byte = 0
Public Const YES As Byte = 1

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

' Image constants
Public Const PIC_X As Integer = 32
Public Const PIC_Y As Integer = 32
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
Public Const TILE_TYPE_CHEST As Byte = 17
Public Const TILE_TYPE_CLASS_CHANGE As Byte = 18
Public Const TILE_TYPE_SCRIPTED As Byte = 19
Public Const TILE_TYPE_NPC_SPAWN As Byte = 20
Public Const TILE_TYPE_BANK As Byte = 21
Public Const TILE_TYPE_POISON As Byte = 26
Public Const TILE_TYPE_COFFRE As Byte = 22
Public Const TILE_TYPE_PORTE_CODE As Byte = 23
Public Const TILE_TYPE_BLOCK_MONTURE As Byte = 24
Public Const TILE_TYPE_BLOCK_NIVEAUX As Byte = 25
Public Const TILE_TYPE_TOIT As Byte = 26
Public Const TILE_TYPE_BLOCK_GUILDE As Byte = 27
Public Const TILE_TYPE_BLOCK_TOIT As Byte = 28
Public Const TILE_TYPE_BLOCK_DIR As Byte = 29

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

' Direction constants
Public Const DIR_UP As Byte = 3
Public Const DIR_DOWN As Byte = 0
Public Const DIR_LEFT As Byte = 1
Public Const DIR_RIGHT As Byte = 2

' Constants for player movement
Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2
Public Const MOVING_VEHICUL As Byte = 3

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

' Speach bubble constants
Public Const DISPLAY_BUBBLE_TIME As Integer = 4000 ' In milliseconds.
Public DISPLAY_BUBBLE_WIDTH As Byte
Public Const MAX_BUBBLE_WIDTH As Byte = 16 ' In tiles. Includes corners.
Public Const MAX_LINE_LENGTH As Byte = 20 ' In characters.
Public Const MAX_LINES As Byte = 5

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

Public Loading As Boolean
Public deco As Boolean

Type ChatBubble
    Text As String
    Created As Long
End Type

Type PlayerInvRec
    num As Long
    Value As Long
    dur As Long
End Type

Type CoffreTempRec
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

Type IndRec
    Data1 As Long
    Data2 As Long
    Data3 As Long
    String1 As String
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
    Guild As String
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
    DEF As Long
    speed As Long
    MAGI As Long
    POINTS As Long
    
    ' Worn equipment
    ArmorSlot As Long
    WeaponSlot As Long
    HelmetSlot As Long
    ShieldSlot As Long
    PetSlot As Long
    
    ' Inventory
    Inv() As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    pet As PetPosRec
    
    ' Position
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte

    ' Client use only
    MaxHp As Long
    MaxMp As Long
    MaxSP As Long
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    CastedSpell As Byte
    PartyIndex As Byte
    
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
    'PAPERDOLL
    Casque As Long
    Armure As Long
    Arme As Long
    Bouclier As Long
    'FIN PAPERDOLL
End Type
    
Type TileRec
    Ground As Long
    Mask As Long
    Anim As Long
    Mask2 As Long
    M2Anim As Long
    Mask3 As Long
    M3Anim As Long
    Fringe As Long
    FAnim As Long
    Fringe2 As Long
    F2Anim As Long
    Fringe3 As Long
    F3Anim As Long
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
    Mask3Set As Byte
    M3AnimSet As Byte
    FringeSet As Byte
    FAnimSet As Byte
    Fringe2Set As Byte
    F2AnimSet As Byte
    Fringe3Set As Byte
    F3AnimSet As Byte
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
    Tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Long
    Npcs(1 To MAX_MAP_NPCS) As NpcMapRec
    PanoInf As String * 50
    TranInf As Byte
    PanoSup As String * 50
    TranSup As Byte
    Fog As Integer
    FogAlpha As Byte
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
    DEF As Long
    speed As Long
    MAGI As Long
    
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
End Type

Type MapItemRec
    num As Long
    Value As Long
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
    DEF As Long
    speed As Long
    MAGI As Long
    MaxHp As Long
    exp As Long
    SpawnTime As Long
    
    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec
    QueteNum As Long
    Inv As Long
    Vol As Long
End Type

Type MapNpcRec
    num As Long
    
    Target As Long
    
    HP As Long
    MaxHp As Long
    MP As Long
    MaxMp As Long
    SP As Long
    
    Map As Long
    x As Byte
    y As Byte
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
    Value(1 To MAX_TRADES) As TradeItemRec
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

' Bubble thing
Public Bubble() As ChatBubble

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte
Public quete() As QueteRec
Public Map() As MapRec
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
Public CoffreTmp(1 To 30) As CoffreTempRec
Public Pets() As PetsRec

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
    LowEffect As Boolean
End Type

Public rac(0 To 13, 0 To 1) As String
Public dragAndDrop As Byte
Public dragAndDropT As Byte

Public AccOpt As TpAccOpt

' Configuration Menu Option des touches
Type optToucheRec
    nom As String
    Value As Byte
End Type
Public nelvl As Long
Public Const TCHMAX = 51
Public optTouche(0 To TCHMAX) As optToucheRec

Sub iniOptTouche()
    optTouche(0).nom = "A"
    optTouche(0).Value = vbKeyA
    optTouche(1).nom = "B"
    optTouche(1).Value = vbKeyB
    optTouche(2).nom = "C"
    optTouche(2).Value = vbKeyC
    optTouche(3).nom = "D"
    optTouche(3).Value = vbKeyD
    optTouche(4).nom = "E"
    optTouche(4).Value = vbKeyE
    optTouche(5).nom = "F"
    optTouche(5).Value = vbKeyF
    optTouche(6).nom = "G"
    optTouche(6).Value = vbKeyG
    optTouche(7).nom = "H"
    optTouche(7).Value = vbKeyH
    optTouche(8).nom = "I"
    optTouche(8).Value = vbKeyI
    optTouche(9).nom = "J"
    optTouche(9).Value = vbKeyJ
    optTouche(10).nom = "K"
    optTouche(10).Value = vbKeyK
    optTouche(11).nom = "L"
    optTouche(11).Value = vbKeyL
    optTouche(12).nom = "M"
    optTouche(12).Value = vbKeyM
    optTouche(13).nom = "N"
    optTouche(13).Value = vbKeyN
    optTouche(14).nom = "O"
    optTouche(14).Value = vbKeyO
    optTouche(15).nom = "P"
    optTouche(15).Value = vbKeyP
    optTouche(16).nom = "Q"
    optTouche(16).Value = vbKeyQ
    optTouche(17).nom = "R"
    optTouche(17).Value = vbKeyR
    optTouche(18).nom = "S"
    optTouche(18).Value = vbKeyS
    optTouche(19).nom = "T"
    optTouche(19).Value = vbKeyT
    optTouche(20).nom = "U"
    optTouche(20).Value = vbKeyU
    optTouche(21).nom = "V"
    optTouche(21).Value = vbKeyV
    optTouche(22).nom = "W"
    optTouche(22).Value = vbKeyW
    optTouche(23).nom = "X"
    optTouche(23).Value = vbKeyX
    optTouche(24).nom = "Y"
    optTouche(24).Value = vbKeyY
    optTouche(25).nom = "Z"
    optTouche(25).Value = vbKeyZ
    optTouche(26).nom = "0"
    optTouche(26).Value = vbKey0
    optTouche(27).nom = "1"
    optTouche(27).Value = vbKey1
    optTouche(28).nom = "2"
    optTouche(28).Value = vbKey2
    optTouche(29).nom = "3"
    optTouche(29).Value = vbKey3
    optTouche(30).nom = "4"
    optTouche(30).Value = vbKey4
    optTouche(31).nom = "5"
    optTouche(31).Value = vbKey5
    optTouche(32).nom = "6"
    optTouche(32).Value = vbKey6
    optTouche(33).nom = "7"
    optTouche(33).Value = vbKey7
    optTouche(34).nom = "8"
    optTouche(34).Value = vbKey8
    optTouche(35).nom = "9"
    optTouche(35).Value = vbKey9
    optTouche(36).nom = "F1"
    optTouche(36).Value = vbKeyF1
    optTouche(37).nom = "F2"
    optTouche(37).Value = vbKeyF2
    optTouche(38).nom = "F3"
    optTouche(38).Value = vbKeyF3
    optTouche(39).nom = "F4"
    optTouche(39).Value = vbKeyF4
    optTouche(40).nom = "F5"
    optTouche(40).Value = vbKeyF5
    optTouche(41).nom = "F6"
    optTouche(41).Value = vbKeyF6
    optTouche(42).nom = "F7"
    optTouche(42).Value = vbKeyF7
    optTouche(43).nom = "F8"
    optTouche(43).Value = vbKeyF8
    optTouche(44).nom = "Haut"
    optTouche(44).Value = vbKeyUp
    optTouche(45).nom = "Bas"
    optTouche(45).Value = vbKeyDown
    optTouche(46).nom = "Gauche"
    optTouche(46).Value = vbKeyLeft
    optTouche(47).nom = "Droite"
    optTouche(47).Value = vbKeyRight
    optTouche(48).nom = "Ctrl"
    optTouche(48).Value = vbKeyControl
    optTouche(49).nom = "Alt"
    optTouche(49).Value = vbKeyMenu
    optTouche(50).nom = "Shift"
    optTouche(50).Value = vbKeyShift
    optTouche(51).nom = "Espace"
    optTouche(51).Value = vbKeySpace
    
    
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

With Player(Index)
    .name = vbNullString
    .Guild = vbNullString
    .Guildaccess = 0
    .Class = 0
    .Level = 0
    .sprite = 0
    .exp = 0
    .Access = 0
    .PK = NO
        
    .HP = 0
    .MP = 0
    .SP = 0
        
    .STR = 0
    .DEF = 0
    .speed = 0
    .MAGI = 0
    
    .QueteEnCour = 0
    .Quetep.Data1 = 0
    .Quetep.Data2 = 0
    .Quetep.Data3 = 0
    .Quetep.String1 = vbNullString
      
    For n = 1 To 15
    .Quetep.indexe(n).Data1 = 0
    .Quetep.indexe(n).Data2 = 0
    .Quetep.indexe(n).Data3 = 0
    .Quetep.indexe(n).String1 = vbNullString
    Next n
        
    For n = 1 To MAX_INV
        .Inv(n).num = 0
        .Inv(n).Value = 0
        .Inv(n).dur = 0
    Next n
        
    .ArmorSlot = 0
    .WeaponSlot = 0
    .HelmetSlot = 0
    .ShieldSlot = 0
    .PetSlot = 0
    
    .Map = 0
    .x = 0
    .y = 0
    .Dir = 0
    
    .pet.Dir = DIR_DOWN
    .pet.y = 1
    .pet.y = 1
    
    ' Client use only
    .MaxHp = 0
    .MaxMp = 0
    .MaxSP = 0
    .XOffset = 0
    .YOffset = 0
    .Moving = 0
    .Attacking = 0
    .AttackTimer = 0
    .MapGetTimer = 0
    .CastedSpell = NO
    .EmoticonNum = -1
    .EmoticonTime = 0
    .EmoticonVar = 0
    
    For i = 1 To MAX_SPELL_ANIM
        .SpellAnim(i).CastedSpell = NO
        .SpellAnim(i).SpellTime = 0
        .SpellAnim(i).SpellVar = 0
        .SpellAnim(i).SpellDone = 0
        
        .SpellAnim(i).Target = 0
        .SpellAnim(i).TargetType = 0
    Next i
    
    .SpellNum = 0
    
    For i = 1 To MAX_BLT_LINE
        BattlePMsg(i).Index = 1
        BattlePMsg(i).Time = i
        BattleMMsg(i).Index = 1
        BattleMMsg(i).Time = i
    Next i
    
    .QueteEnCour = 0
    
    Inventory = 1
End With
End Sub

Sub ClearPlayerQuete(ByVal Index As Long)
Dim i As Long
With Player(MyIndex)
        .QueteEnCour = 0
        .Quetep.Data1 = 0
        .Quetep.Data2 = 0
        .Quetep.Data3 = 0
        .Quetep.String1 = vbNullString
        Accepter = False
        
        For i = 1 To 15
        .Quetep.indexe(i).Data1 = 0
        .Quetep.indexe(i).Data2 = 0
        .Quetep.indexe(i).Data3 = 0
        .Quetep.indexe(i).String1 = 0
        Next i
End With
End Sub

Sub ClearItem(ByVal Index As Long)
With Item(Index)
    .name = vbNullString
    .desc = vbNullString
    
    .Type = 0
    .Data1 = 0
    .Data2 = 0
    .Data3 = 0
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
End With
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
End Sub

Sub ClearMapItem(ByVal Index As Long)
With MapItem(Index)
    .num = 0
    .Value = 0
    .dur = 0
    .x = 0
    .y = 0
End With
End Sub

Sub ClearMap()
Dim i As Long
Dim x As Long
Dim y As Long

For i = 1 To MAX_MAPS
With Map(i)
    .name = vbNullString
    .Revision = 0
    .Moral = 0
    .Up = 0
    .Down = 0
    .Left = 0
    .Right = 0
    .Indoors = 0
        
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With .Tile(x, y)
            .Ground = 0
            .Mask = 0
            .Anim = 0
            .Mask2 = 0
            .M2Anim = 0
            .Mask3 = 0
            .M3Anim = 0
            .Fringe = 0
            .FAnim = 0
            .Fringe2 = 0
            .F2Anim = 0
            .Fringe3 = 0
            .F3Anim = 0
            .Type = 0
            .Data1 = 0
            .Data2 = 0
            .Data3 = 0
            .String1 = vbNullString
            .String2 = vbNullString
            .String3 = vbNullString
            .Light = 0
            .GroundSet = 0
            .MaskSet = 0
            .AnimSet = 0
            .Mask2Set = 0
            .M2AnimSet = 0
            .Mask3Set = 0
            .M3AnimSet = 0
            .FringeSet = 0
            .FAnimSet = 0
            .Fringe2Set = 0
            .F2AnimSet = 0
            .Fringe3Set = 0
            .F3AnimSet = 0
            End With
        Next x
    Next y
    .PanoInf = vbNullString
    .TranInf = 0
    .PanoSup = vbNullString
    .TranSup = 0
    .Fog = 0
    .FogAlpha = 0
End With
Next i
End Sub

Sub ClearMapItems()
Dim x As Long

    For x = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(x)
    Next x
End Sub

Sub ClearMapNpc(ByVal Index As Long)
With MapNpc(Index)
    .num = 0
    .Target = 0
    .HP = 0
    .MP = 0
    .SP = 0
    .Map = 0
    .x = 0
    .y = 0
    .Dir = 0
    
    ' Client use only
    .XOffset = 0
    .YOffset = 0
    .Moving = 0
    .Attacking = 0
    .AttackTimer = 0
End With
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
    GetPlayerGuild = Trim$(Player(Index).Guild)
End Function

Sub SetPlayerGuild(ByVal Index As Long, ByVal Guild As String)
    Player(Index).Guild = Guild
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
    
    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
        Player(Index).HP = GetPlayerMaxHP(Index)
    End If
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
    GetPlayerMaxMP = Player(Index).MaxMp
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
    GetPlayerMaxSP = Player(Index).MaxSP
End Function

Function GetPlayerstr(ByVal Index As Long) As Long
    GetPlayerstr = Player(Index).STR
End Function

Sub SetPlayerstr(ByVal Index As Long, ByVal STR As Long)
    Player(Index).STR = STR
End Sub

Function GetPlayerDEF(ByVal Index As Long) As Long
    GetPlayerDEF = Player(Index).DEF
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal DEF As Long)
    Player(Index).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal Index As Long) As Long
    GetPlayerSPEED = Player(Index).speed
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal speed As Long)
    Player(Index).speed = speed
End Sub

Function GetPlayerMAGI(ByVal Index As Long) As Long
    GetPlayerMAGI = Player(Index).MAGI
End Function

Sub SetPlayerMAGI(ByVal Index As Long, ByVal MAGI As Long)
    Player(Index).MAGI = MAGI
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
If Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    Player(Index).Map = MapNum
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
    GetPlayerInvItemNum = Player(Index).Inv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Inv(InvSlot).num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Inv(InvSlot).dur
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Inv(InvSlot).dur = ItemDur
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

Sub ClearPet(ByVal Index As Long)
    Pets(Index).nom = ""
    Pets(Index).sprite = 0
    Pets(Index).addForce = 0
    Pets(Index).addDefence = 0
End Sub

Function GetPlayerPetSlot(ByVal Index As Long) As Long
    GetPlayerPetSlot = Player(Index).PetSlot
End Function

Sub SetPlayerPetSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).PetSlot = InvNum
End Sub
