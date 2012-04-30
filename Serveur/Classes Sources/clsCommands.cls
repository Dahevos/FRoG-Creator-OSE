VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'Les 129 commandes du Main

' SummonNpc Legend
' -1 = SubScript Out Of Range
' -2 = Bad GameTime
' -3 = Bad TileType
' -4 = Map is Full

Public Enum OrderType
DoNotMove = 0
AllowToMove = 1
KillNpc = 2
DoSpawn = 3
ChangeDir = 4
GoToDir = 5
GoToPos = 6
GoAttackPlayer = 7
GoAttackNpc = 8
End Enum

Function SummonNpc(ByVal MapNum As Integer, ByVal npcnum As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Dir As Byte) As Integer
Dim Packet As String, i As Long
    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or npcnum <= 0 Or npcnum > MAX_NPCS Or X < 0 Or Y < 0 Or X > MAX_MAPX Or Y > MAX_MAPY Then SummonNpc = -1: Exit Function
    
        If GameTime = TIME_NIGHT Then
            If Npc(npcnum).SpawnTime = 1 Then SummonNpc = -2: Exit Function
        Else
            If Npc(npcnum).SpawnTime = 2 Then SummonNpc = -2: Exit Function
        End If
        
        If Map(MapNum).Tile(X, Y).type <> TILE_TYPE_WALKABLE Then SummonNpc = -3: Exit Function
        
        For i = 1 To MAX_MAP_NPCS
            If Map(MapNum).Npc(i) <= 0 And MapNpc(MapNum, i).Num <= 0 Then Exit For
            If i = MAX_MAP_NPCS Then SummonNpc = -4: Exit Function
        Next i
        
        MapNpc(MapNum, i).Num = npcnum
        MapNpc(MapNum, i).Target = 0
        MapNpc(MapNum, i).TargetType = 0
        
        MapNpc(MapNum, i).HP = GetNpcMaxHP(npcnum)
        MapNpc(MapNum, i).MP = GetNpcMaxMP(npcnum) + IIf(MapNpc(Y, X).Amelio.Timer >= GetTickCount, MapNpc(Y, X).Amelio.Power * 2, 0)
        MapNpc(MapNum, i).SP = GetNpcMaxSP(npcnum)
        
        MapNpc(MapNum, i).X = X
        MapNpc(MapNum, i).Y = Y
        MapNpc(MapNum, i).Dir = Dir
        
        Packet = "SPAWNNPC" & SEP_CHAR & i & SEP_CHAR & npcnum & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Dir & END_CHAR
        Call SendDataToMap(MapNum, Packet)
        
        ' Renvoi le MapNpcNum (Nécéssaire pour un NpcOrder)
        SummonNpc = i
End Function

Function NpcOrder(ByVal MapNum As Integer, ByVal MapNpcNum As Byte, ByVal OType As OrderType, Optional ByVal value As Integer) As Integer
NpcOrder = 0
If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then NpcOrder = -1: Exit Function

If OType = DoNotMove Then
    Map(MapNum).Npcs(MapNpcNum).Imobile = 1
ElseIf OType = AllowToMove Then
    Map(MapNum).Npcs(MapNpcNum).Imobile = 0
ElseIf OType = DoSpawn Then
    Call SpawnNpc(MapNpcNum, MapNum)
ElseIf OType = ChangeDir Then
    If value < 0 Or value > 3 Then NpcOrder = -1: Exit Function
    MapNpc(MapNum, MapNpcNum).Dir = value
ElseIf OType = GoToDir Then
    If value < 0 Or value > 3 Then NpcOrder = -1: Exit Function
    If CanNpcMove(MapNum, MapNpcNum, value) Then Call NpcMove(MapNum, MapNpcNum, value, MOVING_WALKING)
ElseIf OType = GoToPos Then
    If value < 0 Or value > (MAX_MAPX + 1) * (MAX_MAPY + 1) Then NpcOrder = -1: Exit Function
    MapNpc(MapNum, MapNpcNum).Target = value
    MapNpc(MapNum, MapNpcNum).TargetType = TARGET_TYPE_CASE
    'X = value - Int(value / (MAX_MAPX + 1)) * (MAX_MAPX + 1)
    'Y = Int(value / (MAX_MAPX + 1))
ElseIf OType = GoAttackPlayer Then
    If value <= 0 Or value > MAX_PLAYERS Then NpcOrder = -1: Exit Function
    MapNpc(MapNum, MapNpcNum).Target = value
    MapNpc(MapNum, MapNpcNum).TargetType = TARGET_TYPE_PLAYER
ElseIf OType = GoAttackNpc Then
    If value <= 0 Or value > MAX_MAP_NPCS Then NpcOrder = -1: Exit Function
    MapNpc(MapNum, MapNpcNum).Target = value
    MapNpc(MapNum, MapNpcNum).TargetType = TARGET_TYPE_NPC
ElseIf OType = KillNpc Then
    MapNpc(MapNum, MapNpcNum).Num = 0
    MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
    MapNpc(MapNum, MapNpcNum).HP = 0
    Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)
End If
End Function

Public Sub SetPlayerQuete(ByVal Index As Integer, ByVal QuestNum As Integer, Optional ByVal Ask As Boolean = True)
    Dim PacketName As String
    
    If QuestNum < 0 Or QuestNum > MAX_QUETES Then Exit Sub
    
    If Ask Then PacketName = "QUETECOUR" Else PacketName = "SETQUETECOUR"
    Call SendDataTo(Index, PacketName & SEP_CHAR & QuestNum & END_CHAR)
    Player(Index).Char(Player(Index).CharNum).QueteEnCour = QuestNum
    If Ask Then Call QueteMsg(Index, Trim$(quete(QuestNum).nom) & " : " & Trim$(quete(QuestNum).description))
End Sub

Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = vbNullString
  
    sSpaces = Space$(5000)
    
    File = App.Path & "\" & File
    
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Sub PutVar(ByVal File As String, Header As String, Var As String, value As String)
    File = App.Path & "\" & File
    Call WritePrivateProfileString(Header, Var, value, File)
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Long)
Dim Packet As String
    
    Packet = "GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Long)
Dim Packet As String
Dim i As Long

    Packet = "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then Call SendDataTo(i, Packet)
    Next i
End Sub

Public Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Long)
Dim Packet As String

    If Not IsPlaying(Index) Then Exit Sub

    Packet = "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub GuildeMsg(ByVal Index As Long, ByVal Msg As String)
    Dim i As Long
    Dim s As String
    
    If Player(Index).Mute Then Exit Sub
       
    If GetPlayerGuild(Index) = vbNullString Then Call PlayerMsg(Index, "Tu n'es pas dans une guilde!", AlertColor): Exit Sub
    
    s = GetPlayerName(Index) & " (" & GetPlayerGuild(Index) & ") : " & Msg
    Call AddLog(s, PLAYER_LOG)
       
    For i = 1 To MAX_PLAYERS
        If GetPlayerGuild(Index) = GetPlayerGuild(i) Then Call PlayerMsg(i, s, CouleurDesGuilde)
    Next i
End Sub

Public Sub QueteMsg(ByVal Index As Long, ByVal Msg As String)
Dim Packet As String

If Mid(Msg, 1, 3) = "**" Then Msg = Mid(Msg, InStr(1, Msg, ":"))
Packet = "QMSG" & SEP_CHAR & Msg & END_CHAR

Call SendDataTo(Index, Packet)
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Long)
Dim Packet As String
Dim text As String

    Packet = "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
Dim Packet As String

    Packet = "ALERTMSG" & SEP_CHAR & Msg & END_CHAR
    
    Call SendDataTo(Index, Packet)
    Call CloseSocket(Index)
End Sub

' :::::::::::::::::::::::::::::::::::::::::::
' :: Functions/Subs Available To The Owner ::
' :::::::::::::::::::::::::::::::::::::::::::
Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function

Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).Char(Player(Index).CharNum).Name)
End Function

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).Char(Player(Index).CharNum).Guild)
End Function

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

Function GetPlayerClassName(ByVal Index As Long) As String
    GetPlayerClassName = GetVar("Classes\Class" & GetPlayerClass(Index) & ".ini", "CLASS", "Name")
End Function

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
    GetPlayerNextLevel = experience(GetPlayerLevel(Index))
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Char(Player(Index).CharNum).Exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    If Player(Index).Char(Player(Index).CharNum).QueteEnCour > 0 Then If quete(Player(Index).Char(Player(Index).CharNum).QueteEnCour).type = QUETE_TYPE_GAGNE_XP Then Call PlayerQueteTypeXp(Index, Player(Index).Char(Player(Index).CharNum).QueteEnCour, Exp)
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
    GetPlayerStr = Player(Index).Char(Player(Index).CharNum).STR
End Function

Sub SetPlayerStr(ByVal Index As Long, ByVal STR As Long)
    Player(Index).Char(Player(Index).CharNum).STR = STR
End Sub

Function GetPlayerDEF(ByVal Index As Long) As Long
    GetPlayerDEF = Player(Index).Char(Player(Index).CharNum).def
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal def As Long)
    Player(Index).Char(Player(Index).CharNum).def = def
End Sub

Function GetPlayerSPEED(ByVal Index As Long) As Long
    GetPlayerSPEED = Player(Index).Char(Player(Index).CharNum).Speed
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal Speed As Long)
    Player(Index).Char(Player(Index).CharNum).Speed = Speed
End Sub

Function GetPlayerMAGI(ByVal Index As Long) As Long
    GetPlayerMAGI = Player(Index).Char(Player(Index).CharNum).magi
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
    GetPlayerInvItemNum = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Num = ItemNum
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

Function GetBootMap(ByVal Index As Long) As Long
    GetBootMap = Map(GetPlayerMap(Index)).BootMap
End Function

Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long)
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then Exit Sub
        
    If (GetPlayerInvItemNum(Index, InvNum) > 0) And (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
        i = FindOpenMapItemSlot(GetPlayerMap(Index))
        
        If i <> 0 Then
            MapItem(GetPlayerMap(Index), i).Dur = 0
            
            ' Check to see if its any sort of ArmorSlot/WeaponSlot
            Select Case item(GetPlayerInvItemNum(Index, InvNum)).type
                Case ITEM_TYPE_ARMOR
                    If InvNum = GetPlayerArmorSlot(Index) Then Call SetPlayerArmorSlot(Index, 0): Call SendInventory(Index): Call SendWornEquipment(Index)
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                
                Case ITEM_TYPE_WEAPON
                    If InvNum = GetPlayerWeaponSlot(Index) Then Call SetPlayerWeaponSlot(Index, 0): Call SendInventory(Index): Call SendWornEquipment(Index)
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                    
                Case ITEM_TYPE_HELMET
                    If InvNum = GetPlayerHelmetSlot(Index) Then Call SetPlayerHelmetSlot(Index, 0): Call SendInventory(Index): Call SendWornEquipment(Index)
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                                    
                Case ITEM_TYPE_SHIELD
                    If InvNum = GetPlayerShieldSlot(Index) Then Call SetPlayerShieldSlot(Index, 0): Call SendInventory(Index): Call SendWornEquipment(Index)
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
            End Select
                                
            MapItem(GetPlayerMap(Index), i).Num = GetPlayerInvItemNum(Index, InvNum)
            MapItem(GetPlayerMap(Index), i).X = GetPlayerX(Index)
            MapItem(GetPlayerMap(Index), i).Y = GetPlayerY(Index)
                        
            If item(GetPlayerInvItemNum(Index, InvNum)).type = ITEM_TYPE_CURRENCY Then
                ' Check if its more then they have and if so drop it all
                If Amount >= GetPlayerInvItemValue(Index, InvNum) Then
                    MapItem(GetPlayerMap(Index), i).value = GetPlayerInvItemValue(Index, InvNum)
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " dépose " & GetPlayerInvItemValue(Index, InvNum) & " " & Trim$(item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemNum(Index, InvNum, 0)
                    Call SetPlayerInvItemValue(Index, InvNum, 0)
                    Call SetPlayerInvItemDur(Index, InvNum, 0)
                Else
                    MapItem(GetPlayerMap(Index), i).value = Amount
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " dépose " & Amount & " " & Trim$(item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemValue(Index, InvNum, GetPlayerInvItemValue(Index, InvNum) - Amount)
                End If
            Else
                ' Its not a currency object so this is easy
                MapItem(GetPlayerMap(Index), i).value = 1
                If item(GetPlayerInvItemNum(Index, InvNum)).type >= ITEM_TYPE_WEAPON And item(GetPlayerInvItemNum(Index, InvNum)).type <= ITEM_TYPE_SHIELD Then
                    If item(GetPlayerInvItemNum(Index, InvNum)).data1 <= -1 Then
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " dépose un " & Trim$(item(GetPlayerInvItemNum(Index, InvNum)).Name) & " - Ind.", Yellow)
                    Else
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " dépose un " & Trim$(item(GetPlayerInvItemNum(Index, InvNum)).Name) & " - " & GetPlayerInvItemDur(Index, InvNum) & "/" & item(GetPlayerInvItemNum(Index, InvNum)).data1 & ".", Yellow)
                    End If
                Else
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " dépose un " & Trim$(item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                End If
                
                Call SetPlayerInvItemNum(Index, InvNum, 0)
                Call SetPlayerInvItemValue(Index, InvNum, 0)
                Call SetPlayerInvItemDur(Index, InvNum, 0)
            End If
                                        
            ' Send inventory update
            Call SendInventoryUpdate(Index, InvNum)
            ' Spawn the item before we set the num or we'll get a different free map item slot
            Call SpawnItemSlot(i, MapItem(GetPlayerMap(Index), i).Num, Amount, MapItem(GetPlayerMap(Index), i).Dur, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
        Else
            Call PlayerMsg(Index, "Trop d'objets sont déjà au sol.", BrightRed)
        End If
    End If
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
Dim Packet As String
Dim i As Long
    
    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
        
    i = MapItemSlot
    
    If i <> 0 And ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
        MapItem(MapNum, i).Num = ItemNum
        MapItem(MapNum, i).value = ItemVal
        
        If ItemNum <> 0 Then
            If (item(ItemNum).type >= ITEM_TYPE_WEAPON) And (item(ItemNum).type <= ITEM_TYPE_SHIELD) Then
                MapItem(MapNum, i).Dur = ItemDur
            Else
                MapItem(MapNum, i).Dur = 0
            End If
        Else
            MapItem(MapNum, i).Dur = 0
        End If
        
        MapItem(MapNum, i).X = X
        MapItem(MapNum, i).Y = Y
            
        Packet = "SPAWNITEM" & SEP_CHAR & i & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & X & SEP_CHAR & Y & END_CHAR
        Call SendDataToMap(MapNum, Packet)
    End If
End Sub

Function IsConnected(ByVal Index As Long) As Boolean
    If frmServer.Socket(Index).State = sckConnected Then IsConnected = True Else IsConnected = False
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If Index < 1 Or Index > MAX_PLAYERS Then IsPlaying = False: Exit Function
    If IsConnected(Index) And Player(Index).InGame Then IsPlaying = True Else IsPlaying = False
End Function

Sub SendInventory(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = "PLAYERINV" & SEP_CHAR & Index & SEP_CHAR
    For i = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(Index, i) & SEP_CHAR & GetPlayerInvItemValue(Index, i) & SEP_CHAR & GetPlayerInvItemDur(Index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
Dim Packet As String
    
    Packet = "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & Index & SEP_CHAR & GetPlayerInvItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(Index, InvSlot) & SEP_CHAR & Index & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendWornEquipment(ByVal Index As Long)
Dim Packet As String
    If IsPlaying(Index) Then
    'CODE ORIGINAL:
       Packet = "PLAYERWORNEQ" & SEP_CHAR & Index & SEP_CHAR & GetPlayerArmorSlot(Index) & SEP_CHAR & GetPlayerWeaponSlot(Index) & SEP_CHAR & GetPlayerHelmetSlot(Index) & SEP_CHAR & GetPlayerShieldSlot(Index) & SEP_CHAR & GetPlayerPetSlot(Index) & END_CHAR
    'CODE MODIFIE POUR PAPERDOLL:
    'Packet = "PLAYERWORNEQ" & SEP_CHAR & Index & SEP_CHAR & GetPlayerArmorSlot(Index) & SEP_CHAR & GetPlayerWeaponSlot(Index) & SEP_CHAR & GetPlayerHelmetSlot(Index) & SEP_CHAR & GetPlayerShieldSlot(Index) & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).Casque & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).armure & SEP_CHAR & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).arme & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).bouclier & END_CHAR
        Call SendDataToMap(GetPlayerMap(Index), Packet)
    End If
End Sub

Sub SendHP(ByVal Index As Long)
Dim Packet As String, X As Byte

    Packet = "PLAYERHP" & SEP_CHAR & GetPlayerMaxHP(Index) & SEP_CHAR & GetPlayerHP(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
    
    If Player(Index).InParty > 0 Then
        For X = 1 To Party.MemberCount(Player(Index).InParty)
            If Player(Index).PartyPlayer <> X Then Call SendDataTo(Party.PlayerIndex(Player(Index).InParty, X), "partyhp" & SEP_CHAR & Index & SEP_CHAR & Player(Index).InParty & SEP_CHAR & GetPlayerMaxHP(Index) & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).HP & SEP_CHAR & GetPlayerMaxMP(Index) & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).MP & END_CHAR)
        Next X
    End If
    
    Packet = "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMP(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(Index) & SEP_CHAR & GetPlayerMP(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendSP(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(Index) & SEP_CHAR & GetPlayerSP(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendStats(ByVal Index As Long)
Dim Packet As String
    
    Packet = "PLAYERSTATSPACKET" & SEP_CHAR & GetPlayerStr(Index) & SEP_CHAR & GetPlayerDEF(Index) & SEP_CHAR & GetPlayerSPEED(Index) & SEP_CHAR & GetPlayerMAGI(Index) & SEP_CHAR & GetPlayerNextLevel(Index) & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & GetPlayerLevel(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub Flash(ByVal Index As Long, ByVal flashfile As String)
    Call SendDataTo(Index, "flashevent" & SEP_CHAR & flashfile & END_CHAR)
End Sub

Sub Prompt(ByVal Index As Long, ByVal question As String, ByVal value As Long)
    Call SendDataTo(Index, "prompt" & SEP_CHAR & question & SEP_CHAR & value & END_CHAR)
End Sub

Sub PlaySound(ByVal Index As Long, ByVal Sound As String)
    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & END_CHAR)
End Sub

Sub SendPlayerData(ByVal Index As Long)
Dim Packet As String

    ' Send index's player data to everyone including himself on th emap
    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & GetPlayerName(Index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(Index) & SEP_CHAR
    Packet = Packet & GetPlayerMap(Index) & SEP_CHAR
    Packet = Packet & GetPlayerX(Index) & SEP_CHAR
    Packet = Packet & GetPlayerY(Index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(Index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(Index) & SEP_CHAR
    Packet = Packet & GetPlayerLevel(Index) & SEP_CHAR
    Packet = Packet & Player(Index).InParty & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendPlayerQuete(ByVal Index As Long)
Dim Packet As String
Dim i As Long
Packet = "PLAYERQUETE" & SEP_CHAR
Packet = Packet & Player(Index).Char(Player(Index).CharNum).QueteEnCour & SEP_CHAR
Packet = Packet & Player(Index).Char(Player(Index).CharNum).Quetep.data1 & SEP_CHAR
Packet = Packet & Player(Index).Char(Player(Index).CharNum).Quetep.data2 & SEP_CHAR
Packet = Packet & Player(Index).Char(Player(Index).CharNum).Quetep.data3 & SEP_CHAR
Packet = Packet & Player(Index).Char(Player(Index).CharNum).Quetep.String1 & SEP_CHAR

For i = 1 To 15
    Packet = Packet & Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data1 & SEP_CHAR
    Packet = Packet & Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data2 & SEP_CHAR
    Packet = Packet & Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data3 & SEP_CHAR
    Packet = Packet & Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).String1 & SEP_CHAR
Next i
Packet = Packet & END_CHAR
Call SendDataTo(Index, Packet)
End Sub

Sub SendPlayerMetier(ByVal Index As Long)
Dim Packet As String

    ' Send index's player data to everyone including himself on th emap
    Packet = "PLAYERMETIER" & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & Player(Index).Char(Player(Index).CharNum).metier & SEP_CHAR
    Packet = Packet & Player(Index).Char(Player(Index).CharNum).MetierLvl & SEP_CHAR
    Packet = Packet & Player(Index).Char(Player(Index).CharNum).MetierExp & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendDataTo(ByVal Index As Long, ByVal data As String)
Dim i As Long, n As Long, startc As Long
    If IsConnected(Index) Then frmServer.Socket(Index).SendData data: DoEvents
End Sub

Sub SendDataToAll(ByVal data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then Call SendDataTo(i, data)
    Next i
End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByVal data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> Index Then Call SendDataTo(i, data)
    Next i
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then If GetPlayerMap(i) = MapNum Then Call SendDataTo(i, data)
    Next i
End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByVal data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then If GetPlayerMap(i) = MapNum And i <> Index Then Call SendDataTo(i, data)
    Next i
End Sub

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Char(Player(Index).CharNum).Name = Name
End Sub

Function GetPlayerCharNum(ByVal Index As Long) As Long
    GetPlayerCharNum = Player(Index).CharNum
End Function

Sub EnvoieFTP(ByVal FTPURL As String, ByVal UTILISATEUR As String, ByVal mdp As String, ByVal FICHIERLOCAL As String, ByVal FICHIERDISTANT As String, ByVal CHEMINDISTANT As String)
    Call Envoi(FTPURL, UTILISATEUR, mdp, App.Path & FICHIERLOCAL, FICHIERDISTANT, CHEMINDISTANT)
End Sub

Sub TelechargementFTP(ByVal FTPURL As String, ByVal UTILISATEUR As String, ByVal mdp As String, ByVal FICHIERLOCAL As String, ByVal FICHIERDISTANT As String, ByVal CHEMINDISTANT As String)
    Call Telecharger(FTPURL, UTILISATEUR, mdp, App.Path & FICHIERLOCAL, FICHIERDISTANT, CHEMINDISTANT)
End Sub

Sub SuppressionFTP(ByVal FTPURL As String, ByVal UTILISATEUR As String, ByVal mdp As String, ByVal FICHIERDISTANT As String, ByVal CHEMINDISTANT As String)
    Call Supprimer(FTPURL, UTILISATEUR, mdp, FICHIERDISTANT, CHEMINDISTANT)
End Sub

Function FindPlayer(ByVal Name As String) As Long
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then FindPlayer = i: Exit Function
        End If
    Next i
    
    FindPlayer = 0
End Function

Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim i As Long
    
    FindOpenInvSlot = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function
           
    If item(ItemNum).type = ITEM_TYPE_CURRENCY Or item(ItemNum).Empilable <> 0 Then
        ' If currency then check to see if they already have an guildSoloView of the item and add it to that
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If
        Next i
    End If
        
    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(Index, i) <= 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If
    Next i
End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
Dim i As Long

    FindOpenMapItemSlot = 0
    
    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Function
        
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(MapNum, i).Num = 0 Then FindOpenMapItemSlot = i: Exit Function
    Next i
End Function

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
Dim Packet As String
Dim OldMap As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
        
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)
    Call SendLeaveMap(Index, OldMap)
    
    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)
                
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then PlayersOnMap(OldMap) = NO
        
    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES

    Player(Index).GettingMap = YES
    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "warp" & END_CHAR)
    Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & END_CHAR)
    
    Call SendInventory(Index)
   'CODE DE BASE QU'ON DOIT MODIFIE POUR LE PAPERDOLL:
   'Call SendWornEquipment(Index)
   'CODE MODIFIE POUR PAPERDOLL:
   Dim i As Integer
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            Call SendInventory(i)
            Call SendWornEquipment(i)
            'Call PlayerPet(i, 0, GetPlayerDir(i))
            Call PetMove(i)
        End If
    Next
End Sub

Sub AddLog(ByVal text As String, ByVal FN As String)
Dim FileName As String
Dim f As Long

    If ServerLog = True Then
        FileName = App.Path & "\logs\" & FN & ".txt"
    
        If Not FileExist(FN) Then
            f = FreeFile
            Open FileName For Output As #f
            Close #f
        End If
    
        f = FreeFile
        Open FileName For Append As #f
            Print #f, Time & ": " & text
        Close #f
    End If
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If Index > 0 Then
        If IsPlaying(Index) Then Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " a été déconnecté pour : (" & Reason & ")", White)
        Call AlertMsg(Index, "Tu as perdu ta connexion avec " & GAME_NAME & "." & Reason)
        Call AffIBMsg(Index, "ATTENTION : Détection d'une tentative de hacking!!(Raison : " & Reason & " Login : " & GetPlayerLogin(Index) & " perso : " & GetPlayerName(Index) & ").", BrightRed, True)
    End If
End Sub

Sub BattleMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Long, ByVal Side As Long)
    Call SendDataTo(Index, "damagedisplay" & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Sub ContrOnOff(ByVal Index As Long)
Dim Packet As String

Packet = "CONOFF" & END_CHAR

Call SendDataTo(Index, Packet)
End Sub

Function Rand(ByVal High As Long, ByVal Low As Long)
Randomize
High = High + 1
Do Until Rand >= Low
    Randomize
    Rand = Int(Rnd * High)
Loop
End Function

Sub SetPlayerGuild(ByVal Index As Long, ByVal Guild As String)
Player(Index).Char(Player(Index).CharNum).Guild = Guild
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

Sub KillPlayer(ByVal Index As Long)
    Player(Index).Char(Player(Index).CharNum).HP = 0
    
    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then Player(Index).Char(Player(Index).CharNum).HP = GetPlayerMaxHP(Index)
    If GetPlayerHP(Index) < 0 Then Player(Index).Char(Player(Index).CharNum).HP = 0
    Call SendStats(Index)
End Sub

Sub Pluie()
    GameWeather = WEATHER_RAINING
    Call SendWeatherToAll
End Sub

Sub Neige()
    GameWeather = WEATHER_SNOWING
    Call SendWeatherToAll
End Sub

Sub Orage()
    GameWeather = WEATHER_THUNDER
    Call SendWeatherToAll
End Sub

Sub Soleil()
    GameWeather = WEATHER_NONE
    Call SendWeatherToAll
End Sub

Sub Jour()
    GameTime = TIME_DAY
    Call SendTimeToAll
End Sub

Sub Nuit()
    GameTime = TIME_NIGHT
    Call SendTimeToAll
End Sub

Public Sub NetObjsCarte(MapNum As Long)
Dim i As Long

For i = 1 To MAX_MAP_ITEMS
    Call ClearMapItem(i, MapNum)
Next

End Sub

Public Sub NetObjsCartes()
    Call ClearMapItems
End Sub

Sub Lancer(ByVal Index As Long, ByVal lien As String)
Dim Packet As String
    Packet = "LANCE" & SEP_CHAR & lien & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SetPlayerItem(Index, item)
Dim n
n = 1
Do
   If GetPlayerInvItemNum(Index, n) = 0 Then
      Call SetPlayerInvItemNum(Index, n, item)
      Call SendInventoryUpdate(Index, n)
      Exit Do
   End If
   n = n + 1
Loop Until n > 24
End Sub

Sub SetPlayerItemValue(Index, item, value)
Dim n
n = 1
Do
   If GetPlayerInvItemNum(Index, n) = item Then
      Call SetPlayerInvItemValue(Index, n, value)
      Call SendInventoryUpdate(Index, n)
      Exit Do
   End If
   n = n + 1
Loop Until n > 24
End Sub

Public Function GetPlayerQueteEtat(ByVal PIndex As Long, ByVal qindex As Long) As Boolean
    GetPlayerQueteEtat = False

    If Player(PIndex).Char(Player(PIndex).CharNum).QueteStatut(qindex) = 2 Then
        GetPlayerQueteEtat = True
    Else
        GetPlayerQueteEtat = False
    End If
End Function

Function GetPlayerHPRegen(ByVal Index As Long)
Dim i As Long
    
    GetPlayerHPRegen = 0
    
    If Val(GetVar(App.Path & "\Data.ini", "CONFIG", "HPRegen")) >= 1 Then
        ' Prevent subscript out of range
        If Not IsPlaying(Index) Or Index <= 0 Or Index > MAX_PLAYERS Then GetPlayerHPRegen = 0: Exit Function
        
        i = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "HPRegen")) '(GetPlayerDEF(Index) \ 2)
        If i < 2 Then i = 2
        
        GetPlayerHPRegen = i
    End If
End Function

Function GetPlayerMPRegen(ByVal Index As Long)
Dim i As Long
    
    GetPlayerMPRegen = 0
    
    If Val(GetVar(App.Path & "\Data.ini", "CONFIG", "MPRegen")) >= 1 Then
        ' Prevent subscript out of range
        If Not IsPlaying(Index) Or Index <= 0 Or Index > MAX_PLAYERS Then GetPlayerMPRegen = 0: Exit Function
        
        i = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "MPRegen")) '(GetPlayerMAGI(Index) \ 2)
        If i < 2 Then i = 2
        
        GetPlayerMPRegen = i
    End If
End Function

Function GetPlayerSPRegen(ByVal Index As Long)
Dim i As Long
    
    GetPlayerSPRegen = 0
    
    If Val(GetVar(App.Path & "\Data.ini", "CONFIG", "SPRegen")) >= 1 Then
        ' Prevent subscript out of range
        If Not IsPlaying(Index) Or Index <= 0 Or Index > MAX_PLAYERS Then GetPlayerSPRegen = 0: Exit Function
                
        i = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "SPRegen")) '(GetPlayerSPEED(Index) \ 2)
        If i < 2 Then i = 2
        
        GetPlayerSPRegen = i
    End If
End Function

Sub SuprObjet(ByVal Index As Long, ByVal ItemNum As Long)
Dim i As Long

For i = 1 To MAX_INV
    If GetPlayerInvItemNum(Index, i) = ItemNum Then
        If GetPlayerInvItemValue(Index, i) > 1 Then
            Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - 1)
        Else
            Call SetPlayerInvItemNum(Index, i, 0)
            Call SetPlayerInvItemValue(Index, i, 0)
            Call SetPlayerInvItemDur(Index, i, 0)
        End If
        Exit Sub
    End If
Next i

End Sub

Sub AjoutObjet(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long

i = FindOpenInvSlot(Index, ItemNum)
If i = 0 Then Call PlayerMsg(Index, "Ton inventaire est plein", Red): Exit Sub
Call SetPlayerInvItemNum(Index, i, ItemNum)
Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + ItemVal)
If (item(ItemNum).type >= ITEM_TYPE_WEAPON) And (item(ItemNum).type <= ITEM_TYPE_SHIELD) Then Call SetPlayerInvItemDur(Index, i, item(ItemNum).data1) Else Call SetPlayerInvItemDur(Index, i, 0)
Call SendInventory(Index)
End Sub

Sub ChangeClasse(ByVal Index As Long, ByVal Classes As Long)

If Val(Classes) < 0 Or Val(Classes) > Max_Classes Then Exit Sub

If Player(Index).Char(Player(Index).CharNum).Sex = 0 Then
    If GetPlayerSprite(Index) <> Classe(GetPlayerClass(Index)).MaleSprite Then
        Call SetPlayerSprite(Index, Classe(Classes).MaleSprite)
    End If
Else
    If GetPlayerSprite(Index) <> Classe(GetPlayerClass(Index)).FemaleSprite Then
        Call SetPlayerSprite(Index, Classe(Classes).FemaleSprite)
    End If
End If

Call SetPlayerStr(Index, (Player(Index).Char(Player(Index).CharNum).STR - Classe(GetPlayerClass(Index)).STR))
Call SetPlayerDEF(Index, (Player(Index).Char(Player(Index).CharNum).def - Classe(GetPlayerClass(Index)).def))
Call SetPlayerMAGI(Index, (Player(Index).Char(Player(Index).CharNum).magi - Classe(GetPlayerClass(Index)).magi))
Call SetPlayerSPEED(Index, (Player(Index).Char(Player(Index).CharNum).Speed - Classe(GetPlayerClass(Index)).Speed))

Call SetPlayerClass(Index, Classes)

Call SetPlayerStr(Index, (Player(Index).Char(Player(Index).CharNum).STR + Classe(GetPlayerClass(Index)).STR + GetVar("Classes\Class" & GetPlayerClass(Index) & ".ini", "CLASSCHANGE", "AddStr")))
Call SetPlayerDEF(Index, (Player(Index).Char(Player(Index).CharNum).def + Classe(GetPlayerClass(Index)).def + GetVar("Classes\Class" & GetPlayerClass(Index) & ".ini", "CLASSCHANGE", "AddDef")))
Call SetPlayerMAGI(Index, (Player(Index).Char(Player(Index).CharNum).magi + Classe(GetPlayerClass(Index)).magi + GetVar("Classes\Class" & GetPlayerClass(Index) & ".ini", "CLASSCHANGE", "AddMagi")))
Call SetPlayerSPEED(Index, (Player(Index).Char(Player(Index).CharNum).Speed + Classe(GetPlayerClass(Index)).Speed + GetVar("Classes\Class" & GetPlayerClass(Index) & ".ini", "CLASSCHANGE", "AddSpeed")))


Call PlayerMsg(Index, "Ta nouvelle classe est " & Trim$(Classe(GetPlayerClass(Index)).Name) & "!", BrightGreen)

Call SendStats(Index)
Call SendHP(Index)
Call SendMP(Index)
Call SendSP(Index)
Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
End Sub

Sub SauvJoueur(ByVal Index As Long)
    Call SavePlayer(Index)
End Sub

Function CibleJoueur(ByVal Index As Long) As Long
CibleJoueur = Player(Index).Target
End Function

Function TypeCibleJoueur(ByVal Index As Long) As Byte
TypeCibleJoueur = Player(Index).TargetType
End Function

Function ASort(ByVal Index As Long, ByVal SpellNum As Long) As Long
Dim i As Long

    ASort = 0
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, i) = SpellNum Then ASort = i: Exit Function
    Next i
End Function

Function AObjet(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim i As Long
    
    AObjet = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function
        
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            AObjet = i
            Exit Function
        End If
    Next i
End Function

Sub AppObjet(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
        
    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    
    Call SpawnItemSlot(i, ItemNum, ItemVal, item(ItemNum).data1, MapNum, X, Y)
End Sub

Sub Bloque(ByVal Index As Long)
If GetPlayerDir(Index) = DIR_UP Then
    Call SetPlayerY(Index, GetPlayerY(Index) + 1)
ElseIf GetPlayerDir(Index) = DIR_DOWN Then
    Call SetPlayerY(Index, GetPlayerY(Index) - 1)
ElseIf GetPlayerDir(Index) = DIR_RIGHT Then
    Call SetPlayerX(Index, GetPlayerX(Index) - 1)
ElseIf GetPlayerDir(Index) = DIR_LEFT Then
    Call SetPlayerX(Index, GetPlayerX(Index) + 1)
End If
Call SendPlayerData(Index)
End Sub

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

Function GetTime() As Byte
    'Jour = 0, Nuit = 1
    GetTime = GameTime
End Function

Function GetCaseBloque(ByVal M As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    GetCaseBloque = False
    
    If Map(M).Tile(X, Y).type = TILE_TYPE_CBLOCK Then
        GetCaseBloque = True
    End If
    
End Function

Public Sub PlayerAnimStart(ByVal Index As Long, ByVal Anim As Long)
Dim Packet As String

    If Not IsPlaying(Index) Then Exit Sub

    Packet = "PLAYERANIMSTART" & SEP_CHAR & Index & SEP_CHAR & Anim & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Public Sub PlayerAnimStop(ByVal Index As Long)
Dim Packet As String

    If Not IsPlaying(Index) Then Exit Sub

    Packet = "PLAYERANIMSTOP" & SEP_CHAR & Index & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Public Sub PlayerAnim(ByVal Index As Long, ByVal Anim As Long, ByVal ms As Long)
Dim Packet As String

    If Not IsPlaying(Index) Then Exit Sub

    Packet = "PLAYERANIM" & SEP_CHAR & Index & SEP_CHAR & Anim & SEP_CHAR & ms & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Public Sub PlayerAnimRetour(ByVal Index As Long, ByVal Anim As Long, ByVal ms As Long, ByVal script As Long)
Dim Packet As String

    If Not IsPlaying(Index) Then Exit Sub

    Packet = "PLAYERANIMRT" & SEP_CHAR & Index & SEP_CHAR & Anim & SEP_CHAR & ms & SEP_CHAR & script & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

