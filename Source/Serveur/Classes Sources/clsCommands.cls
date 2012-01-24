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
            If Map(MapNum).Npc(i) <= 0 And MapNpc(MapNum, i).num <= 0 Then Exit For
            If i = MAX_MAP_NPCS Then SummonNpc = -4: Exit Function
        Next i
        
        MapNpc(MapNum, i).num = npcnum
        MapNpc(MapNum, i).Target = 0
        MapNpc(MapNum, i).TargetType = 0
        
        MapNpc(MapNum, i).HP = GetNpcMaxHP(npcnum)
        MapNpc(MapNum, i).MP = GetNpcMaxMP(npcnum) + IIf(MapNpc(Y, X).Amelio.Timer >= GetTickCount, MapNpc(Y, X).Amelio.Power * 2, 0)
        MapNpc(MapNum, i).SP = GetNpcMaxSP(npcnum)
        
        MapNpc(MapNum, i).X = X
        MapNpc(MapNum, i).Y = Y
        MapNpc(MapNum, i).Dir = Dir
        
        Packet = "SPAWNNPC" & SEP_CHAR & i & SEP_CHAR & npcnum & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Dir & SEP_CHAR & END_CHAR
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
    MapNpc(MapNum, MapNpcNum).num = 0
    MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
    MapNpc(MapNum, MapNpcNum).HP = 0
    Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
End If
End Function

Public Sub SetPlayerQuete(ByVal index As Integer, ByVal QuestNum As Integer, Optional ByVal Ask As Boolean = True)
    Dim PacketName As String
    
    If QuestNum < 0 Or QuestNum > MAX_QUETES Then Exit Sub
    
    If Ask Then PacketName = "QUETECOUR" Else PacketName = "SETQUETECOUR"
    Call SendDataTo(index, PacketName & SEP_CHAR & QuestNum & SEP_CHAR & END_CHAR)
    Player(index).Char(Player(index).CharNum).QueteEnCour = QuestNum
    If Ask Then Call QueteMsg(index, Trim$(quete(QuestNum).nom) & " : " & Trim$(quete(QuestNum).description))
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
    
    Packet = "GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Long)
Dim Packet As String
Dim i As Long

    Packet = "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then Call SendDataTo(i, Packet)
    Next i
End Sub

Public Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Long)
Dim Packet As String

    If Not IsPlaying(index) Then Exit Sub

    Packet = "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub GuildeMsg(ByVal index As Long, ByVal Msg As String)
    Dim i As Long
    Dim s As String
    
    If Player(index).Mute Then Exit Sub
       
    If GetPlayerGuild(index) = vbNullString Then Call PlayerMsg(index, "Tu n'es pas dans une guilde!", AlertColor): Exit Sub
    
    s = GetPlayerName(index) & " (" & GetPlayerGuild(index) & ") : " & Msg
    Call AddLog(s, PLAYER_LOG)
       
    For i = 1 To MAX_PLAYERS
        If GetPlayerGuild(index) = GetPlayerGuild(i) Then Call PlayerMsg(i, s, CouleurDesGuilde)
    Next i
End Sub

Public Sub QueteMsg(ByVal index As Long, ByVal Msg As String)
Dim Packet As String

If Mid(Msg, 1, 3) = "**" Then Msg = Mid(Msg, InStr(1, Msg, ":"))
Packet = "QMSG" & SEP_CHAR & Msg & SEP_CHAR & END_CHAR

Call SendDataTo(index, Packet)
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Long)
Dim Packet As String
Dim text As String

    Packet = "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Public Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
Dim Packet As String

    Packet = "ALERTMSG" & SEP_CHAR & Msg & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
    Call CloseSocket(index)
End Sub

' :::::::::::::::::::::::::::::::::::::::::::
' :: Functions/Subs Available To The Owner ::
' :::::::::::::::::::::::::::::::::::::::::::
Function GetPlayerLogin(ByVal index As Long) As String
    GetPlayerLogin = Trim$(Player(index).Login)
End Function

Function GetPlayerName(ByVal index As Long) As String
    GetPlayerName = Trim$(Player(index).Char(Player(index).CharNum).Name)
End Function

Function GetPlayerGuild(ByVal index As Long) As String
    GetPlayerGuild = Trim$(Player(index).Char(Player(index).CharNum).Guild)
End Function

Function GetPlayerGuildAccess(ByVal index As Long) As Long
    GetPlayerGuildAccess = Player(index).Char(Player(index).CharNum).Guildaccess
End Function

Sub SetPlayerGuildAccess(ByVal index As Long, ByVal Guildaccess As Long)
    Player(index).Char(Player(index).CharNum).Guildaccess = Guildaccess
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    GetPlayerClass = Player(index).Char(Player(index).CharNum).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    Player(index).Char(Player(index).CharNum).Class = ClassNum
End Sub

Function GetPlayerClassName(ByVal index As Long) As String
    GetPlayerClassName = GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "CLASS", "Name")
End Function

Function GetPlayerSprite(ByVal index As Long) As Long
    GetPlayerSprite = Player(index).Char(Player(index).CharNum).sprite
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal sprite As Long)
    Player(index).Char(Player(index).CharNum).sprite = sprite
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long
    GetPlayerLevel = Player(index).Char(Player(index).CharNum).Level
End Function

Sub SetPlayerLevel(ByVal index As Long, ByVal Level As Long)
    If GetPlayerLevel(index) > MAX_LEVEL Then Exit Sub
    Player(index).Char(Player(index).CharNum).Level = Level
End Sub

Function GetPlayerNextLevel(ByVal index As Long) As Long
    If GetPlayerLevel(index) > MAX_LEVEL Then Exit Function
    GetPlayerNextLevel = experience(GetPlayerLevel(index))
End Function

Function GetPlayerExp(ByVal index As Long) As Long
    GetPlayerExp = Player(index).Char(Player(index).CharNum).Exp
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal Exp As Long)
    If Player(index).Char(Player(index).CharNum).QueteEnCour > 0 Then If quete(Player(index).Char(Player(index).CharNum).QueteEnCour).type = QUETE_TYPE_GAGNE_XP Then Call PlayerQueteTypeXp(index, Player(index).Char(Player(index).CharNum).QueteEnCour, Exp)
    Player(index).Char(Player(index).CharNum).Exp = Exp
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long
    GetPlayerAccess = Player(index).Char(Player(index).CharNum).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    Player(index).Char(Player(index).CharNum).Access = Access
End Sub

Function GetPlayerPK(ByVal index As Long) As Long
    GetPlayerPK = Player(index).Char(Player(index).CharNum).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    Player(index).Char(Player(index).CharNum).PK = PK
End Sub

Function GetPlayerHP(ByVal index As Long) As Long
    GetPlayerHP = Player(index).Char(Player(index).CharNum).HP
End Function

Sub SetPlayerHP(ByVal index As Long, ByVal HP As Long)
    Player(index).Char(Player(index).CharNum).HP = HP
    
    If GetPlayerHP(index) > GetPlayerMaxHP(index) Then Player(index).Char(Player(index).CharNum).HP = GetPlayerMaxHP(index)
    If GetPlayerHP(index) < 0 Then Player(index).Char(Player(index).CharNum).HP = 0
    Call SendStats(index)
End Sub

Function GetPlayerMP(ByVal index As Long) As Long
    GetPlayerMP = Player(index).Char(Player(index).CharNum).MP
End Function

Sub SetPlayerMP(ByVal index As Long, ByVal MP As Long)
    Player(index).Char(Player(index).CharNum).MP = MP

    If GetPlayerMP(index) > GetPlayerMaxMP(index) Then Player(index).Char(Player(index).CharNum).MP = GetPlayerMaxMP(index)
    If GetPlayerMP(index) < 0 Then Player(index).Char(Player(index).CharNum).MP = 0
End Sub

Function GetPlayerSP(ByVal index As Long) As Long
    GetPlayerSP = Player(index).Char(Player(index).CharNum).SP
End Function

Sub SetPlayerSP(ByVal index As Long, ByVal SP As Long)
    Player(index).Char(Player(index).CharNum).SP = SP

    If GetPlayerSP(index) > GetPlayerMaxSP(index) Then Player(index).Char(Player(index).CharNum).SP = GetPlayerMaxSP(index)
    If GetPlayerSP(index) < 0 Then Player(index).Char(Player(index).CharNum).SP = 0
End Sub

Function GetPlayerMaxHP(ByVal index As Long) As Long
Dim CharNum As Long
Dim i As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(index) > 0 Then add = item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddHP
    If GetPlayerArmorSlot(index) > 0 Then add = add + item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddHP
    If GetPlayerShieldSlot(index) > 0 Then add = add + item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddHP
    If GetPlayerHelmetSlot(index) > 0 Then add = add + item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddHP
    
    CharNum = Player(index).CharNum
    'GetPlayerMaxHP = ((Player(index).Char(CharNum).Level + Int(GetPlayerstr(index) / 2) + ClassE(Player(index).Char(CharNum).Class).STR) * 2) + add
    GetPlayerMaxHP = (GetPlayerLevel(index) * AddHP.Level) + (GetPlayerStr(index) * AddHP.STR) + (GetPlayerDEF(index) * AddHP.def) + (GetPlayerMAGI(index) * AddHP.magi) + (GetPlayerSPEED(index) * AddHP.Speed) + add
End Function

Function GetPlayerMaxMP(ByVal index As Long) As Long
Dim CharNum As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(index) > 0 Then add = item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddMP
    If GetPlayerArmorSlot(index) > 0 Then add = add + item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddMP
    If GetPlayerShieldSlot(index) > 0 Then add = add + item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddMP
    If GetPlayerHelmetSlot(index) > 0 Then add = add + item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddMP
    
    CharNum = Player(index).CharNum
    'GetPlayerMaxMP = ((Player(index).Char(CharNum).Level + Int(GetPlayerMAGI(index) / 2) + Class(Player(index).Char(CharNum).Class).MAGI) * 2) + add
    GetPlayerMaxMP = (GetPlayerLevel(index) * AddMP.Level) + (GetPlayerStr(index) * AddMP.STR) + (GetPlayerDEF(index) * AddMP.def) + (GetPlayerMAGI(index) * AddMP.magi) + (GetPlayerSPEED(index) * AddMP.Speed) + add
End Function

Function GetPlayerMaxSP(ByVal index As Long) As Long
Dim CharNum As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(index) > 0 Then add = item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddSP
    If GetPlayerArmorSlot(index) > 0 Then add = add + item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddSP
    If GetPlayerShieldSlot(index) > 0 Then add = add + item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddSP
    If GetPlayerHelmetSlot(index) > 0 Then add = add + item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddSP
    
    CharNum = Player(index).CharNum
    'GetPlayerMaxSP = ((Player(index).Char(CharNum).Level + Int(GetPlayerSPEED(index) / 2) + Class(Player(index).Char(CharNum).Class).SPEED) * 2) + add
    GetPlayerMaxSP = (GetPlayerLevel(index) * AddSP.Level) + (GetPlayerStr(index) * AddSP.STR) + (GetPlayerDEF(index) * AddSP.def) + (GetPlayerMAGI(index) * AddSP.magi) + (GetPlayerSPEED(index) * AddSP.Speed) + add
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

Function GetPlayerStr(ByVal index As Long) As Long
    GetPlayerStr = Player(index).Char(Player(index).CharNum).STR
End Function

Sub SetPlayerStr(ByVal index As Long, ByVal STR As Long)
    Player(index).Char(Player(index).CharNum).STR = STR
End Sub

Function GetPlayerDEF(ByVal index As Long) As Long
    GetPlayerDEF = Player(index).Char(Player(index).CharNum).def
End Function

Sub SetPlayerDEF(ByVal index As Long, ByVal def As Long)
    Player(index).Char(Player(index).CharNum).def = def
End Sub

Function GetPlayerSPEED(ByVal index As Long) As Long
    GetPlayerSPEED = Player(index).Char(Player(index).CharNum).Speed
End Function

Sub SetPlayerSPEED(ByVal index As Long, ByVal Speed As Long)
    Player(index).Char(Player(index).CharNum).Speed = Speed
End Sub

Function GetPlayerMAGI(ByVal index As Long) As Long
    GetPlayerMAGI = Player(index).Char(Player(index).CharNum).magi
End Function

Sub SetPlayerMAGI(ByVal index As Long, ByVal magi As Long)
    Player(index).Char(Player(index).CharNum).magi = magi
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long
    GetPlayerPOINTS = Player(index).Char(Player(index).CharNum).POINTS
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    Player(index).Char(Player(index).CharNum).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal index As Long) As Long
    GetPlayerMap = Player(index).Char(Player(index).CharNum).Map
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal MapNum As Long)
    If MapNum > 0 And MapNum <= MAX_MAPS Then Player(index).Char(Player(index).CharNum).Map = MapNum
End Sub

Function GetPlayerX(ByVal index As Long) As Long
    GetPlayerX = Player(index).Char(Player(index).CharNum).X
End Function

Sub SetPlayerX(ByVal index As Long, ByVal X As Long)
    Player(index).Char(Player(index).CharNum).X = X
End Sub

Function GetPlayerY(ByVal index As Long) As Long
    GetPlayerY = Player(index).Char(Player(index).CharNum).Y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal Y As Long)
    Player(index).Char(Player(index).CharNum).Y = Y
End Sub

<<<<<<< HEAD
Function GetPlayerSex(ByVal index As Long) As Byte
    GetPlayerSex = Player(index).Char(Player(index).CharNum).Sex
End Function

Sub SetPlayerSex(ByVal index As Long, ByVal Sex As Byte)
    Player(index).Char(Player(index).CharNum).Sex = Sex
=======
Function GetPlayerSex(ByVal Index As Long) As Byte
    GetPlayerSex = Player(Index).Char(Player(Index).CharNum).Sex
End Function

Sub SetPlayerSex(ByVal Index As Long, ByVal Sex As Byte)
    Player(Index).Char(Player(Index).CharNum).Sex = Sex
>>>>>>> 78e891abe783f077a019bb7ae3b3f68bd63dddde
End Sub

Function GetPlayerDir(ByVal index As Long) As Long
    GetPlayerDir = Player(index).Char(Player(index).CharNum).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    Player(index).Char(Player(index).CharNum).Dir = Dir
End Sub

Function GetPlayerIP(ByVal index As Long) As String
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(index).Char(Player(index).CharNum).Inv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(index).Char(Player(index).CharNum).Inv(InvSlot).value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(index).Char(Player(index).CharNum).Spell(SpellSlot)
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
    Player(index).Char(Player(index).CharNum).Spell(SpellSlot) = SpellNum
End Sub

Function GetPlayerArmorSlot(ByVal index As Long) As Long
    GetPlayerArmorSlot = Player(index).Char(Player(index).CharNum).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal index As Long) As Long
    GetPlayerWeaponSlot = Player(index).Char(Player(index).CharNum).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal index As Long) As Long
    GetPlayerHelmetSlot = Player(index).Char(Player(index).CharNum).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal index As Long) As Long
    GetPlayerShieldSlot = Player(index).Char(Player(index).CharNum).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).ShieldSlot = InvNum
End Sub

Function GetBootMap(ByVal index As Long) As Long
    GetBootMap = Map(GetPlayerMap(index)).BootMap
End Function

Sub PlayerMapDropItem(ByVal index As Long, ByVal InvNum As Long, ByVal Amount As Long)
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then Exit Sub
        
    If (GetPlayerInvItemNum(index, InvNum) > 0) And (GetPlayerInvItemNum(index, InvNum) <= MAX_ITEMS) Then
        i = FindOpenMapItemSlot(GetPlayerMap(index))
        
        If i <> 0 Then
            MapItem(GetPlayerMap(index), i).Dur = 0
            
            ' Check to see if its any sort of ArmorSlot/WeaponSlot
            Select Case item(GetPlayerInvItemNum(index, InvNum)).type
                Case ITEM_TYPE_ARMOR
                    If InvNum = GetPlayerArmorSlot(index) Then Call SetPlayerArmorSlot(index, 0): Call SendInventory(index): Call SendWornEquipment(index)
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                
                Case ITEM_TYPE_WEAPON
                    If InvNum = GetPlayerWeaponSlot(index) Then Call SetPlayerWeaponSlot(index, 0): Call SendInventory(index): Call SendWornEquipment(index)
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                    
                Case ITEM_TYPE_HELMET
                    If InvNum = GetPlayerHelmetSlot(index) Then Call SetPlayerHelmetSlot(index, 0): Call SendInventory(index): Call SendWornEquipment(index)
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                                    
                Case ITEM_TYPE_SHIELD
                    If InvNum = GetPlayerShieldSlot(index) Then Call SetPlayerShieldSlot(index, 0): Call SendInventory(index): Call SendWornEquipment(index)
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
            End Select
                                
            MapItem(GetPlayerMap(index), i).num = GetPlayerInvItemNum(index, InvNum)
            MapItem(GetPlayerMap(index), i).X = GetPlayerX(index)
            MapItem(GetPlayerMap(index), i).Y = GetPlayerY(index)
                        
            If item(GetPlayerInvItemNum(index, InvNum)).type = ITEM_TYPE_CURRENCY Then
                ' Check if its more then they have and if so drop it all
                If Amount >= GetPlayerInvItemValue(index, InvNum) Then
                    MapItem(GetPlayerMap(index), i).value = GetPlayerInvItemValue(index, InvNum)
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " dépose " & GetPlayerInvItemValue(index, InvNum) & " " & Trim$(item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemNum(index, InvNum, 0)
                    Call SetPlayerInvItemValue(index, InvNum, 0)
                    Call SetPlayerInvItemDur(index, InvNum, 0)
                Else
                    MapItem(GetPlayerMap(index), i).value = Amount
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " dépose " & Amount & " " & Trim$(item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemValue(index, InvNum, GetPlayerInvItemValue(index, InvNum) - Amount)
                End If
            Else
                ' Its not a currency object so this is easy
                MapItem(GetPlayerMap(index), i).value = 1
                If item(GetPlayerInvItemNum(index, InvNum)).type >= ITEM_TYPE_WEAPON And item(GetPlayerInvItemNum(index, InvNum)).type <= ITEM_TYPE_SHIELD Then
                    If item(GetPlayerInvItemNum(index, InvNum)).data1 <= -1 Then
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " dépose un " & Trim$(item(GetPlayerInvItemNum(index, InvNum)).Name) & " - Ind.", Yellow)
                    Else
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " dépose un " & Trim$(item(GetPlayerInvItemNum(index, InvNum)).Name) & " - " & GetPlayerInvItemDur(index, InvNum) & "/" & item(GetPlayerInvItemNum(index, InvNum)).data1 & ".", Yellow)
                    End If
                Else
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " dépose un " & Trim$(item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                End If
                
                Call SetPlayerInvItemNum(index, InvNum, 0)
                Call SetPlayerInvItemValue(index, InvNum, 0)
                Call SetPlayerInvItemDur(index, InvNum, 0)
            End If
                                        
            ' Send inventory update
            Call SendInventoryUpdate(index, InvNum)
            ' Spawn the item before we set the num or we'll get a different free map item slot
            Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).num, Amount, MapItem(GetPlayerMap(index), i).Dur, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
        Else
            Call PlayerMsg(index, "Trop d'objets sont déjà au sol.", BrightRed)
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
        MapItem(MapNum, i).num = ItemNum
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
            
        Packet = "SPAWNITEM" & SEP_CHAR & i & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, Packet)
    End If
End Sub

Function IsConnected(ByVal index As Long) As Boolean
    If frmServer.Socket(index).State = sckConnected Then IsConnected = True Else IsConnected = False
End Function

Function IsPlaying(ByVal index As Long) As Boolean
    If index < 1 Or index > MAX_PLAYERS Then IsPlaying = False: Exit Function
    If IsConnected(index) And Player(index).InGame Then IsPlaying = True Else IsPlaying = False
End Function

Sub SendInventory(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "PLAYERINV" & SEP_CHAR & index & SEP_CHAR
    For i = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(index, i) & SEP_CHAR & GetPlayerInvItemValue(index, i) & SEP_CHAR & GetPlayerInvItemDur(index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal InvSlot As Long)
Dim Packet As String
    
    Packet = "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & index & SEP_CHAR & GetPlayerInvItemNum(index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(index, InvSlot) & SEP_CHAR & index & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendWornEquipment(ByVal index As Long)
Dim Packet As String
    If IsPlaying(index) Then
    'CODE ORIGINAL:
       Packet = "PLAYERWORNEQ" & SEP_CHAR & index & SEP_CHAR & GetPlayerArmorSlot(index) & SEP_CHAR & GetPlayerWeaponSlot(index) & SEP_CHAR & GetPlayerHelmetSlot(index) & SEP_CHAR & GetPlayerShieldSlot(index) & SEP_CHAR & GetPlayerPetSlot(index) & SEP_CHAR & END_CHAR
    'CODE MODIFIE POUR PAPERDOLL:
    'Packet = "PLAYERWORNEQ" & SEP_CHAR & Index & SEP_CHAR & GetPlayerArmorSlot(Index) & SEP_CHAR & GetPlayerWeaponSlot(Index) & SEP_CHAR & GetPlayerHelmetSlot(Index) & SEP_CHAR & GetPlayerShieldSlot(Index) & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).Casque & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).armure & SEP_CHAR & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).arme & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).bouclier & END_CHAR
        Call SendDataToMap(GetPlayerMap(index), Packet)
    End If
End Sub

Sub SendHP(ByVal index As Long)
Dim Packet As String, X As Byte

    Packet = "PLAYERHP" & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & GetPlayerHP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
    
    If Player(index).InParty > 0 Then
        For X = 1 To Party.MemberCount(Player(index).InParty)
            If Player(index).PartyPlayer <> X Then Call SendDataTo(Party.PlayerIndex(Player(index).InParty, X), "partyhp" & SEP_CHAR & index & SEP_CHAR & Player(index).InParty & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & Player(index).Char(Player(index).CharNum).HP & SEP_CHAR & GetPlayerMaxMP(index) & SEP_CHAR & Player(index).Char(Player(index).CharNum).MP & SEP_CHAR & END_CHAR)
        Next X
    End If
    
    Packet = "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendMP(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(index) & SEP_CHAR & GetPlayerMP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendSP(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(index) & SEP_CHAR & GetPlayerSP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendStats(ByVal index As Long)
Dim Packet As String
    
    Packet = "PLAYERSTATSPACKET" & SEP_CHAR & GetPlayerStr(index) & SEP_CHAR & GetPlayerDEF(index) & SEP_CHAR & GetPlayerSPEED(index) & SEP_CHAR & GetPlayerMAGI(index) & SEP_CHAR & GetPlayerNextLevel(index) & SEP_CHAR & GetPlayerExp(index) & SEP_CHAR & GetPlayerLevel(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub Flash(ByVal index As Long, ByVal flashfile As String)
    Call SendDataTo(index, "flashevent" & SEP_CHAR & flashfile & SEP_CHAR & END_CHAR)
End Sub

Sub Prompt(ByVal index As Long, ByVal question As String, ByVal value As Long)
    Call SendDataTo(index, "prompt" & SEP_CHAR & question & SEP_CHAR & value & SEP_CHAR & END_CHAR)
End Sub

Sub PlaySound(ByVal index As Long, ByVal Sound As String)
    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & SEP_CHAR & END_CHAR)
End Sub

Sub SendPlayerData(ByVal index As Long)
Dim Packet As String

    ' Send index's player data to everyone including himself on th emap
    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & index & SEP_CHAR
    Packet = Packet & GetPlayerName(index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(index) & SEP_CHAR
    Packet = Packet & GetPlayerMap(index) & SEP_CHAR
    Packet = Packet & GetPlayerX(index) & SEP_CHAR
    Packet = Packet & GetPlayerY(index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(index) & SEP_CHAR
    Packet = Packet & GetPlayerLevel(index) & SEP_CHAR
    Packet = Packet & Player(index).InParty & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendPlayerQuete(ByVal index As Long)
Dim Packet As String
Dim i As Long
Packet = "PLAYERQUETE" & SEP_CHAR
Packet = Packet & Player(index).Char(Player(index).CharNum).QueteEnCour & SEP_CHAR
Packet = Packet & Player(index).Char(Player(index).CharNum).Quetep.data1 & SEP_CHAR
Packet = Packet & Player(index).Char(Player(index).CharNum).Quetep.data2 & SEP_CHAR
Packet = Packet & Player(index).Char(Player(index).CharNum).Quetep.data3 & SEP_CHAR
Packet = Packet & Player(index).Char(Player(index).CharNum).Quetep.String1 & SEP_CHAR

For i = 1 To 15
    Packet = Packet & Player(index).Char(Player(index).CharNum).Quetep.indexe(i).data1 & SEP_CHAR
    Packet = Packet & Player(index).Char(Player(index).CharNum).Quetep.indexe(i).data2 & SEP_CHAR
    Packet = Packet & Player(index).Char(Player(index).CharNum).Quetep.indexe(i).data3 & SEP_CHAR
    Packet = Packet & Player(index).Char(Player(index).CharNum).Quetep.indexe(i).String1 & SEP_CHAR
Next i
Packet = Packet & END_CHAR
Call SendDataTo(index, Packet)
End Sub

Sub SendPlayerMetier(ByVal index As Long)
Dim Packet As String

    ' Send index's player data to everyone including himself on th emap
    Packet = "PLAYERMETIER" & SEP_CHAR
    Packet = Packet & index & SEP_CHAR
    Packet = Packet & Player(index).Char(Player(index).CharNum).metier & SEP_CHAR
    Packet = Packet & Player(index).Char(Player(index).CharNum).MetierLvl & SEP_CHAR
    Packet = Packet & Player(index).Char(Player(index).CharNum).MetierExp & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendDataTo(ByVal index As Long, ByVal data As String)
Dim i As Long, n As Long, startc As Long
    If IsConnected(index) Then frmServer.Socket(index).SendData data: DoEvents
End Sub

Sub SendDataToAll(ByVal data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then Call SendDataTo(i, data)
    Next i
End Sub

Sub SendDataToAllBut(ByVal index As Long, ByVal data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index Then Call SendDataTo(i, data)
    Next i
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then If GetPlayerMap(i) = MapNum Then Call SendDataTo(i, data)
    Next i
End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal MapNum As Long, ByVal data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then If GetPlayerMap(i) = MapNum And i <> index Then Call SendDataTo(i, data)
    Next i
End Sub

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    Player(index).Char(Player(index).CharNum).Name = Name
End Sub

Function GetPlayerCharNum(ByVal index As Long) As Long
    GetPlayerCharNum = Player(index).CharNum
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

Function FindOpenInvSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
Dim i As Long
    
    FindOpenInvSlot = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function
           
<<<<<<< HEAD
    If item(ItemNum).type = ITEM_TYPE_CURRENCY Or item(ItemNum).Empilable <> 0 Then
=======
    If item(ItemNum).Type = ITEM_TYPE_CURRENCY Or item(ItemNum).Empilable <> 0 Then
>>>>>>> 78e891abe783f077a019bb7ae3b3f68bd63dddde
        ' If currency then check to see if they already have an guildSoloView of the item and add it to that
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If
        Next i
    End If
        
    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, i) <= 0 Then
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
        If MapItem(MapNum, i).num = 0 Then FindOpenMapItemSlot = i: Exit Function
    Next i
End Function

Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
Dim Packet As String
Dim OldMap As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
        
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)
    Call SendLeaveMap(index, OldMap)
    
    Call SetPlayerMap(index, MapNum)
    Call SetPlayerX(index, X)
    Call SetPlayerY(index, Y)
                
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then PlayersOnMap(OldMap) = NO
        
    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES

    Player(index).GettingMap = YES
    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "warp" & SEP_CHAR & END_CHAR)
    Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & END_CHAR)
    
    Call SendInventory(index)
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

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)
    If index > 0 Then
        If IsPlaying(index) Then Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " a été déconnecté pour : (" & Reason & ")", White)
        Call AlertMsg(index, "Tu as perdu ta connexion avec " & GAME_NAME & "." & Reason)
        Call AffIBMsg(index, "ATTENTION : Détection d'une tentative de hacking!!(Raison : " & Reason & " Login : " & GetPlayerLogin(index) & " perso : " & GetPlayerName(index) & ").", BrightRed, True)
    End If
End Sub

Sub BattleMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Long, ByVal Side As Long)
    Call SendDataTo(index, "damagedisplay" & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR)
End Sub

Public Sub ContrOnOff(ByVal index As Long)
Dim Packet As String

Packet = "CONOFF" & SEP_CHAR & END_CHAR

Call SendDataTo(index, Packet)
End Sub

Function Rand(ByVal High As Long, ByVal Low As Long)
Randomize
High = High + 1
Do Until Rand >= Low
    Randomize
    Rand = Int(Rnd * High)
Loop
End Function

Sub SetPlayerGuild(ByVal index As Long, ByVal Guild As String)
Player(index).Char(Player(index).CharNum).Guild = Guild
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

Sub KillPlayer(ByVal index As Long)
    Player(index).Char(Player(index).CharNum).HP = 0
    
    If GetPlayerHP(index) > GetPlayerMaxHP(index) Then Player(index).Char(Player(index).CharNum).HP = GetPlayerMaxHP(index)
    If GetPlayerHP(index) < 0 Then Player(index).Char(Player(index).CharNum).HP = 0
    Call SendStats(index)
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

Sub Lancer(ByVal index As Long, ByVal lien As String)
Dim Packet As String
    Packet = "LANCE" & SEP_CHAR & lien & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SetPlayerItem(index, item)
Dim n
n = 1
Do
   If GetPlayerInvItemNum(index, n) = 0 Then
      Call SetPlayerInvItemNum(index, n, item)
      Call SendInventoryUpdate(index, n)
      Exit Do
   End If
   n = n + 1
Loop Until n > 24
End Sub

Sub SetPlayerItemValue(index, item, value)
Dim n
n = 1
Do
   If GetPlayerInvItemNum(index, n) = item Then
      Call SetPlayerInvItemValue(index, n, value)
      Call SendInventoryUpdate(index, n)
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

Function GetPlayerHPRegen(ByVal index As Long)
Dim i As Long
    
    GetPlayerHPRegen = 0
    
    If Val(GetVar(App.Path & "\Data.ini", "CONFIG", "HPRegen")) >= 1 Then
        ' Prevent subscript out of range
        If Not IsPlaying(index) Or index <= 0 Or index > MAX_PLAYERS Then GetPlayerHPRegen = 0: Exit Function
        
        i = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "HPRegen")) '(GetPlayerDEF(Index) \ 2)
        If i < 2 Then i = 2
        
        GetPlayerHPRegen = i
    End If
End Function

Function GetPlayerMPRegen(ByVal index As Long)
Dim i As Long
    
    GetPlayerMPRegen = 0
    
    If Val(GetVar(App.Path & "\Data.ini", "CONFIG", "MPRegen")) >= 1 Then
        ' Prevent subscript out of range
        If Not IsPlaying(index) Or index <= 0 Or index > MAX_PLAYERS Then GetPlayerMPRegen = 0: Exit Function
        
        i = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "MPRegen")) '(GetPlayerMAGI(Index) \ 2)
        If i < 2 Then i = 2
        
        GetPlayerMPRegen = i
    End If
End Function

Function GetPlayerSPRegen(ByVal index As Long)
Dim i As Long
    
    GetPlayerSPRegen = 0
    
    If Val(GetVar(App.Path & "\Data.ini", "CONFIG", "SPRegen")) >= 1 Then
        ' Prevent subscript out of range
        If Not IsPlaying(index) Or index <= 0 Or index > MAX_PLAYERS Then GetPlayerSPRegen = 0: Exit Function
                
        i = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "SPRegen")) '(GetPlayerSPEED(Index) \ 2)
        If i < 2 Then i = 2
        
        GetPlayerSPRegen = i
    End If
End Function

Sub SuprObjet(ByVal index As Long, ByVal ItemNum As Long)
Dim i As Long

For i = 1 To MAX_INV
    If GetPlayerInvItemNum(index, i) = ItemNum Then
        If GetPlayerInvItemValue(index, i) > 1 Then
            Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - 1)
        Else
            Call SetPlayerInvItemNum(index, i, 0)
            Call SetPlayerInvItemValue(index, i, 0)
            Call SetPlayerInvItemDur(index, i, 0)
        End If
        Exit Sub
    End If
Next i

End Sub

Sub AjoutObjet(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long

i = FindOpenInvSlot(index, ItemNum)
If i = 0 Then Call PlayerMsg(index, "Ton inventaire est plein", Red): Exit Sub
Call SetPlayerInvItemNum(index, i, ItemNum)
Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)
If (item(ItemNum).type >= ITEM_TYPE_WEAPON) And (item(ItemNum).type <= ITEM_TYPE_SHIELD) Then Call SetPlayerInvItemDur(index, i, item(ItemNum).data1) Else Call SetPlayerInvItemDur(index, i, 0)
Call SendInventory(index)
End Sub

Sub ChangeClasse(ByVal index As Long, ByVal Classes As Long)

If Val(Classes) < 0 Or Val(Classes) > Max_Classes Then Exit Sub

<<<<<<< HEAD
If Player(index).Char(Player(index).CharNum).Sex = 0 Then
    If GetPlayerSprite(index) <> Classe(GetPlayerClass(index)).MaleSprite Then
        Call SetPlayerSprite(index, Classe(Classes).MaleSprite)
=======
If Player(Index).Char(Player(Index).CharNum).Sex = 0 Then
    If GetPlayerSprite(Index) <> Classe(GetPlayerClass(Index)).MaleSprite Then
        Call SetPlayerSprite(Index, Classe(Classes).MaleSprite)
>>>>>>> 78e891abe783f077a019bb7ae3b3f68bd63dddde
    End If
Else
    If GetPlayerSprite(index) <> Classe(GetPlayerClass(index)).FemaleSprite Then
        Call SetPlayerSprite(index, Classe(Classes).FemaleSprite)
    End If
End If

Call SetPlayerStr(index, (Player(index).Char(Player(index).CharNum).STR - Classe(GetPlayerClass(index)).STR))
Call SetPlayerDEF(index, (Player(index).Char(Player(index).CharNum).def - Classe(GetPlayerClass(index)).def))
Call SetPlayerMAGI(index, (Player(index).Char(Player(index).CharNum).magi - Classe(GetPlayerClass(index)).magi))
Call SetPlayerSPEED(index, (Player(index).Char(Player(index).CharNum).Speed - Classe(GetPlayerClass(index)).Speed))

Call SetPlayerClass(index, Classes)

Call SetPlayerStr(index, (Player(index).Char(Player(index).CharNum).STR + Classe(GetPlayerClass(index)).STR + GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "CLASSCHANGE", "AddStr")))
Call SetPlayerDEF(index, (Player(index).Char(Player(index).CharNum).def + Classe(GetPlayerClass(index)).def + GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "CLASSCHANGE", "AddDef")))
Call SetPlayerMAGI(index, (Player(index).Char(Player(index).CharNum).magi + Classe(GetPlayerClass(index)).magi + GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "CLASSCHANGE", "AddMagi")))
Call SetPlayerSPEED(index, (Player(index).Char(Player(index).CharNum).Speed + Classe(GetPlayerClass(index)).Speed + GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "CLASSCHANGE", "AddSpeed")))


Call PlayerMsg(index, "Ta nouvelle classe est " & Trim$(Classe(GetPlayerClass(index)).Name) & "!", BrightGreen)

Call SendStats(index)
Call SendHP(index)
Call SendMP(index)
Call SendSP(index)
Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & END_CHAR)
End Sub

Sub SauvJoueur(ByVal index As Long)
    Call SavePlayer(index)
End Sub

Function CibleJoueur(ByVal index As Long) As Long
CibleJoueur = Player(index).Target
End Function

Function TypeCibleJoueur(ByVal index As Long) As Byte
TypeCibleJoueur = Player(index).TargetType
End Function

Function ASort(ByVal index As Long, ByVal SpellNum As Long) As Long
Dim i As Long

    ASort = 0
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) = SpellNum Then ASort = i: Exit Function
    Next i
End Function

Function AObjet(ByVal index As Long, ByVal ItemNum As Long) As Long
Dim i As Long
    
    AObjet = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function
        
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
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

Sub Bloque(ByVal index As Long)
If GetPlayerDir(index) = DIR_UP Then
    Call SetPlayerY(index, GetPlayerY(index) + 1)
ElseIf GetPlayerDir(index) = DIR_DOWN Then
    Call SetPlayerY(index, GetPlayerY(index) - 1)
ElseIf GetPlayerDir(index) = DIR_RIGHT Then
    Call SetPlayerX(index, GetPlayerX(index) - 1)
ElseIf GetPlayerDir(index) = DIR_LEFT Then
    Call SetPlayerX(index, GetPlayerX(index) + 1)
End If
Call SendPlayerData(index)
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

Public Sub PlayerAnimStart(ByVal index As Long, ByVal Anim As Long)
Dim Packet As String

    If Not IsPlaying(index) Then Exit Sub

    Packet = "PLAYERANIMSTART" & SEP_CHAR & index & SEP_CHAR & Anim & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Public Sub PlayerAnimStop(ByVal index As Long)
Dim Packet As String

    If Not IsPlaying(index) Then Exit Sub

    Packet = "PLAYERANIMSTOP" & SEP_CHAR & index & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Public Sub PlayerAnim(ByVal index As Long, ByVal Anim As Long, ByVal ms As Long)
Dim Packet As String

    If Not IsPlaying(index) Then Exit Sub

    Packet = "PLAYERANIM" & SEP_CHAR & index & SEP_CHAR & Anim & SEP_CHAR & ms & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Public Sub PlayerAnimRetour(ByVal index As Long, ByVal Anim As Long, ByVal ms As Long, ByVal script As Long)
Dim Packet As String

    If Not IsPlaying(index) Then Exit Sub

    Packet = "PLAYERANIMRT" & SEP_CHAR & index & SEP_CHAR & Anim & SEP_CHAR & ms & SEP_CHAR & script & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

