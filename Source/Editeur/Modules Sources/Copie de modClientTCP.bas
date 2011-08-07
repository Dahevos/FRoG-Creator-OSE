Attribute VB_Name = "modClientTCP"
Option Explicit

Public ServerIP As String
Public PlayerBuffer As String
Public InGame As Boolean
Public TradePlayer As Long
Private MapNumS As Long

Sub TcpInit()
    Call EcrireEtat("Initialisation Tcp")
    
    SEP_CHAR = Chr(0)
    END_CHAR = Chr(237)
    PlayerBuffer = vbNullString
    
    Dim filename As String
    filename = App.Path & "\Config\Serveur.ini"

    frmMirage.Socket.RemoteHost = ReadINI("SERVER0", "IP", filename)
    frmMirage.Socket.RemotePort = Val(ReadINI("SERVER0", "Port", filename))
End Sub

Sub TcpDestroy()
    Call EcrireEtat("Destruction Tcp")
    
    frmMirage.Socket.Close
    
    If frmChars.Visible Then frmChars.Visible = False
    If frmpet.Visible Then frmpet.Visible = False
    
    If frmLogin.Visible Then frmLogin.Visible = False
    
    
End Sub

Sub IncomingData(ByVal DataLength As Long)
Dim Buffer As String
Dim Packet As String
Dim Top As String * 3
Dim Start As Long

    Call EcrireEtat("Traitement des données")
    frmMirage.Socket.GetData Buffer, vbString, DataLength
    PlayerBuffer = PlayerBuffer & Buffer
        
    Start = InStr(PlayerBuffer, END_CHAR)
    Do While Start > 0
        Packet = Mid(PlayerBuffer, 1, Start - 1)
        PlayerBuffer = Mid(PlayerBuffer, Start + 1, Len(PlayerBuffer))
        Start = InStr(PlayerBuffer, END_CHAR)
        If Len(Packet) > 0 Then
            Call HandleData(Packet)
        End If
    Loop
End Sub

Sub HandleData(ByVal data As String)
Dim Parse() As String
Dim name As String
Dim Password As String
Dim Sex As Long
Dim ClassNum As Long
Dim CharNum As Long
Dim Msg As String
Dim IPMask As String
Dim BanSlot As Long
Dim MsgTo As Long
Dim Dir As Long
Dim InvNum As Long
Dim Ammount As Long
Dim Damage As Long
Dim PointType As Long
Dim BanPlayer As Long
Dim Level As Long
Dim i As Long, n As Long, X As Long, Y As Long
Dim ShopNum As Long, GiveItem As Long, GiveValue As Long, GetItem As Long, GetValue As Long
Dim z As Long


    Call EcrireEtat("Analyse des données")
    ' Handle Data
    Parse = Split(data, SEP_CHAR)
    
    ' Add the data to the debug window if we are in debug mode
    If Trim(Command) = "-debug" Then
        If frmDebug.Visible = False Then frmDebug.Visible = True
        Call TextAdd(frmDebug.txtDebug, "((( Processed Packet " & Parse(0) & " )))", True)
    End If
    
    ' :::::::::::::::::::::::
    ' :: Get players stats ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "datacofr" Then
        Dim cofitemnum
        Dim cofitemval
        Dim slotcof
        
        cofitemnum = Val(Parse(1))
        cofitemval = Val(Parse(2))
        slotcof = Val(Parse(3))
        
        If cofitemnum = 0 Then GoTo jlu:
        If Item(cofitemnum).Type = 12 And cofitemval > 0 Then
        Else
        cofitemval = 1
        End If

        frmbank.Label3(slotcof).Caption = Item(cofitemnum).name
        frmbank.valcof(slotcof).Caption = "Valeur : " & cofitemval
        frmbank.numcof(slotcof).Caption = cofitemnum
jlu:
        Exit Sub
    End If
        
    If LCase(Parse(0)) = "picvalue" Then
        PIC_PL = Val(Parse(1))
        PIC_NPC1 = Val(Parse(2))
        PIC_NPC2 = Val(Parse(3))
        If PIC_NPC1 <= 0 Then PIC_NPC1 = 2
        If PIC_PL <= 0 Then PIC_PL = 64
        If PIC_NPC2 < 0 Then PIC_NPC2 = 32
        Call WriteINI("INFO", "PIC_PL", Val(PIC_PL), App.Path & "\config.ini")
        Call WriteINI("INFO", "PIC_NPC1", Val(PIC_NPC1), App.Path & "\config.ini")
        Call WriteINI("INFO", "PIC_NPC2", Val(PIC_NPC1), App.Path & "\config.ini")
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "maxinfo" Then
        GAME_NAME = Trim(Parse(1))
        MAX_PLAYERS = Val(Parse(2))
        MAX_ITEMS = Val(Parse(3))
        MAX_NPCS = Val(Parse(4))
        MAX_SHOPS = Val(Parse(5))
        MAX_SPELLS = Val(Parse(6))
        MAX_MAPS = Val(Parse(7))
        MAX_MAP_ITEMS = Val(Parse(8))
        MAX_MAPX = Val(Parse(9))
        MAX_MAPY = Val(Parse(10))
        MAX_EMOTICONS = Val(Parse(11))
        MAX_LEVEL = Val(Parse(12))
        MAX_QUETES = Val(Parse(13))
        'MAX_INV = Val(Parse(13))
'        MAX_PLAYER_SPELLS = Val(Parse(14))
        Call WriteINI("INFO", "GameName", GAME_NAME, App.Path & "\config.ini")
        Call WriteINI("INFO", "Maxplayers", Val(MAX_PLAYERS), App.Path & "\config.ini")
        Call WriteINI("INFO", "Maxitems", Val(MAX_ITEMS), App.Path & "\config.ini")
        Call WriteINI("INFO", "Maxnpcs", Val(MAX_NPCS), App.Path & "\config.ini")
        Call WriteINI("INFO", "Maxshops", Val(MAX_SHOPS), App.Path & "\config.ini")
        Call WriteINI("INFO", "Maxspells", Val(MAX_SPELLS), App.Path & "\config.ini")
        Call WriteINI("INFO", "Maxmaps", Val(MAX_MAPS), App.Path & "\config.ini")
        Call WriteINI("INFO", "Maxmapitems", Val(MAX_MAP_ITEMS), App.Path & "\config.ini")
        Call WriteINI("INFO", "Maxmapx", Val(MAX_MAPX), App.Path & "\config.ini")
        Call WriteINI("INFO", "Maxmapy", Val(MAX_MAPY), App.Path & "\config.ini")
        Call WriteINI("INFO", "Maxemots", Val(MAX_EMOTICONS), App.Path & "\config.ini")
        Call WriteINI("INFO", "Maxlevel", Val(MAX_LEVEL), App.Path & "\config.ini")
        Call WriteINI("INFO", "Maxinv", Val(MAX_INV), App.Path & "\config.ini")
        Call WriteINI("INFO", "Maxpspel", Val(MAX_PLAYER_SPELLS), App.Path & "\config.ini")
        Call WriteINI("INFO", "Maxquet", Val(MAX_QUETES), App.Path & "\config.ini")
        
        ReDim quete(1 To MAX_QUETES) As QueteRec
        ReDim Map(1 To MAX_MAPS) As MapRec
        ReDim MapAttributeNpc(1 To MAX_ATTRIBUTE_NPCS, 0 To MAX_MAPX, 0 To MAX_MAPY) As MapNpcRec
        ReDim SaveMapAttributeNpc(1 To MAX_ATTRIBUTE_NPCS, 0 To MAX_MAPX, 0 To MAX_MAPY) As MapNpcRec
        ReDim Player(1 To MAX_PLAYERS) As PlayerRec
        ReDim Item(1 To MAX_ITEMS) As ItemRec
        ReDim Npc(1 To MAX_NPCS) As NpcRec
        ReDim MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
        ReDim Shop(1 To MAX_SHOPS) As ShopRec
        ReDim Spell(1 To MAX_SPELLS) As SpellRec
        ReDim Bubble(1 To MAX_PLAYERS) As ChatBubble
        ReDim SaveMapItem(1 To MAX_MAP_ITEMS) As MapItemRec
        For i = 1 To MAX_MAPS
            ReDim Map(i).tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
            ReDim Map(i).tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
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
            Call LoadMap(i)
        Next i
        Call ChargerCartes
        
        frmMirage.Caption = "Editeur pour le jeu : " & Trim(GAME_NAME) & " Mettez votre souris sur un élément pour plus de détails."
        App.Title = GAME_NAME
 
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Multi-Serveur ::
    ' :::::::::::::::::::
    If LCase(Parse(0)) = "serverresults" Then
        frmServerChooser.lstServers.AddItem ReadINI("SERVER" & Val(Parse(1)), "Name", App.Path & "\Config\Serveur.ini") & " - Ouvert! (" & Val(Parse(2)) & "/" & Val(Parse(3)) & ")"
        CHECK_WAIT = False
        Exit Sub
    End If
        
    ' :::::::::::::::::::
    ' :: Npc hp packet ::
    ' :::::::::::::::::::
    If LCase(Parse(0)) = "npchp" Then
        n = Val(Parse(1))
 
        MapNpc(n).HP = Val(Parse(2))
        MapNpc(n).MaxHp = Val(Parse(3))
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Npc hp packet ::
    ' :::::::::::::::::::
    If LCase(Parse(0)) = "attributenpchp" Then
        n = Val(Parse(1))
 
        MapAttributeNpc(n, Val(Parse(4)), Val(Parse(5))).HP = Val(Parse(2))
        MapAttributeNpc(n, Val(Parse(4)), Val(Parse(5))).MaxHp = Val(Parse(3))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Alert message packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "alertmsg" Then
        frmMirage.Visible = False
        frmsplash.Visible = False
        frmMainMenu.Visible = True
        DoEvents

        Msg = Parse(1)
        Call MsgBox(Msg, vbOKOnly, GAME_NAME)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Plain message packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "plainmsg" Then
        frmsplash.Visible = False
        n = Val(Parse(2))
        
        
        If n = 3 Then frmMainMenu.Show
        
        If n = 5 Then frmChars.Show
        
        Msg = Parse(1)
        Call MsgBox(Msg, vbOKOnly, GAME_NAME)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: All characters packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "allchars" Then
        n = 1
        If frmasistantcreat.Visible = True And frmasistantcreat.AutoRedraw = False Then
        frmsplash.Visible = False
        
        frmChars.lstChars.Clear
        
        For i = 1 To MAX_CHARS
            name = Parse(n)
            Msg = Parse(n + 1)
            Level = Val(Parse(n + 2))
            
            If Trim(name) = vbNullString Then
                frmChars.lstChars.AddItem "Emplacement libre"
            Else
                frmChars.lstChars.AddItem name & " - niveaux " & Level & " - " & Msg
            End If
            
            n = n + 3
        Next i
        
        frmChars.lstChars.ListIndex = 0
        
       
        Exit Sub
        Else
        frmChars.Visible = True
        frmsplash.Visible = False
        frmChars.Visible = True
        frmChars.lstChars.Clear
        
        For i = 1 To MAX_CHARS
            name = Parse(n)
            Msg = Parse(n + 1)
            Level = Val(Parse(n + 2))
            
            If Trim(name) = vbNullString Then
                frmChars.lstChars.AddItem "Emplacement libre"
            Else
                frmChars.lstChars.AddItem name & " - niveaux " & Level & " - " & Msg
            End If
            
            n = n + 3
        Next i
        
        frmChars.lstChars.ListIndex = 0
        
       
        Exit Sub
        End If
    End If

    ' :::::::::::::::::::::::::::::::::
    ' :: Login was successful packet ::
    ' :::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "loginok" Then
        ' Now we can receive game data
        MyIndex = Val(Parse(1))
        
        frmsplash.Visible = True
        frmChars.Visible = False
        
        frmsplash.chrg.value = 50
        Call SetStatus("Réception des données en cours...")
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: New character classes data packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "newcharclasses" Then
        n = 1
        
        ' Max classes
        Max_Classes = Val(Parse(n))
        ReDim Class(0 To Max_Classes) As ClassRec
        
        n = n + 1

        For i = 0 To Max_Classes
            Class(i).name = Parse(n)
            
            Class(i).HP = Val(Parse(n + 1))
            Class(i).MP = Val(Parse(n + 2))
            Class(i).SP = Val(Parse(n + 3))
            
            Class(i).STR = Val(Parse(n + 4))
            Class(i).def = Val(Parse(n + 5))
            Class(i).speed = Val(Parse(n + 6))
            Class(i).magi = Val(Parse(n + 7))
            'Class(i).INTEL = Val(Parse(n + 8))
            Class(i).MaleSprite = Val(Parse(n + 8))
            Class(i).FemaleSprite = Val(Parse(n + 9))
            Class(i).Locked = Val(Parse(n + 10))
        
        n = n + 11
        Next i
        
        ' Used for if the player is creating a new character
        
        frmsplash.Visible = False

        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Classes data packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "classesdata" Then
        n = 1
        
        ' Max classes
        Max_Classes = Val(Parse(n))
        Call WriteINI("INFO", "Maxclasses", Val(Max_Classes), App.Path & "\config.ini")
        ReDim Class(0 To Max_Classes) As ClassRec
        
        n = n + 1
        
        For i = 0 To Max_Classes
            Class(i).name = Parse(n)
            
            Class(i).HP = Val(Parse(n + 1))
            Class(i).MP = Val(Parse(n + 2))
            Class(i).SP = Val(Parse(n + 3))
            
            Class(i).STR = Val(Parse(n + 4))
            Class(i).def = Val(Parse(n + 5))
            Class(i).speed = Val(Parse(n + 6))
            Class(i).magi = Val(Parse(n + 7))
            
            Class(i).Locked = Val(Parse(n + 8))
            
            n = n + 9
        Next i
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: In game packet ::
    ' ::::::::::::::::::::
    If LCase(Parse(0)) = "ingame" Then
        InGame = True
        If Player(MyIndex).Access < 1 Then
        MsgBox "Vous n'avez pas un acces suffisant pour éditer le jeu!!"
        MsgBox "Le logiciel va se terminer!!"
        Call GameDestroy
        End If
        Call GameInit
        Call GameLoop
        
        If Parse(1) = END_CHAR Then
            MsgBox ("here")
            End
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player inventory packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerinv" Then
        n = 2
        z = Val(Parse(1))
        
        For i = 1 To MAX_INV
            Call SetPlayerInvItemNum(z, i, Val(Parse(n)))
            Call SetPlayerInvItemValue(z, i, Val(Parse(n + 1)))
            Call SetPlayerInvItemDur(z, i, Val(Parse(n + 2)))
            
            n = n + 3
        Next i
        
        If z = MyIndex Then Call UpdateVisInv
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::::
    ' :: Player inventory update packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerinvupdate" Then
        n = Val(Parse(1))
        z = Val(Parse(2))
        
        Call SetPlayerInvItemNum(z, n, Val(Parse(3)))
        Call SetPlayerInvItemValue(z, n, Val(Parse(4)))
        Call SetPlayerInvItemDur(z, n, Val(Parse(5)))
        If z = MyIndex Then Call UpdateVisInv
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::
    ' :: Player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerworneq" Then
        z = Val(Parse(1))
        If z <= 0 Then Exit Sub
        
        Call SetPlayerArmorSlot(z, Val(Parse(2)))
        Call SetPlayerWeaponSlot(z, Val(Parse(3)))
        Call SetPlayerHelmetSlot(z, Val(Parse(4)))
        Call SetPlayerShieldSlot(z, Val(Parse(5)))
        
        If z = MyIndex Then
            Call UpdateVisInv
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Player points packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerpoints" Then
        Player(MyIndex).POINTS = Val(Parse(1))
        frmMirage.lblPoints.Caption = "Points :" & Val(Parse(1))
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::
    ' :: Player hp packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "playerhp" Then
        Player(MyIndex).MaxHp = Val(Parse(1))
        Call SetPlayerHP(MyIndex, Val(Parse(2)))
        If GetPlayerMaxHP(MyIndex) > 0 Then
            frmMirage.shpHP.FillColor = RGB(208, 11, 0)
            frmMirage.shpHP.Width = (((GetPlayerHP(MyIndex) / 126) / (GetPlayerMaxHP(MyIndex) / 126)) * 126)
            frmMirage.lblHP.Caption = GetPlayerHP(MyIndex) & " / " & GetPlayerMaxHP(MyIndex)
            'frmMirage.monnom.Caption = GetPlayerName(MyIndex)
            'frmMirage.maclasse.Caption = Trim(Class(GetPlayerClass(MyIndex)).name)
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Player mp packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "playermp" Then
        Player(MyIndex).MaxMP = Val(Parse(1))
        Call SetPlayerMP(MyIndex, Val(Parse(2)))
        If GetPlayerMaxMP(MyIndex) > 0 Then
            frmMirage.shpMP.FillColor = RGB(208, 11, 0)
            frmMirage.shpMP.Width = (((GetPlayerMP(MyIndex) / 126) / (GetPlayerMaxMP(MyIndex) / 126)) * 126)
            frmMirage.lblMP.Caption = GetPlayerMP(MyIndex) & " / " & GetPlayerMaxMP(MyIndex)
        End If
        Exit Sub
    End If
    
    ' speech bubble parse
    If (LCase(Parse(0)) = "mapmsg2") Then
        Bubble(Val(Parse(2))).Text = Parse(1)
        Bubble(Val(Parse(2))).Created = GetTickCount()
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Player sp packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "playersp" Then
       ' Player(MyIndex).MaxSP = Val(Parse(1))
        'Call SetPlayerSP(MyIndex, Val(Parse(2)))
        'If GetPlayerMaxSP(MyIndex) > 0 Then
            'frmMirage.lblSP.Caption = Int(GetPlayerSP(MyIndex) / GetPlayerMaxSP(MyIndex) * 100) & "%"
        'End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Player Stats Packet ::
    ' :::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playerstatspacket") Then
        Dim SubStr As Long, SubDef As Long, SubMagi As Long, SubSpeed As Long
        SubStr = 0
        SubDef = 0
        SubMagi = 0
        SubSpeed = 0
        On Error GoTo mont:
        If GetPlayerWeaponSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerArmorSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerShieldSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerHelmetSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddSpeed
        End If
mont:
        If SubStr > 0 Then
            frmMirage.lblSTR.Caption = "Force : " & Val(Parse(1)) - SubStr & " (+" & SubStr & ")"
        Else
            frmMirage.lblSTR.Caption = "Force : " & Val(Parse(1))
        End If
        If SubDef > 0 Then
            frmMirage.lblDEF.Caption = "Défense : " & Val(Parse(2)) - SubDef & " (+" & SubDef & ")"
        Else
            frmMirage.lblDEF.Caption = "Défense : " & Val(Parse(2))
        End If
        If SubMagi > 0 Then
            frmMirage.lblMAGI.Caption = "Magie : " & Val(Parse(4)) - SubMagi & " (+" & SubMagi & ")"
        Else
            frmMirage.lblMAGI.Caption = "Magie : " & Val(Parse(4))
        End If
        If SubSpeed > 0 Then
            frmMirage.lblSPEED.Caption = "Vitesse : " & Val(Parse(3)) - SubSpeed & " (+" & SubSpeed & ")"
        Else
            frmMirage.lblSPEED.Caption = "Vitesse : " & Val(Parse(3))
        End If
        frmMirage.lblEXP.Caption = Val(Parse(6)) & " / " & Val(Parse(5))
        
        frmMirage.shpTNL.Width = (((Val(Parse(6)) / 127) / (Val(Parse(5)) / 127)) * 127)
        frmMirage.lblLevel.Caption = "niveau : " & Val(Parse(7))
        frmMirage.monnom.Caption = Player(MyIndex).name
        frmMirage.maclasse.Caption = Class(Player(MyIndex).Class).name
        Player(MyIndex).Level = Val(Parse(7))
        Exit Sub
    End If
                

    ' ::::::::::::::::::::::::
    ' :: Player data packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerdata" Then
        i = Val(Parse(1))
        Call SetPlayerName(i, Parse(2))
        Call SetPlayerSprite(i, Val(Parse(3)))
        Call SetPlayerMap(i, Val(Parse(4)))
        Call SetPlayerX(i, Val(Parse(5)))
        Call SetPlayerY(i, Val(Parse(6)))
        Call SetPlayerDir(i, Val(Parse(7)))
        Call SetPlayerAccess(i, Val(Parse(8)))
        Call SetPlayerPK(i, Val(Parse(9)))
        Call SetPlayerGuild(i, Parse(10))
        Call SetPlayerGuildAccess(i, Val(Parse(11)))
        Call SetPlayerClass(i, Val(Parse(12)))
        Call SetPlayerLevel(i, Val(Parse(13)))

        ' Make sure they aren't walking
        Player(i).Moving = 0
        Player(i).XOffset = 0
        Player(i).YOffset = 0
        
        ' Check if the player is the client player, and if so reset directions
        If i = MyIndex Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Player movement packet ::
    ' ::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playermove") Then
        i = Val(Parse(1))
        X = Val(Parse(2))
        Y = Val(Parse(3))
        Dir = Val(Parse(4))
        n = Val(Parse(5))

        If Dir < DIR_UP Or Dir > DIR_RIGHT Then Exit Sub

        Call SetPlayerX(i, X)
        Call SetPlayerY(i, Y)
        Call SetPlayerDir(i, Dir)
                
        Player(i).XOffset = 0
        Player(i).YOffset = 0
        Player(i).Moving = n
        
        Select Case GetPlayerDir(i)
            Case DIR_UP
                Player(i).YOffset = PIC_Y
            Case DIR_DOWN
                Player(i).YOffset = PIC_Y * -1
            Case DIR_LEFT
                Player(i).XOffset = PIC_X
            Case DIR_RIGHT
                Player(i).XOffset = PIC_X * -1
        End Select
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Npc movement packet ::
    ' :::::::::::::::::::::::::
    If (LCase(Parse(0)) = "npcmove") Then
        i = Val(Parse(1))
        X = Val(Parse(2))
        Y = Val(Parse(3))
        Dir = Val(Parse(4))
        n = Val(Parse(5))

        MapNpc(i).X = X
        MapNpc(i).Y = Y
        MapNpc(i).Dir = Dir
        MapNpc(i).XOffset = 0
        MapNpc(i).YOffset = 0
        MapNpc(i).Moving = n
        
        Select Case MapNpc(i).Dir
            Case DIR_UP
                MapNpc(i).YOffset = PIC_Y
            Case DIR_DOWN
                MapNpc(i).YOffset = PIC_Y * -1
            Case DIR_LEFT
                MapNpc(i).XOffset = PIC_X
            Case DIR_RIGHT
                MapNpc(i).XOffset = PIC_X * -1
        End Select
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Npc movement packet ::
    ' :::::::::::::::::::::::::
    If (LCase(Parse(0)) = "attributenpcmove") Then
        i = Val(Parse(1))
        X = Val(Parse(2))
        Y = Val(Parse(3))
        Dir = Val(Parse(4))
        n = Val(Parse(5))

        MapAttributeNpc(i, Val(Parse(6)), Val(Parse(7))).X = X
        MapAttributeNpc(i, Val(Parse(6)), Val(Parse(7))).Y = Y
        MapAttributeNpc(i, Val(Parse(6)), Val(Parse(7))).Dir = Dir
        MapAttributeNpc(i, Val(Parse(6)), Val(Parse(7))).XOffset = 0
        MapAttributeNpc(i, Val(Parse(6)), Val(Parse(7))).YOffset = 0
        MapAttributeNpc(i, Val(Parse(6)), Val(Parse(7))).Moving = n
        
        Select Case MapAttributeNpc(i, Val(Parse(6)), Val(Parse(7))).Dir
            Case DIR_UP
                MapAttributeNpc(i, Val(Parse(6)), Val(Parse(7))).YOffset = PIC_Y
            Case DIR_DOWN
                MapAttributeNpc(i, Val(Parse(6)), Val(Parse(7))).YOffset = PIC_Y * -1
            Case DIR_LEFT
                MapAttributeNpc(i, Val(Parse(6)), Val(Parse(7))).XOffset = PIC_X
            Case DIR_RIGHT
                MapAttributeNpc(i, Val(Parse(6)), Val(Parse(7))).XOffset = PIC_X * -1
        End Select
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player direction packet ::
    ' :::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playerdir") Then
        i = Val(Parse(1))
        Dir = Val(Parse(2))
        
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then Exit Sub

        Call SetPlayerDir(i, Dir)
        
        Player(i).XOffset = 0
        Player(i).YOffset = 0
        Player(i).Moving = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: NPC direction packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "npcdir") Then
        i = Val(Parse(1))
        Dir = Val(Parse(2))
        MapNpc(i).Dir = Dir
        
        MapNpc(i).XOffset = 0
        MapNpc(i).YOffset = 0
        MapNpc(i).Moving = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: NPC direction packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "attributenpcdir") Then
        i = Val(Parse(1))
        Dir = Val(Parse(2))
        MapAttributeNpc(i, Val(Parse(3)), Val(Parse(4))).Dir = Dir
        
        MapAttributeNpc(i, Val(Parse(3)), Val(Parse(4))).XOffset = 0
        MapAttributeNpc(i, Val(Parse(3)), Val(Parse(4))).YOffset = 0
        MapAttributeNpc(i, Val(Parse(3)), Val(Parse(4))).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Player XY location packet ::
    ' :::::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playerxy") Then
        X = Val(Parse(1))
        Y = Val(Parse(2))
        
        Call SetPlayerX(MyIndex, X)
        Call SetPlayerY(MyIndex, Y)
        
        ' Make sure they aren't walking
        Player(MyIndex).Moving = 0
        Player(MyIndex).XOffset = 0
        Player(MyIndex).YOffset = 0
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "attack") Then
        i = Val(Parse(1))
        
        ' Set player to attacking
        Player(i).Attacking = 1
        Player(i).AttackTimer = GetTickCount
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' ::Player controle packet::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "conoff") Then
        
        If ConOff = False Then
            ConOff = True
        Else
            ConOff = False
        End If
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: NPC attack packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "npcattack") Then
        i = Val(Parse(1))
        
        ' Set player to attacking
        MapNpc(i).Attacking = 1
        MapNpc(i).AttackTimer = GetTickCount
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: NPC attack packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "attributenpcattack") Then
        i = Val(Parse(1))
        
        ' Set player to attacking
        MapAttributeNpc(i, Val(Parse(2)), Val(Parse(3))).Attacking = 1
        MapAttributeNpc(i, Val(Parse(2)), Val(Parse(3))).AttackTimer = GetTickCount
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Check for map packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "checkformap") Then
        ' Erase all players except self
        For i = 1 To MAX_PLAYERS
            If i <> MyIndex Then
                Call SetPlayerMap(i, 0)
            End If
        Next i
        
        ' Erase all temporary tile values
        Call ClearTempTile

        ' Get map num
        X = Val(Parse(1))
        
        ' Get revision
        Y = Val(Parse(2))
        
        If FileExiste("maps\map" & X & ".fcc") Then
            ' Check to see if the revisions match
            If GetMapRevision(X) = Y Then
                ' We do so we dont need the map
                
                ' Load the map
                'Call LoadMap(X)
                
                Call SendData("needmap" & SEP_CHAR & "no" & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        End If
        
        ' Either the revisions didn't match or we dont have the map, so we need it
        Call SendData("needmap" & SEP_CHAR & "yes" & SEP_CHAR & END_CHAR)
        'minmap = 0 'temp
        'frmMirage.Timer2.Enabled = True 'temp
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Map data packet ::
    ' :::::::::::::::::::::
    
    If LCase(Parse(0)) = "cartesave" Then
    Call Unload(frmmsg)
    'Call MsgBox("Carte" & Player(MyIndex).Map & " sauvegarder sur le serveur!!")
    Exit Sub
    End If
       
    If LCase(Parse(0)) = "mapdatas" Then
        
        n = 1
        MapNumS = Val(Parse(1))
        'Stop
        Map(Val(Parse(1))).name = Parse(n + 1)
        Map(Val(Parse(1))).Revision = Val(Parse(n + 2))
        Map(Val(Parse(1))).Moral = Val(Parse(n + 3))
        Map(Val(Parse(1))).Up = Val(Parse(n + 4))
        Map(Val(Parse(1))).Down = Val(Parse(n + 5))
        Map(Val(Parse(1))).Left = Val(Parse(n + 6))
        Map(Val(Parse(1))).Right = Val(Parse(n + 7))
        Map(Val(Parse(1))).music = Parse(n + 8)
        Map(Val(Parse(1))).BootMap = Val(Parse(n + 9))
        Map(Val(Parse(1))).BootX = Val(Parse(n + 10))
        Map(Val(Parse(1))).BootY = Val(Parse(n + 11))
        Map(Val(Parse(1))).Indoors = Val(Parse(n + 12))
        
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "maptiles" Then
        n = 1
        
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Map(MapNumS).tile(X, Y).Ground = Val(Parse(n))
                Map(MapNumS).tile(X, Y).Mask = Val(Parse(n + 1))
                Map(MapNumS).tile(X, Y).Anim = Val(Parse(n + 2))
                Map(MapNumS).tile(X, Y).Mask2 = Val(Parse(n + 3))
                Map(MapNumS).tile(X, Y).M2Anim = Val(Parse(n + 4))
                Map(MapNumS).tile(X, Y).Mask3 = Val(Parse(n + 32)) '<--
                Map(MapNumS).tile(X, Y).M3Anim = Val(Parse(n + 30)) '<--
                Map(MapNumS).tile(X, Y).Fringe = Val(Parse(n + 5))
                Map(MapNumS).tile(X, Y).FAnim = Val(Parse(n + 6))
                Map(MapNumS).tile(X, Y).Fringe2 = Val(Parse(n + 7))
                Map(MapNumS).tile(X, Y).F2Anim = Val(Parse(n + 8))
                Map(MapNumS).tile(X, Y).Fringe3 = Val(Parse(n + 26)) '<--
                Map(MapNumS).tile(X, Y).F3Anim = Val(Parse(n + 27)) '<--
                Map(MapNumS).tile(X, Y).Type = Val(Parse(n + 9))
                Map(MapNumS).tile(X, Y).Data1 = Val(Parse(n + 10))
                Map(MapNumS).tile(X, Y).Data2 = Val(Parse(n + 11))
                Map(MapNumS).tile(X, Y).Data3 = Val(Parse(n + 12))
                Map(MapNumS).tile(X, Y).String1 = Parse(n + 13)
                Map(MapNumS).tile(X, Y).String2 = Parse(n + 14)
                Map(MapNumS).tile(X, Y).String3 = Parse(n + 15)
                Map(MapNumS).tile(X, Y).Light = Val(Parse(n + 16))
                Map(MapNumS).tile(X, Y).GroundSet = Val(Parse(n + 17))
                Map(MapNumS).tile(X, Y).MaskSet = Val(Parse(n + 18))
                Map(MapNumS).tile(X, Y).AnimSet = Val(Parse(n + 19))
                Map(MapNumS).tile(X, Y).Mask2Set = Val(Parse(n + 20))
                Map(MapNumS).tile(X, Y).M2AnimSet = Val(Parse(n + 21))
                Map(MapNumS).tile(X, Y).Mask3Set = Val(Parse(n + 33)) '<--
                Map(MapNumS).tile(X, Y).M3AnimSet = Val(Parse(n + 31)) '<--
                Map(MapNumS).tile(X, Y).FringeSet = Val(Parse(n + 22))
                Map(MapNumS).tile(X, Y).FAnimSet = Val(Parse(n + 23))
                Map(MapNumS).tile(X, Y).Fringe2Set = Val(Parse(n + 24))
                Map(MapNumS).tile(X, Y).F2AnimSet = Val(Parse(n + 25))
                Map(MapNumS).tile(X, Y).Fringe3Set = Val(Parse(n + 28)) '<--
                Map(MapNumS).tile(X, Y).F3AnimSet = Val(Parse(n + 29)) '<--
                
                n = n + 34
            Next X
        Next Y
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "mapnpcs" Then
        n = 1
        For X = 1 To MAX_MAP_NPCS
            Map(MapNumS).Npc(X) = Val(Parse(n))
            n = n + 1
            Map(MapNumS).Npcs(X).X = Val(Parse(n))
            n = n + 1
            Map(MapNumS).Npcs(X).Y = Val(Parse(n))
            n = n + 1
            Map(MapNumS).Npcs(X).x1 = Val(Parse(n))
            n = n + 1
            Map(MapNumS).Npcs(X).y1 = Val(Parse(n))
            n = n + 1
            Map(MapNumS).Npcs(X).x2 = Val(Parse(n))
            n = n + 1
            Map(MapNumS).Npcs(X).y2 = Val(Parse(n))
            n = n + 1
            Map(MapNumS).Npcs(X).Hasardm = Val(Parse(n))
            n = n + 1
            Map(MapNumS).Npcs(X).Hasardp = Val(Parse(n))
            n = n + 1
            Map(MapNumS).Npcs(X).boucle = Val(Parse(n))
            n = n + 1
            Map(MapNumS).Npcs(X).Imobile = Val(Parse(n))
            n = n + 1
        Next X
                
        ' Save the map
        Call SaveLocalMap(MapNumS)
        
        ' Check if we get a map from someone else and if we were editing a map cancel it out
        If InEditor Then
            'InEditor = False
            frmMirage.Visible = False
            
            frmMirage.Show
            'frmMirage.picMapEditor.Visible = False
            
            If frmMapWarp.Visible Then
                Unload frmMapWarp
            End If
            
            If frmMapProperties.Visible Then
                Unload frmMapProperties
            End If
        End If
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: Map items data packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapitemdata" Then
        n = 1
        
        For i = 1 To MAX_MAP_ITEMS
            SaveMapItem(i).num = Val(Parse(n))
            SaveMapItem(i).value = Val(Parse(n + 1))
            SaveMapItem(i).Dur = Val(Parse(n + 2))
            SaveMapItem(i).X = Val(Parse(n + 3))
            SaveMapItem(i).Y = Val(Parse(n + 4))
            
            n = n + 5
        Next i
        
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::::
    ' :: Map npc data packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapnpcdata" Then
        n = 1
        
        For i = 1 To MAX_MAP_NPCS
            SaveMapNpc(i).num = Val(Parse(n))
            SaveMapNpc(i).X = Val(Parse(n + 1))
            SaveMapNpc(i).Y = Val(Parse(n + 2))
            SaveMapNpc(i).Dir = Val(Parse(n + 3))
            
            n = n + 4
        Next i
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Map npc data packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapattributenpcdata" Then
        n = 3
        
        X = Val(Parse(1))
        Y = Val(Parse(2))
        
        For i = 1 To MAX_ATTRIBUTE_NPCS
            'If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
                'If i <= Map(GetPlayerMap(MyIndex)).Tile(x, y).Data2 Then
                    SaveMapAttributeNpc(i, X, Y).num = Val(Parse(n))
                    SaveMapAttributeNpc(i, X, Y).X = Val(Parse(n + 1))
                    SaveMapAttributeNpc(i, X, Y).Y = Val(Parse(n + 2))
                    SaveMapAttributeNpc(i, X, Y).Dir = Val(Parse(n + 3))
    
                    n = n + 4
                'End If
            'End If
        Next i
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Map send completed packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapdone" Then
        'Map = SaveMap
        
        For i = 1 To MAX_MAP_ITEMS
            MapItem(i) = SaveMapItem(i)
        Next i
        
        For i = 1 To MAX_MAP_NPCS
            MapNpc(i) = SaveMapNpc(i)
        Next i
        
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                If Map(GetPlayerMap(MyIndex)).tile(X, Y).Type = TILE_TYPE_NPC_SPAWN Then
                    For i = 1 To MAX_ATTRIBUTE_NPCS
                        If i <= Map(GetPlayerMap(MyIndex)).tile(X, Y).Data2 Then
                            MapAttributeNpc(i, X, Y) = SaveMapAttributeNpc(i, X, Y)
                        End If
                    Next i
                End If
            Next X
        Next Y
        
        GettingMap = False
        
        ' Play music
        If Trim(Map(GetPlayerMap(MyIndex)).music) <> "None" Then
            Call PlayMidi(Trim(Map(GetPlayerMap(MyIndex)).music))
        Else
            Call StopMidi
        End If
        'frmMirage.Timer2.Enabled = False 'temp
        'MsgBox minmap 'temp
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "saymsg") Or (LCase(Parse(0)) = "broadcastmsg") Or (LCase(Parse(0)) = "globalmsg") Or (LCase(Parse(0)) = "playermsg") Or (LCase(Parse(0)) = "mapmsg") Or (LCase(Parse(0)) = "adminmsg") Then
        If Len(Parse(1)) > 76 Then
        For i = 0 To Int((Len(Parse(1)) / 76))
        If i > 0 Then
        Call AddText(Mid(Parse(1), 76 * i, 76), Val(Parse(2)))
        Else
        Call AddText(Mid(Parse(1), 1, 76), Val(Parse(2)))
        End If
        Next i
        Else
        Call AddText(Parse(1), Val(Parse(2)))
        End If
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "bank") Then
    frmbank.Show
    frmMirage.TxtQ2.Text = Map(GetPlayerMap(MyIndex)).tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).String1
    frmMirage.txtQ.Visible = True
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "qmsg") Then
    frmMirage.txtQ.Visible = True
    frmMirage.TxtQ2.Text = Parse(1)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Item spawn packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "spawnitem" Then
        n = Val(Parse(1))
        
        MapItem(n).num = Val(Parse(2))
        MapItem(n).value = Val(Parse(3))
        MapItem(n).Dur = Val(Parse(4))
        MapItem(n).X = Val(Parse(5))
        MapItem(n).Y = Val(Parse(6))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Item editor packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "itemeditor") Then
        InItemsEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_ITEMS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Item(i).name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update item packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updateitem") Then
        n = Val(Parse(1))
        
        ' Update the item
        Item(n).name = Parse(2)
        Item(n).Pic = Val(Parse(3))
        Item(n).Type = Val(Parse(4))
        Item(n).Data1 = Val(Parse(5))
        Item(n).Data2 = Val(Parse(6))
        Item(n).Data3 = Val(Parse(7))
        Item(n).StrReq = Val(Parse(8))
        Item(n).DefReq = Val(Parse(9))
        Item(n).SpeedReq = Val(Parse(10))
        Item(n).ClassReq = Val(Parse(11))
        Item(n).AccessReq = Val(Parse(12))
        
        Item(n).AddHP = Val(Parse(13))
        Item(n).AddMP = Val(Parse(14))
        Item(n).AddSP = Val(Parse(15))
        Item(n).AddStr = Val(Parse(16))
        Item(n).AddDef = Val(Parse(17))
        Item(n).AddMagi = Val(Parse(18))
        Item(n).AddSpeed = Val(Parse(19))
        Item(n).AddEXP = Val(Parse(20))
        Item(n).desc = Parse(21)
        Item(n).AttackSpeed = Val(Parse(22))
        Exit Sub
    End If
       
    ' ::::::::::::::::::::::
    ' :: Edit item packet :: <- Used for item editor admins only
    ' ::::::::::::::::::::::
    If (LCase(Parse(0)) = "edititem") Then
        n = Val(Parse(1))
        
        ' Update the item
        Item(n).name = Parse(2)
        Item(n).Pic = Val(Parse(3))
        Item(n).Type = Val(Parse(4))
        Item(n).Data1 = Val(Parse(5))
        Item(n).Data2 = Val(Parse(6))
        Item(n).Data3 = Val(Parse(7))
        Item(n).StrReq = Val(Parse(8))
        Item(n).DefReq = Val(Parse(9))
        Item(n).SpeedReq = Val(Parse(10))
        Item(n).ClassReq = Val(Parse(11))
        Item(n).AccessReq = Val(Parse(12))
        
        Item(n).AddHP = Val(Parse(13))
        Item(n).AddMP = Val(Parse(14))
        Item(n).AddSP = Val(Parse(15))
        Item(n).AddStr = Val(Parse(16))
        Item(n).AddDef = Val(Parse(17))
        Item(n).AddMagi = Val(Parse(18))
        Item(n).AddSpeed = Val(Parse(19))
        Item(n).AddEXP = Val(Parse(20))
        Item(n).desc = Parse(21)
        Item(n).AttackSpeed = Val(Parse(22))
        
        ' Initialize the item editor
        Call ItemEditorInit

        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "spawnnpc" Then
        n = Val(Parse(1))
        
        MapNpc(n).num = Val(Parse(2))
        MapNpc(n).X = Val(Parse(3))
        MapNpc(n).Y = Val(Parse(4))
        MapNpc(n).Dir = Val(Parse(5))
        MapNpc(n).Big = Val(Parse(6))
        
        ' Client use only
        MapNpc(n).XOffset = 0
        MapNpc(n).YOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "spawnattributenpc" Then
        n = Val(Parse(1))
        
        X = Val(Parse(7))
        Y = Val(Parse(8))
        
        MapAttributeNpc(n, X, Y).num = Val(Parse(2))
        MapAttributeNpc(n, X, Y).X = Val(Parse(3))
        MapAttributeNpc(n, X, Y).Y = Val(Parse(4))
        MapAttributeNpc(n, X, Y).Dir = Val(Parse(5))
        MapAttributeNpc(n, X, Y).Big = Val(Parse(6))
        
        ' Client use only
        MapAttributeNpc(n, X, Y).XOffset = 0
        MapAttributeNpc(n, X, Y).YOffset = 0
        MapAttributeNpc(n, X, Y).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "npcdead" Then
        n = Val(Parse(1))
        
        MapNpc(n).num = 0
        MapNpc(n).X = 0
        MapNpc(n).Y = 0
        MapNpc(n).Dir = 0
        
        ' Client use only
        MapNpc(n).XOffset = 0
        MapNpc(n).YOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "attributenpcdead" Then
        n = Val(Parse(1))
        
        MapAttributeNpc(n, Val(Parse(2)), Val(Parse(3))).num = 0
        MapAttributeNpc(n, Val(Parse(2)), Val(Parse(3))).X = 0
        MapAttributeNpc(n, Val(Parse(2)), Val(Parse(3))).Y = 0
        MapAttributeNpc(n, Val(Parse(2)), Val(Parse(3))).Dir = 0
        
        ' Client use only
        MapAttributeNpc(n, Val(Parse(2)), Val(Parse(3))).XOffset = 0
        MapAttributeNpc(n, Val(Parse(2)), Val(Parse(3))).YOffset = 0
        MapAttributeNpc(n, Val(Parse(2)), Val(Parse(3))).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Npc editor packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "npceditor") Then
        InNpcEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_NPCS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Npc(i).name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Update npc packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "updatenpc") Then
        n = Val(Parse(1))
        
        ' Update the item
        Npc(n).name = Parse(2)
        Npc(n).AttackSay = vbNullString
        Npc(n).Sprite = Val(Parse(3))
        Npc(n).SpawnSecs = 0
        Npc(n).Behavior = Val(Parse(7))
        Npc(n).Range = 0
        For i = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(i).Chance = 0
            Npc(n).ItemNPC(i).ItemNum = 0
            Npc(n).ItemNPC(i).ItemValue = 0
        Next i
        Npc(n).STR = 0
        Npc(n).def = 0
        Npc(n).speed = 0
        Npc(n).magi = 0
        Npc(n).Big = Val(Parse(4))
        Npc(n).MaxHp = Val(Parse(5))
        Npc(n).exp = 0
        Npc(n).quetenum = Val(Parse(6))
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit npc packet :: <- Used for item editor admins only
    ' :::::::::::::::::::::
    If (LCase(Parse(0)) = "editnpc") Then
        n = Val(Parse(1))
        
        ' Update the npc
        Npc(n).name = Parse(2)
        Npc(n).AttackSay = Parse(3)
        Npc(n).Sprite = Val(Parse(4))
        Npc(n).SpawnSecs = Val(Parse(5))
        Npc(n).Behavior = Val(Parse(6))
        Npc(n).Range = Val(Parse(7))
        Npc(n).STR = Val(Parse(8))
        Npc(n).def = Val(Parse(9))
        Npc(n).speed = Val(Parse(10))
        Npc(n).magi = Val(Parse(11))
        Npc(n).Big = Val(Parse(12))
        Npc(n).MaxHp = Val(Parse(13))
        Npc(n).exp = Val(Parse(14))
        Npc(n).SpawnTime = Val(Parse(15))
        Npc(n).quetenum = Val(Parse(16))
        
        z = 17
        For i = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(i).Chance = Val(Parse(z))
            Npc(n).ItemNPC(i).ItemNum = Val(Parse(z + 1))
            Npc(n).ItemNPC(i).ItemValue = Val(Parse(z + 2))
            z = z + 3
        Next i
        
        ' Initialize the npc editor
        Call NpcEditorInit

        Exit Sub
    End If

    ' ::::::::::::::::::::
    ' :: Map key packet ::
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "mapkey") Then
        X = Val(Parse(1))
        Y = Val(Parse(2))
        n = Val(Parse(3))
                
        TempTile(X, Y).DoorOpen = n
        
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit map packet ::
    ' :::::::::::::::::::::
    If (LCase(Parse(0)) = "editmap") Then
        Call EditorInit
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Shop editor packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "shopeditor") Then
        InShopEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_SHOPS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Shop(i).name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' ::Quetes editor packet::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "queteeditor") Then
        InQuetesEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_QUETES
            frmIndex.lstIndex.AddItem i & ": " & Trim(quete(i).nom)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update shop packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updateshop") Then
        n = Val(Parse(1))
        
        ' Update the shop name
        Shop(n).name = Parse(2)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Edit shop packet :: <- Used for shop editor admins only
    ' ::::::::::::::::::::::
    If (LCase(Parse(0)) = "editshop") Then
        ShopNum = Val(Parse(1))
        
        ' Update the shop
        Shop(ShopNum).name = Parse(2)
        Shop(ShopNum).JoinSay = Parse(3)
        Shop(ShopNum).LeaveSay = Parse(4)
        Shop(ShopNum).FixesItems = Val(Parse(5))
        
        n = 6
        For z = 1 To 6
            For i = 1 To MAX_TRADES
                
                GiveItem = Val(Parse(n))
                GiveValue = Val(Parse(n + 1))
                GetItem = Val(Parse(n + 2))
                GetValue = Val(Parse(n + 3))
                
                Shop(ShopNum).TradeItem(z).value(i).GiveItem = GiveItem
                Shop(ShopNum).TradeItem(z).value(i).GiveValue = GiveValue
                Shop(ShopNum).TradeItem(z).value(i).GetItem = GetItem
                Shop(ShopNum).TradeItem(z).value(i).GetValue = GetValue
                
                n = n + 4
            Next i
        Next z
        
        ' Initialize the shop editor
        Call ShopEditorInit

        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Spell editor packet ::
    ' :::::::::::::::::::::::::
    If (LCase(Parse(0)) = "spelleditor") Then
        InSpellEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_SPELLS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Spell(i).name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' ::    quete packet    ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "quetecour") Then
        Player(MyIndex).QueteEnCour = Val(Parse(1))
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "playerquete") Then
        Player(MyIndex).QueteEnCour = Val(Parse(1))
        Player(MyIndex).Quetep.Data1 = Val(Parse(2))
        Player(MyIndex).Quetep.Data2 = Val(Parse(3))
        Player(MyIndex).Quetep.Data3 = Val(Parse(4))
        Player(MyIndex).Quetep.String1 = Val(Parse(5))
        n = 5
        
        For i = 1 To 15
        n = n + 1
        Player(MyIndex).Quetep.indexe(i).Data1 = Val(Parse(n))
        n = n + 1
        Player(MyIndex).Quetep.indexe(i).Data2 = Val(Parse(n))
        n = n + 1
        Player(MyIndex).Quetep.indexe(i).Data3 = Val(Parse(n))
        n = n + 1
        Player(MyIndex).Quetep.indexe(i).String1 = Val(Parse(n))
        Next i
       
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "finquete") Then
        Call MsgBox("Bravo!! vous venez de finir la quête : " & Trim(quete(Player(MyIndex).QueteEnCour).nom) & " retourné voir celui qui vous la donné pour avoir vos récompenses")
        frmMirage.quetetimersec.Enabled = False
        frmMirage.tmpsquete.Visible = False
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "terminequete") Then
        Call ClearPlayerQuete(MyIndex)
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "tempsquete") Then
        If Val(Parse(1)) > 0 Then
        frmMirage.tmpsquete.Visible = True
        frmMirage.quetetimersec.Interval = 1000
        Seco = Val(Parse(1)) - (Int(Val(Parse(1)) / 60) * 60)
        Minu = Int(Val(Parse(1)) / 60)
        If Len(STR(Minu)) > 2 Then
        frmMirage.Minute.Caption = Minu & ":"
        Else
        frmMirage.Minute.Caption = "0" & Minu & ":"
        End If
        If Len(STR(Seco)) > 2 Then
        frmMirage.seconde.Caption = Seco
        Else
        frmMirage.seconde.Caption = "0" & Seco
        End If
        frmMirage.quetetimersec.Enabled = True
        Exit Sub
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update quete packet::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updatequete") Then
        n = Val(Parse(1))
        'la
        ' Update the quete
        quete(n).nom = Parse(2)
        quete(n).Data1 = Val(Parse(3))
        quete(n).Data2 = Val(Parse(4))
        quete(n).Data3 = Val(Parse(5))
        quete(n).description = Parse(6)
        quete(n).reponse = Parse(7)
        quete(n).String1 = Parse(8)
        quete(n).Temps = Val(Parse(9))
        quete(n).Type = Val(Parse(10))
        
        Dim l As Long
        i = 10
        For l = 1 To 15
        i = i + 1
        quete(n).indexe(l).Data1 = Val(Parse(i))
        i = i + 1
        quete(n).indexe(l).Data2 = Val(Parse(i))
        i = i + 1
        quete(n).indexe(l).Data3 = Val(Parse(i))
        i = i + 1
        quete(n).indexe(l).String1 = Parse(i)
        Next l
        quete(n).Recompence.exp = Val(Parse(i + 1))
        quete(n).Recompence.objn1 = Val(Parse(i + 2))
        quete(n).Recompence.objn2 = Val(Parse(i + 3))
        quete(n).Recompence.objn3 = Val(Parse(i + 4))
        quete(n).Recompence.objq1 = Val(Parse(i + 5))
        quete(n).Recompence.objq2 = Val(Parse(i + 6))
        quete(n).Recompence.objq3 = Val(Parse(i + 7))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update spell packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updatespell") Then
        n = Val(Parse(1))
        
        ' Update the spell name
        Spell(n).name = Parse(2)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit quetes packet:: <- Used for quetes editor admins only
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "editquetes") Then
        n = Val(Parse(1))
        
        ' Update the quetes
        quete(n).nom = Parse(2)
        quete(n).Data1 = Val(Parse(3))
        quete(n).Data2 = Val(Parse(4))
        quete(n).Data3 = Val(Parse(5))
        quete(n).description = Parse(6)
        quete(n).reponse = Parse(7)
        quete(n).String1 = Parse(8)
        quete(n).Temps = Val(Parse(9))
        quete(n).Type = Val(Parse(10))
        
        ' Initialize the quetes editor
        Call QuetesEditorInit

        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit spell packet :: <- Used for spell editor admins only
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "editspell") Then
        n = Val(Parse(1))
        
        ' Update the spell
        Spell(n).name = Parse(2)
        Spell(n).ClassReq = Val(Parse(3))
        Spell(n).LevelReq = Val(Parse(4))
        Spell(n).Type = Val(Parse(5))
        Spell(n).Data1 = Val(Parse(6))
        Spell(n).Data2 = Val(Parse(7))
        Spell(n).Data3 = Val(Parse(8))
        Spell(n).MPCost = Val(Parse(9))
        Spell(n).Sound = Val(Parse(10))
        Spell(n).Range = Val(Parse(11))
        Spell(n).SpellAnim = Val(Parse(12))
        Spell(n).SpellTime = Val(Parse(13))
        Spell(n).SpellDone = Val(Parse(14))
        Spell(n).AE = Val(Parse(15))
                        
        ' Initialize the spell editor
        Call SpellEditorInit

        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    If (LCase(Parse(0)) = "trade") Then
        
        ShopNum = Val(Parse(1))
        If Val(Parse(2)) = 1 Then
            frmTrade.picFixItems.Visible = True
        Else
            frmTrade.picFixItems.Visible = False
        End If
        
        n = 3
        For z = 1 To 6
            For i = 1 To MAX_TRADES
                GiveItem = Val(Parse(n))
                GiveValue = Val(Parse(n + 1))
                GetItem = Val(Parse(n + 2))
                GetValue = Val(Parse(n + 3))
                
                Trade(z).Items(i).ItemGetNum = GetItem
                Trade(z).Items(i).ItemGiveNum = GiveItem
                Trade(z).Items(i).ItemGetVal = GetValue
                Trade(z).Items(i).ItemGiveVal = GiveValue
                
                n = n + 4
            Next i
        Next z
        
        Dim xx As Long
        For xx = 1 To 6
            Trade(xx).Selected = NO
        Next xx
        
        Trade(1).Selected = YES
                    
        frmTrade.shopType.Top = frmTrade.Label1.Top
        frmTrade.shopType.Left = frmTrade.Label1.Left
        frmTrade.shopType.Height = frmTrade.Label1.Height
        frmTrade.shopType.Width = frmTrade.Label1.Width
        Trade(1).SelectedItem = 1
        
        Call ItemSelected(1, 1)
            
        frmTrade.Show vbModeless, frmMirage
        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If (LCase(Parse(0)) = "spells") Then
        
        frmMirage.picPlayerSpells.Visible = True
        frmMirage.lstSpells.Clear
        
        ' Put spells known in player record
        For i = 1 To MAX_PLAYER_SPELLS
            Player(MyIndex).Spell(i) = Val(Parse(i))
            If Player(MyIndex).Spell(i) <> 0 Then
                frmMirage.lstSpells.AddItem i & ": " & Trim(Spell(Player(MyIndex).Spell(i)).name)
            Else
                frmMirage.lstSpells.AddItem "<Emplacement de sort libre>"
            End If
        Next i
        
        frmMirage.lstSpells.ListIndex = 0
    End If

    ' ::::::::::::::::::::
    ' :: Weather packet ::
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "weather") Then
        If Val(Parse(1)) = WEATHER_RAINING And GameWeather <> WEATHER_RAINING Then
            Call AddText("La pluie commence a tomber!", BrightGreen)
        End If
        If Val(Parse(1)) = WEATHER_THUNDER And GameWeather <> WEATHER_THUNDER Then
            Call AddText("Le tonnerre commence a gronder!", BrightGreen)
        End If
        If Val(Parse(1)) = WEATHER_SNOWING And GameWeather <> WEATHER_SNOWING Then
            Call AddText("La neige commence a tomber", BrightGreen)
        End If
        
        If Val(Parse(1)) = WEATHER_NONE Then
            If GameWeather = WEATHER_RAINING Then
                Call AddText("La pluie se calme.", BrightGreen)
            ElseIf GameWeather = WEATHER_SNOWING Then
                Call AddText("La neige se calme.", BrightGreen)
            ElseIf GameWeather = WEATHER_THUNDER Then
                Call AddText("Les éclaires commence a disparaitre.", BrightGreen)
            End If
        End If
        GameWeather = Val(Parse(1))
        RainIntensity = Val(Parse(2))
        If MAX_RAINDROPS <> RainIntensity Then
            MAX_RAINDROPS = RainIntensity
            ReDim DropRain(1 To MAX_RAINDROPS) As DropRainRec
            ReDim DropSnow(1 To MAX_RAINDROPS) As DropRainRec
        End If
    End If

    ' :::::::::::::::::
    ' :: Time packet ::
    ' :::::::::::::::::
    If (LCase(Parse(0)) = "time") Then
        GameTime = Val(Parse(1))
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Get Online List ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "onlinelist" Then
    frmMirage.lstOnline.Clear
    
        n = 2
        z = Val(Parse(1))
        For X = n To (z + 1)
            frmMirage.lstOnline.AddItem Trim(Parse(n))
            n = n + 2
        Next X
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Blit Player Damage ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "blitplayerdmg" Then
        DmgDamage = Val(Parse(1))
        NPCWho = Val(Parse(2))
        DmgTime = GetTickCount
        iii = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Blit NPC Damage ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "blitnpcdmg" Then
        NPCDmgDamage = Val(Parse(1))
        NPCDmgTime = GetTickCount
        ii = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::
    ' :: Retrieve the player's inventory ::
    ' :::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "pptrading" Then
        frmPlayerTrade.Items1.Clear
        frmPlayerTrade.Items2.Clear
        For i = 1 To MAX_PLAYER_TRADES
            Trading(i).InvNum = 0
            Trading(i).InvName = vbNullString
            Trading2(i).InvNum = 0
            Trading2(i).InvName = vbNullString
            frmPlayerTrade.Items1.AddItem i & ": <Nothing>"
            frmPlayerTrade.Items2.AddItem i & ": <Nothing>"
        Next i
        
        frmPlayerTrade.Items1.ListIndex = 0
        
        Call UpdateTradeInventory
        frmPlayerTrade.Show vbModeless, frmMirage
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "qtrade" Then
        For i = 1 To MAX_PLAYER_TRADES
            Trading(i).InvNum = 0
            Trading(i).InvName = vbNullString
            Trading2(i).InvNum = 0
            Trading2(i).InvName = vbNullString
        Next i
        
        frmPlayerTrade.Command1.ForeColor = &H0&
        frmPlayerTrade.Command2.ForeColor = &H0&
        
        frmPlayerTrade.Visible = False
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "updatetradeitem" Then
            n = Val(Parse(1))
            
            Trading2(n).InvNum = Val(Parse(2))
            Trading2(n).InvName = Parse(3)
            
            If STR(Trading2(n).InvNum) <= 0 Then
                frmPlayerTrade.Items2.List(n - 1) = n & ": <Nothing>"
            Else
                frmPlayerTrade.Items2.List(n - 1) = n & ": " & Trim(Trading2(n).InvName)
            End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "trading" Then
        n = Val(Parse(1))
            If n = 0 Then frmPlayerTrade.Command2.ForeColor = &H0&
            If n = 1 Then frmPlayerTrade.Command2.ForeColor = &HFF00&
        Exit Sub
    End If
    
' :::::::::::::::::::::::::
' :: Chat System Packets ::
' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "ppchatting" Then
        frmPlayerChat.txtChat.Text = vbNullString
        frmPlayerChat.txtSay.Text = vbNullString
        frmPlayerChat.Label1.Caption = "Vous chattez avec : " & Trim(Player(Val(Parse(1))).name)

        frmPlayerChat.Show vbModeless, frmMirage
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "qchat" Then
        frmPlayerChat.txtChat.Text = vbNullString
        frmPlayerChat.txtSay.Text = vbNullString
        frmPlayerChat.Visible = False
        frmPlayerTrade.Command2.ForeColor = &H8000000F
        frmPlayerTrade.Command1.ForeColor = &H8000000F
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "sendchat" Then
        Dim s As String
  
        s = vbNewLine & GetPlayerName(Val(Parse(2))) & "> " & Trim(Parse(1))
        frmPlayerChat.txtChat.SelStart = Len(frmPlayerChat.txtChat.Text)
        frmPlayerChat.txtChat.SelColor = QBColor(Brown)
        frmPlayerChat.txtChat.SelText = s
        frmPlayerChat.txtChat.SelStart = Len(frmPlayerChat.txtChat.Text) - 1
        Exit Sub
    End If
' :::::::::::::::::::::::::::::
' :: END Chat System Packets ::
' :::::::::::::::::::::::::::::

    ' :::::::::::::::::::::::
    ' :: Play Sound Packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "sound" Then
        s = LCase(Parse(1))
        Select Case Trim(s)
            Case "attack"
                Call PlaySound("sword.wav")
            Case "critical"
                Call PlaySound("critical.wav")
            Case "miss"
                Call PlaySound("miss.wav")
            Case "key"
                Call PlaySound("key.wav")
            Case "magic"
                Call PlaySound("magic" & Val(Parse(2)) & ".wav")
            Case "warp"
                Call PlaySound("warp.wav")
            Case "pain"
                Call PlaySound("pain.wav")
            Case "soundattribute"
                Call PlaySound(Trim(Parse(2)))
        End Select
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: Sprite Change Confirmation Packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "spritechange" Then
        If Val(Parse(1)) = 1 Then
            i = MsgBox("Êtes-vous sur de vouloir acheter ce sprite?", 4, "Acheter un Sprite")
            If i = 6 Then
                Call SendData("buysprite" & SEP_CHAR & END_CHAR)
            End If
        Else
            Call SendData("buysprite" & SEP_CHAR & END_CHAR)
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::::
    ' :: Change Player Direction Packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "changedir" Then
        Player(Val(Parse(2))).Dir = Val(Parse(1))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Flash Movie Event Packet ::
    ' ::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "flashevent" Then
        If LCase(Mid(Trim(Parse(1)), 1, 7)) = "http://" Then
            WriteINI "CONFIG", "Music", 0, App.Path & "\Config\Account.ini"
            WriteINI "CONFIG", "Sound", 0, App.Path & "\Config\Account.ini"
            Call StopMidi
            Call StopSound
            frmFlash.Flash.LoadMovie 0, Trim(Parse(1))
            frmFlash.Flash.Play
            frmFlash.Check.Enabled = True
            frmFlash.Show vbModeless, frmMirage
        ElseIf FileExiste("Flashs\" & Trim(Parse(1))) = True Then
            WriteINI "CONFIG", "Music", 0, App.Path & "\Config\Account.ini"
            WriteINI "CONFIG", "Sound", 0, App.Path & "\Config\Account.ini"
            Call StopMidi
            Call StopSound
            frmFlash.Flash.LoadMovie 0, App.Path & "\Flashs\" & Trim(Parse(1))
            frmFlash.Flash.Play
            frmFlash.Check.Enabled = True
            frmFlash.Show vbModeless, frmMirage
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Prompt Packet ::
    ' :::::::::::::::::::
    If LCase(Parse(0)) = "prompt" Then
        i = MsgBox(Trim(Parse(1)), vbYesNo)
        Call SendData("prompt" & SEP_CHAR & i & SEP_CHAR & Val(Parse(2)) & SEP_CHAR & END_CHAR)
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Emoticon editor packet ::
    ' ::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "emoticoneditor") Then
        InEmoticonEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        For i = 0 To MAX_EMOTICONS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Emoticons(i).Command)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "updateemoticon") Then
        n = Val(Parse(1))
        
        Emoticons(n).Command = Parse(2)
        Emoticons(n).Pic = Val(Parse(3))
        Exit Sub
    End If

    If (LCase(Parse(0)) = "editemoticon") Then
        n = Val(Parse(1))

        Emoticons(n).Command = Parse(2)
        Emoticons(n).Pic = Val(Parse(3))
        
        Call EmoticonEditorInit
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "updateemoticon") Then
        n = Val(Parse(1))
        
        Emoticons(n).Command = Parse(2)
        Emoticons(n).Pic = Val(Parse(3))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Arrow editor packet ::
    ' ::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "arroweditor") Then
        InArrowEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        For i = 1 To MAX_ARROWS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Arrows(i).name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "updatearrow") Then
        n = Val(Parse(1))
        
        Arrows(n).name = Parse(2)
        Arrows(n).Pic = Val(Parse(3))
        Arrows(n).Range = Val(Parse(4))
        Exit Sub
    End If

    If (LCase(Parse(0)) = "editarrow") Then
        n = Val(Parse(1))

        Arrows(n).name = Parse(2)
        
        Call ArrowEditorInit
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "updatearrow") Then
        n = Val(Parse(1))
        
        Arrows(n).name = Parse(2)
        Arrows(n).Pic = Val(Parse(3))
        Arrows(n).Range = Val(Parse(4))
        Exit Sub
    End If

    If (LCase(Parse(0)) = "checkarrows") Then
        n = Val(Parse(1))
        z = Val(Parse(2))
        i = Val(Parse(3))
        
        For X = 1 To MAX_PLAYER_ARROWS
            If Player(n).Arrow(X).Arrow = 0 Then
                Player(n).Arrow(X).Arrow = 1
                Player(n).Arrow(X).ArrowNum = z
                Player(n).Arrow(X).ArrowAnim = Arrows(z).Pic
                Player(n).Arrow(X).ArrowTime = GetTickCount
                Player(n).Arrow(X).ArrowVarX = 0
                Player(n).Arrow(X).ArrowVarY = 0
                Player(n).Arrow(X).ArrowY = GetPlayerY(n)
                Player(n).Arrow(X).ArrowX = GetPlayerX(n)
                
                If i = DIR_DOWN Then
                    Player(n).Arrow(X).ArrowY = GetPlayerY(n) + 1
                    Player(n).Arrow(X).ArrowPosition = 0
                    If Player(n).Arrow(X).ArrowY - 1 > MAX_MAPY Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                If i = DIR_UP Then
                    Player(n).Arrow(X).ArrowY = GetPlayerY(n) - 1
                    Player(n).Arrow(X).ArrowPosition = 1
                    If Player(n).Arrow(X).ArrowY + 1 < 0 Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                If i = DIR_RIGHT Then
                    Player(n).Arrow(X).ArrowX = GetPlayerX(n) + 1
                    Player(n).Arrow(X).ArrowPosition = 2
                    If Player(n).Arrow(X).ArrowX - 1 > MAX_MAPX Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                If i = DIR_LEFT Then
                    Player(n).Arrow(X).ArrowX = GetPlayerX(n) - 1
                    Player(n).Arrow(X).ArrowPosition = 3
                    If Player(n).Arrow(X).ArrowX + 1 < 0 Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                Exit For
            End If
        Next X
        Exit Sub
    End If

    If (LCase(Parse(0)) = "checksprite") Then
        n = Val(Parse(1))
        
        Player(n).Sprite = Val(Parse(2))
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "mapreport") Then
        n = 1
        
        frmMirage.lstIndex.Clear
        For i = 1 To MAX_MAPS
            frmMirage.lstIndex.AddItem i & ": " & Trim(Parse(n))
            n = n + 1
        Next i
        
        
        Exit Sub
    End If
    
    ' :::::::::::::::::
    ' :: Time packet ::
    ' :::::::::::::::::
    If (LCase(Parse(0)) = "time") Then
        GameTime = Val(Parse(1))
        If GameTime = TIME_DAY Then
            Call AddText("Le jours ce leve sur ce royaume.", White)
        Else
            Call AddText("La nuit tombe dans ce royaume..", White)
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Spell anim packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "spellanim") Then
        Dim SpellNum As Long
        SpellNum = Val(Parse(1))
        
        Spell(SpellNum).SpellAnim = Val(Parse(2))
        Spell(SpellNum).SpellTime = Val(Parse(3))
        Spell(SpellNum).SpellDone = Val(Parse(4))
        
        Player(Val(Parse(5))).SpellNum = SpellNum
        
        For i = 1 To MAX_SPELL_ANIM
            If Player(Val(Parse(5))).SpellAnim(i).CastedSpell = NO Then
                Player(Val(Parse(5))).SpellAnim(i).SpellDone = 0
                Player(Val(Parse(5))).SpellAnim(i).SpellVar = 0
                Player(Val(Parse(5))).SpellAnim(i).SpellTime = GetTickCount
                Player(Val(Parse(5))).SpellAnim(i).TargetType = Val(Parse(6))
                Player(Val(Parse(5))).SpellAnim(i).Target = Val(Parse(7))
                Player(Val(Parse(5))).SpellAnim(i).CastedSpell = YES
                Exit For
            End If
        Next i
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "checkemoticons") Then
        n = Val(Parse(1))
        
        Player(n).EmoticonNum = Val(Parse(2))
        Player(n).EmoticonTime = GetTickCount
        Player(n).EmoticonVar = 0
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "levelup" Then
        Player(Val(Parse(1))).LevelUpT = GetTickCount
        Player(Val(Parse(1))).LevelUp = 1
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "damagedisplay" Then
        For i = 1 To MAX_BLT_LINE
            If Val(Parse(1)) = 0 Then
                If BattlePMsg(i).index <= 0 Then
                    BattlePMsg(i).index = 1
                    BattlePMsg(i).Msg = Parse(2)
                    BattlePMsg(i).Color = Val(Parse(3))
                    BattlePMsg(i).Time = GetTickCount
                    BattlePMsg(i).Done = 1
                    BattlePMsg(i).Y = 0
                    Exit Sub
                Else
                    BattlePMsg(i).Y = BattlePMsg(i).Y - 15
                End If
            Else
                If BattleMMsg(i).index <= 0 Then
                    BattleMMsg(i).index = 1
                    BattleMMsg(i).Msg = Parse(2)
                    BattleMMsg(i).Color = Val(Parse(3))
                    BattleMMsg(i).Time = GetTickCount
                    BattleMMsg(i).Done = 1
                    BattleMMsg(i).Y = 0
                    Exit Sub
                Else
                    BattleMMsg(i).Y = BattleMMsg(i).Y - 15
                End If
            End If
        Next i
        
        z = 1
        If Val(Parse(1)) = 0 Then
            For i = 1 To MAX_BLT_LINE
                If i < MAX_BLT_LINE Then
                    If BattlePMsg(i).Y < BattlePMsg(i + 1).Y Then z = i
                Else
                    If BattlePMsg(i).Y < BattlePMsg(1).Y Then z = i
                End If
            Next i
                        
            BattlePMsg(z).index = 1
            BattlePMsg(z).Msg = Parse(2)
            BattlePMsg(z).Color = Val(Parse(3))
            BattlePMsg(z).Time = GetTickCount
            BattlePMsg(z).Done = 1
            BattlePMsg(z).Y = 0
        Else
            For i = 1 To MAX_BLT_LINE
                If i < MAX_BLT_LINE Then
                    If BattleMMsg(i).Y < BattleMMsg(i + 1).Y Then z = i
                Else
                    If BattleMMsg(i).Y < BattleMMsg(1).Y Then z = i
                End If
            Next i
                        
            BattleMMsg(z).index = 1
            BattleMMsg(z).Msg = Parse(2)
            BattleMMsg(z).Color = Val(Parse(3))
            BattleMMsg(z).Time = GetTickCount
            BattleMMsg(z).Done = 1
            BattleMMsg(z).Y = 0
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "itembreak" Then
        ItemDur(Val(Parse(1))).Item = Val(Parse(2))
        ItemDur(Val(Parse(1))).Dur = Val(Parse(3))
        ItemDur(Val(Parse(1))).Done = 1
        Exit Sub
    End If
End Sub

Function ConnectToServer() As Boolean
Dim Wait As Long

    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    frmMirage.Socket.Close
    frmMirage.Socket.Connect
    
    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsConnected) And (GetTickCount <= Wait + 3000)
        DoEvents
    Loop
    
    If IsConnected Then
        ConnectToServer = True
    Else
        ConnectToServer = False
    End If
End Function

Function IsConnected() As Boolean
    If frmMirage.Socket.State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function

Function IsPlaying(ByVal index As Long) As Boolean
    If GetPlayerName(index) <> "" Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If
End Function

Sub SendData(ByVal data As String)
    Call EcrireEtat("Envoie des données")
    If IsConnected Then
        frmMirage.Socket.SendData data
        DoEvents
    End If
End Sub

Sub SendNewAccount(ByVal name As String, ByVal Password As String)
Dim Packet As String

    Call EcrireEtat("Création du compte")
    Packet = "newfaccountied" & SEP_CHAR & Trim(name) & SEP_CHAR & Trim(Password) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDelAccount(ByVal name As String, ByVal Password As String)
Dim Packet As String
    
    Call EcrireEtat("Supression du compte")
    Packet = "delimaccounted" & SEP_CHAR & Trim(name) & SEP_CHAR & Trim(Password) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendLogin(ByVal name As String, ByVal Password As String)
Dim Packet As String

    Call EcrireEtat("Envoie du login")
    Packet = "logination" & SEP_CHAR & Trim(name) & SEP_CHAR & Trim(Password) & SEP_CHAR & App.Major & SEP_CHAR & App.Minor & SEP_CHAR & App.Revision & SEP_CHAR & SEC_CODE1 & SEP_CHAR & SEC_CODE2 & SEP_CHAR & SEC_CODE3 & SEP_CHAR & SEC_CODE4 & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendAddChar(ByVal name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Slot As Long)
Dim Packet As String
    
    Call EcrireEtat("Création du personnage")
    Packet = "addachara" & SEP_CHAR & Trim(name) & SEP_CHAR & Sex & SEP_CHAR & ClassNum & SEP_CHAR & Slot & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDelChar(ByVal Slot As Long)
Dim Packet As String
    
    Call EcrireEtat("Supression du personnage")
    Packet = "delimbocharu" & SEP_CHAR & Slot & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendGetClasses()
Dim Packet As String

    Call EcrireEtat("Demande des classes")
    Packet = "gatglasses" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendUseChar(ByVal CharSlot As Long)
Dim Packet As String

    Call EcrireEtat("Utilisation d'un personnage")
    Packet = "usagakarim" & SEP_CHAR & CharSlot & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SayMsg(ByVal Text As String)
Dim Packet As String

    Call EcrireEtat("Envoie d'un message(parler)")
    Packet = "saymsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub GlobalMsg(ByVal Text As String)
Dim Packet As String

    Call EcrireEtat("Envoie d'un message a tout le monde")
    Packet = "globalmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub BroadcastMsg(ByVal Text As String)
Dim Packet As String

    Call EcrireEtat("Envoie d'un message(spéciale)")
    Packet = "broadcastmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub EmoteMsg(ByVal Text As String)
Dim Packet As String

    Call EcrireEtat("Envoie d'un émoticons")
    Packet = "emotemsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub GuildeMsg(ByVal Text As String)
Dim Packet As String

   Call EcrireEtat("Envoie d'un message a la guilde")
   Packet = "guildemsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
   Call SendData(Packet)
End Sub

Sub MapMsg(ByVal Text As String)
Dim Packet As String

    Call EcrireEtat("Envoie d'un message a tout ceux de la carte")
    Packet = "mapmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
Dim Packet As String

    Call EcrireEtat("Envoie d'un message à un joueur")
    Packet = "playermsg" & SEP_CHAR & MsgTo & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub AdminMsg(ByVal Text As String)
Dim Packet As String

    Call EcrireEtat("Envoie d'un message au admins")
    Packet = "adminmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerMove()
Dim Packet As String

    Call EcrireEtat("Envoie du déplacement du joueur")
    Packet = "playermove" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & Player(MyIndex).Moving & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerDir()
Dim Packet As String

    Call EcrireEtat("Envoie de la direction du joueur")
    Packet = "playerdir" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerRequestNewMap()
Dim Packet As String
    
    Call EcrireEtat("Demande de réception d'une carte")
    Packet = "requestnewmap" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendMap()
Dim Packet As String, P1 As String, P2 As String
Dim X As Long
Dim Y As Long

    Call EcrireEtat("Envoie/Sauvegarde d'une carte")
    Packet = "MAPDATA" & SEP_CHAR & GetPlayerMap(MyIndex) & SEP_CHAR & Trim(Map(GetPlayerMap(MyIndex)).name) & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Revision & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Moral & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Up & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Down & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Left & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Right & SEP_CHAR & Map(GetPlayerMap(MyIndex)).music & SEP_CHAR & Map(GetPlayerMap(MyIndex)).BootMap & SEP_CHAR & Map(GetPlayerMap(MyIndex)).BootX & SEP_CHAR & Map(GetPlayerMap(MyIndex)).BootY & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Indoors & SEP_CHAR
    
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            With Map(GetPlayerMap(MyIndex)).tile(X, Y)
                Packet = Packet & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR & .Light & SEP_CHAR
                Packet = Packet & .GroundSet & SEP_CHAR & .MaskSet & SEP_CHAR & .AnimSet & SEP_CHAR & .Mask2Set & SEP_CHAR & .M2AnimSet & SEP_CHAR & .FringeSet & SEP_CHAR & .FAnimSet & SEP_CHAR & .Fringe2Set & SEP_CHAR & .F2AnimSet & SEP_CHAR & .Fringe3 & SEP_CHAR & .F3Anim & SEP_CHAR & .Fringe3Set & SEP_CHAR & .F3AnimSet & SEP_CHAR & .M3Anim & SEP_CHAR & .M3AnimSet & SEP_CHAR & .Mask3 & SEP_CHAR & .Mask3Set & SEP_CHAR  '<--2
            End With
        Next X
    Next Y
    
    For X = 1 To MAX_MAP_NPCS
        Packet = Packet & Map(GetPlayerMap(MyIndex)).Npc(X) & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(MyIndex)).Npcs(X).X & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(MyIndex)).Npcs(X).Y & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(MyIndex)).Npcs(X).x1 & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(MyIndex)).Npcs(X).y1 & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(MyIndex)).Npcs(X).x2 & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(MyIndex)).Npcs(X).y2 & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(MyIndex)).Npcs(X).Hasardm & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(MyIndex)).Npcs(X).Hasardp & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(MyIndex)).Npcs(X).boucle & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(MyIndex)).Npcs(X).Imobile & SEP_CHAR
    Next X
    
    Packet = Packet & END_CHAR
    
    X = Int(Len(Packet) / 2)
    P1 = Mid(Packet, 1, X)
    P2 = Mid(Packet, X + 1, Len(Packet) - X)
    Call SendData(Packet)
End Sub

Sub WarpMeTo(ByVal name As String)
Dim Packet As String

    Call EcrireEtat("Téléportation j'usqua " & name)
    Packet = "WARPMETO" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub WarpToMe(ByVal name As String)
Dim Packet As String

    Call EcrireEtat("Téléporattion de " & name & " vers le joueur")
    Packet = "WARPTOME" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub WarpTo(ByVal MapNum As Long)
Dim Packet As String
    
    Call EcrireEtat("Téléportation à la carte " & MapNum)
    Packet = "WARPTO" & SEP_CHAR & MapNum & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetAccess(ByVal name As String, ByVal Access As Byte)
Dim Packet As String

    Call EcrireEtat("Changement d'acces de " & name & ". Nouvelle acces : " & Access)
    Packet = "SETACCESS" & SEP_CHAR & name & SEP_CHAR & Access & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetSprite(ByVal SpriteNum As Long)
Dim Packet As String

    Call EcrireEtat("Changement de sprite/skin")
    Packet = "SETSPRITE" & SEP_CHAR & SpriteNum & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetName(ByVal nom As String)
Dim Packet As String

    Call EcrireEtat("Changement de nom")
    Packet = "SETNAME" & SEP_CHAR & nom & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendGetStats()
Dim Packet As String

    Call EcrireEtat("Demande des stats")
    Packet = "GETSTATS" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendGetOtherStats(ByVal name As String)
Dim Packet As String

    Call EcrireEtat("Demande des stats de" & name)
    Packet = "GETOTHERSTATS" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerInfoRequest(ByVal name As String)
Dim Packet As String

    Call EcrireEtat("demande d'info sur " & name)
    Packet = "PLAYERINFOREQUEST" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendKick(ByVal name As String)
Dim Packet As String

    Call EcrireEtat("Kick de " & name)
    Packet = "KICKPLAYER" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
    End Sub

Sub SendBan(ByVal name As String)
Dim Packet As String

    Packet = "BANPLAYER" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendBanList()
Dim Packet As String

    Packet = "BANLIST" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditItem()
Dim Packet As String

    Packet = "REQUESTEDITITEM" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveItem(ByVal ItemNum As Long)
Dim Packet As String

Dim filename As String
Dim f  As Long
filename = App.Path & "\items\item" & ItemNum & ".fco"
        
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , Item(ItemNum)
    Close #f
    filename = ReadINI("modif", "objet" & ItemNum, App.Path & "\config.ini")
    If "objet" & ItemNum = filename Then Exit Sub
    Call WriteINI("modif", "objet" & ItemNum, "1", App.Path & "\config.ini")
If HORS_LIGNE = 1 Then Exit Sub
    Packet = "SAVEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).desc & SEP_CHAR & Item(ItemNum).AttackSpeed
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendData(Packet)

End Sub
                
Sub SendRequestEditEmoticon()
Dim Packet As String

    Packet = "REQUESTEDITEMOTICON" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveEmoticon(ByVal EmoNum As Long)
Dim Packet As String

Dim filename As String

    filename = App.Path & "\emoticons.ini"
    
    Call WriteINI("EMOTICONS", "EmoticonC" & EmoNum, Trim(Emoticons(EmoNum).Command), filename)
    Call WriteINI("EMOTICONS", "Emoticon" & EmoNum, Val(Emoticons(EmoNum).Pic), filename)
    filename = ReadINI("modif", "emot" & EmoNum, App.Path & "\config.ini")
    If "emot" & EmoNum = filename Then Exit Sub
    Call WriteINI("modif", "emot" & EmoNum, "1", App.Path & "\config.ini")
If HORS_LIGNE = 1 Then Exit Sub
    Packet = "SAVEEMOTICON" & SEP_CHAR & EmoNum & SEP_CHAR & Trim(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & SEP_CHAR & END_CHAR
    Call SendData(Packet)
    
End Sub

Sub SendRequestEditArrow()
Dim Packet As String

    Packet = "REQUESTEDITARROW" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveArrow(ByVal ArrowNum As Long)
Dim Packet As String

Dim filename As String

    filename = App.Path & "\Arrows.ini"
    
    Call WriteINI("Arrow" & ArrowNum, "ArrowName", Trim(Arrows(ArrowNum).name), filename)
    Call WriteINI("Arrow" & ArrowNum, "ArrowPic", Val(Arrows(ArrowNum).Pic), filename)
    Call WriteINI("Arrow" & ArrowNum, "ArrowRange", Val(Arrows(ArrowNum).Range), filename)
    filename = ReadINI("modif", "flêche" & ArrowNum, App.Path & "\config.ini")
    If "flêche" & ArrowNum = filename Then Exit Sub
    Call WriteINI("modif", "flêche" & ArrowNum, "1", App.Path & "\config.ini")
If HORS_LIGNE = 1 Then Exit Sub
    Packet = "SAVEARROW" & SEP_CHAR & ArrowNum & SEP_CHAR & Trim(Arrows(ArrowNum).name) & SEP_CHAR & Arrows(ArrowNum).Pic & SEP_CHAR & Arrows(ArrowNum).Range & SEP_CHAR & END_CHAR
    Call SendData(Packet)

End Sub
                
Sub SendRequestEditNpc()
Dim Packet As String

    Packet = "REQUESTEDITNPC" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveNpc(ByVal NpcNum As Long)
Dim Packet As String
Dim i As Long

    Dim filename As String
    Dim f As Long
    filename = App.Path & "\pnjs\npc" & NpcNum & ".fcp"
        
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , Npc(NpcNum)
    Close #f
    filename = ReadINI("modif", "pnj" & NpcNum, App.Path & "\config.ini")
    If "pnj" & NpcNum = filename Then Exit Sub
    Call WriteINI("modif", "pnj" & NpcNum, "1", App.Path & "\config.ini")
    If HORS_LIGNE = 1 Then Exit Sub
    Packet = "SAVENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).name) & SEP_CHAR & Trim(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).def & SEP_CHAR & Npc(NpcNum).speed & SEP_CHAR & Npc(NpcNum).magi & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).exp & SEP_CHAR & Npc(NpcNum).SpawnTime & SEP_CHAR & Npc(NpcNum).quetenum & SEP_CHAR
    For i = 1 To MAX_NPC_DROPS
        Packet = Packet & Npc(NpcNum).ItemNPC(i).Chance
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemNum
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemValue & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    Call SendData(Packet)
    
End Sub

Sub SendMapRespawn()
Dim Packet As String

    Packet = "MAPRESPAWN" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendUseItem(ByVal InvNum As Long)
Dim Packet As String

    Packet = "USEITEM" & SEP_CHAR & InvNum & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDropItem(ByVal InvNum, ByVal Ammount As Long)
Dim Packet As String

    Packet = "MAPDROPITEM" & SEP_CHAR & InvNum & SEP_CHAR & Ammount & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendWhosOnline()
Dim Packet As String

    Packet = "WHOSONLINE" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
Sub SendOnlineList()
Dim Packet As String

Packet = "ONLINELIST" & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub
            
Sub SendMOTDChange(ByVal motd As String)
Dim Packet As String

    Packet = "SETMOTD" & SEP_CHAR & motd & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditShop()
Dim Packet As String

    Packet = "REQUESTEDITSHOP" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveShop(ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long, z As Long

Dim filename As String
Dim f As Long

    filename = App.Path & "\shops\shop" & ShopNum & ".fcm"
        
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , Shop(ShopNum)
    Close #f
    filename = ReadINI("modif", "magasin" & ShopNum, App.Path & "\config.ini")
    If "magasin" & ShopNum = filename Then Exit Sub
    Call WriteINI("modif", "magasin" & ShopNum, "1", App.Path & "\config.ini")
If HORS_LIGNE = 1 Then Exit Sub
    Packet = "SAVESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).name) & SEP_CHAR & Trim(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For i = 1 To 6
        For z = 1 To MAX_TRADES
            Packet = Packet & Shop(ShopNum).TradeItem(i).value(z).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).value(z).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).value(z).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).value(z).GetValue & SEP_CHAR
        Next z
    Next i
    Packet = Packet & END_CHAR
    Call SendData(Packet)
    
End Sub

Sub SendRequestEditQuetes()
Dim Packet As String

    Packet = "REQUESTEDITQUETES" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditSpell()
Dim Packet As String

    Packet = "REQUESTEDITSPELL" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveSpell(ByVal SpellNum As Long)
Dim Packet As String

Dim filename As String
Dim f As Long

    filename = App.Path & "\spells\spells" & SpellNum & ".fcs"
        
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , Spell(SpellNum)
    Close #f
    filename = ReadINI("modif", "sort" & SpellNum, App.Path & "\config.ini")
    If "sort" & SpellNum = filename Then Exit Sub
    Call WriteINI("modif", "sort" & SpellNum, "1", App.Path & "\config.ini")
If HORS_LIGNE = 1 Then Exit Sub
    Packet = "SAVESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Trim(Spell(SpellNum).Sound) & SEP_CHAR & Spell(SpellNum).Range & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Spell(SpellNum).AE
    
    
    Call SendData(Packet)

End Sub

Sub SendSaveQuete(ByVal quetenum As Long)
Dim Packet As String

Dim filename As String
Dim f As Long

    filename = App.Path & "\quetes\quete" & quetenum & ".fcq"
        
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , quete(quetenum)
    Close #f
    filename = ReadINI("modif", "quête" & quetenum, App.Path & "\config.ini")
    If "quete" & quetenum = filename Then Exit Sub
    Call WriteINI("modif", "quête" & quetenum, "1", App.Path & "\config.ini")
If HORS_LIGNE = 1 Then Exit Sub
    Packet = "SAVEQUETE" & SEP_CHAR & quetenum & SEP_CHAR & Trim(quete(quetenum).nom) & SEP_CHAR & quete(quetenum).Data1 & SEP_CHAR & quete(quetenum).Data2 & SEP_CHAR & quete(quetenum).Data3 & SEP_CHAR & quete(quetenum).description & SEP_CHAR & quete(quetenum).reponse & SEP_CHAR & quete(quetenum).String1 & SEP_CHAR & quete(quetenum).Temps & SEP_CHAR & quete(quetenum).Type
    Dim i As Long
    For i = 1 To 15
    Packet = Packet & SEP_CHAR & quete(quetenum).indexe(i).Data1 & SEP_CHAR & quete(quetenum).indexe(i).Data2 & SEP_CHAR & quete(quetenum).indexe(i).Data3 & SEP_CHAR & quete(quetenum).indexe(i).String1
    Next i
    Packet = Packet & SEP_CHAR & quete(quetenum).Recompence.exp & SEP_CHAR & quete(quetenum).Recompence.objn1 & SEP_CHAR & quete(quetenum).Recompence.objn2 & SEP_CHAR & quete(quetenum).Recompence.objn3 & SEP_CHAR & quete(quetenum).Recompence.objq1 & SEP_CHAR & quete(quetenum).Recompence.objq2 & SEP_CHAR & quete(quetenum).Recompence.objq3
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendData(Packet)

End Sub

Sub SendRequestEditMap()
Dim Packet As String

    Packet = "REQUESTEDITMAP" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendTradeRequest(ByVal name As String)
Dim Packet As String

    Packet = "PPTRADE" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendAcceptTrade()
Dim Packet As String

    Packet = "ATRADE" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDeclineTrade()
Dim Packet As String

    Packet = "DTRADE" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPartyRequest(ByVal name As String)
Dim Packet As String

    Packet = "PARTY" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendJoinParty()
Dim Packet As String

    Packet = "JOINPARTY" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendLeaveParty()
Dim Packet As String

    Packet = "LEAVEPARTY" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendBanDestroy()
Dim Packet As String
    
    Packet = "BANDESTROY" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestLocation()
Dim Packet As String

    Packet = "REQUESTLOCATION" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
Sub SendSetPlayerSprite(ByVal name As String, ByVal SpriteNum As Byte)
Dim Packet As String

    Packet = "SETPLAYERSPRITE" & SEP_CHAR & name & SEP_CHAR & SpriteNum & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetPlayerName(ByVal name As String, ByVal Nouveau As String)
Dim Packet As String

    Packet = "SETPLAYERNAME" & SEP_CHAR & name & SEP_CHAR & Nouveau & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetPlayerStr(ByVal name As String, ByVal num As Long)
Dim Packet As String

    Packet = "SETPLAYERSTR" & SEP_CHAR & name & SEP_CHAR & num & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetPlayerDef(ByVal name As String, ByVal num As Long)
Dim Packet As String

    Packet = "SETPLAYERDEF" & SEP_CHAR & name & SEP_CHAR & num & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetPlayerVit(ByVal name As String, ByVal num As Long)
Dim Packet As String

    Packet = "SETPLAYERVIT" & SEP_CHAR & name & SEP_CHAR & num & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetPlayerMagi(ByVal name As String, ByVal num As Long)
Dim Packet As String

    Packet = "SETPLAYERMAGI" & SEP_CHAR & name & SEP_CHAR & num & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetPlayerPk(ByVal name As String, ByVal num As Long)
Dim Packet As String

    Packet = "SETPLAYERPK" & SEP_CHAR & name & SEP_CHAR & num & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetPlayerNiveau(ByVal name As String, ByVal num As Long)
Dim Packet As String

    Packet = "SETPLAYERNIVEAU" & SEP_CHAR & name & SEP_CHAR & num & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub


Sub SendSetPlayerExp(ByVal name As String, ByVal num As Long)
Dim Packet As String

    Packet = "SETPLAYEREXP" & SEP_CHAR & name & SEP_CHAR & num & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetPlayerPoint(ByVal name As String, ByVal num As Long)
Dim Packet As String

    Packet = "SETPLAYERPOINT" & SEP_CHAR & name & SEP_CHAR & num & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetPlayerMaxPv(ByVal name As String, ByVal num As Long)
Dim Packet As String

    Packet = "SETPLAYERMAXPV" & SEP_CHAR & name & SEP_CHAR & num & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetPlayerMaxPm(ByVal name As String, ByVal num As Long)
Dim Packet As String

    Packet = "SETPLAYERMAXPM" & SEP_CHAR & name & SEP_CHAR & num & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendGetAdminHelp()
Dim Packet As String

    Packet = "GETADMINHELP" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
