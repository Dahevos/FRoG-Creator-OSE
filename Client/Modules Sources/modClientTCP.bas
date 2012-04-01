Attribute VB_Name = "modClientTCP"
Option Explicit
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Public ServerIP As String
Public PlayerBuffer As String
Public InGame As Boolean
Public TradePlayer As Long
Private MapNumS As Long

Sub TcpInit()
    SEP_CHAR = Chr$(0)
    END_CHAR = Chr$(237)
    PlayerBuffer = vbNullString
    
    Dim filename As String
    filename = App.Path & "\Config\IpConfig.ini"

    frmMirage.Socket.RemoteHost = ReadINI("IPCONFIG", "IP", filename)
    frmMirage.Socket.RemotePort = Val(ReadINI("IPCONFIG", "PORT", filename))
End Sub

Sub TcpDestroy()
    frmMirage.Socket.Close
    
    If frmMainMenu.fraPers.Visible Then frmMainMenu.fraPers.Visible = False
    If frmMainMenu.fraLogin.Visible Then frmMainMenu.fraLogin.Visible = False
    If frmMainMenu.fraNewAccount.Visible Then frmMainMenu.fraNewAccount.Visible = False
    If frmNewChar.Visible Then frmNewChar.Visible = False
End Sub

Sub IncomingData(ByVal DataLength As Long)
Dim Buffer As String
Dim Packet As String
'Dim Top As String * 3
Dim Start As Long

    frmMirage.Socket.GetData Buffer, vbString, DataLength
    PlayerBuffer = PlayerBuffer & Buffer
        
    Start = InStr(PlayerBuffer, END_CHAR)
    Do While Start > 0
        Packet = Mid$(PlayerBuffer, 1, Start - 1)
        PlayerBuffer = Mid$(PlayerBuffer, Start + 1, Len(PlayerBuffer))
        Start = InStr(PlayerBuffer, END_CHAR)
        If Len(Packet) > 0 Then Call HandleData(Packet)
        Sleep 1
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
Dim level As Long
Dim i As Long, n As Long, x As Long, y As Long
Dim ShopNum As Long, GiveItem As Long, GiveValue As Long, GetItem As Long, GetValue As Long
Dim z As Long
Dim Ending As String

'On Error GoTo erreur:

    ' Handle Data
    Parse = Split(data, SEP_CHAR)
        
    ' :::::::::::::::::::::::
    ' :: Get players stats ::
    ' :::::::::::::::::::::::
    If LCase$(Parse(0)) = "datacofr" Then
        
        For i = 1 To 30
            CoffreTmp(i).Numeros = Val(Parse((i * 3) - 2))
            CoffreTmp(i).Valeur = Val(Parse((i * 3) - 1))
            CoffreTmp(i).Durabiliter = Val(Parse((i * 3)))
        Next i
        Call frmbank.ActPic
        Exit Sub
    End If
        
    If LCase$(Parse(0)) = "picvalue" Then
        PIC_PL = Val(Parse(1))
        PIC_NPC1 = Val(Parse(2))
        PIC_NPC2 = Val(Parse(3))
        If PIC_NPC1 <= 0 Then PIC_NPC1 = 2
        If PIC_PL <= 0 Then PIC_PL = 64
        If PIC_NPC2 < 0 Then PIC_NPC2 = 32
        AccModo = Val(Parse(4))
        AccMapeur = Val(Parse(5))
        AccDevelopeur = Val(Parse(6))
        AccAdmin = Val(Parse(7))
        Exit Sub
    End If
    
    If LCase$(Parse(0)) = "maxinfo" Then
        GAME_NAME = Trim$(Parse(1))
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
        MAX_INV = Val(Parse(14))
        MAX_PETS = Val(Parse(16))
        MAX_METIER = Val(Parse(17))
        MAX_RECETTE = Val(Parse(18))
        
        For i = 1 To MAX_INV - 1
            If Loading = False Then Load frmMirage.picInv(i)
            
            x = Int(i / 3)
            frmMirage.picInv(i).Top = 8 + 40 * x
            frmMirage.picInv(i).Left = 8 + (i - x * 3) * 40
            frmMirage.picInv(i).Visible = True
        Next
        
        frmMirage.Picture9.Height = frmMirage.picInv(i - 1).Top + 40
        
        For i = 1 To MAX_PLAYER_SPELLS - 1
            If Loading = False Then Load frmMirage.picspell(i)
            
            x = Int(i / 3)
            frmMirage.picspell(i).Top = 8 + 40 * x
            frmMirage.picspell(i).Left = 8 + (i - x * 3) * 40
            frmMirage.picspell(i).Visible = True
        Next
        
        If MAX_MAPX <= 20 Then PicScHeight = (MAX_MAPY + 1) * PIC_Y: PicScWidth = (MAX_MAPX + 1) * PIC_X
        
        ReDim quete(1 To MAX_QUETES) As QueteRec
        ReDim Map(1 To MAX_MAPS) As MapRec
        ReDim Player(1 To MAX_PLAYERS) As PlayerRec
        ReDim PlayerAnim(1 To MAX_PLAYERS, 0 To 4) As Long
        For i = 1 To MAX_PLAYERS
            PlayerAnim(i, 0) = 0
        Next i
        ReDim Item(1 To MAX_ITEMS) As ItemRec
        ReDim Npc(1 To MAX_NPCS) As NpcRec
        ReDim MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
        ReDim Shop(1 To MAX_SHOPS) As ShopRec
        ReDim Spell(1 To MAX_SPELLS) As SpellRec
        ReDim Bubble(1 To MAX_PLAYERS) As ChatBubble
        ReDim SaveMapItem(1 To MAX_MAP_ITEMS) As MapItemRec
        For i = 1 To MAX_MAPS
            ReDim Map(i).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
            ReDim Map(i).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        Next i
        ReDim TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
        ReDim Emoticons(0 To MAX_EMOTICONS) As EmoRec
        ReDim MapReport(1 To MAX_MAPS) As MapRec
        ReDim Pets(1 To MAX_PETS) As PetsRec
        ReDim Metier(1 To MAX_METIER) As MetierRec
        ReDim recette(1 To MAX_RECETTE) As RecetteRec
        
        MAX_SPELL_ANIM = MAX_MAPX * MAX_MAPY
        
        MAX_BLT_LINE = 10
        ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
        ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec
        
        For i = 1 To MAX_PLAYERS
            ReDim Player(i).Inv(1 To MAX_INV) As PlayerInvRec
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
        
        Call ClearMap
        
        For i = 1 To MAX_MAPS
            DoEvents
            Call LoadMap(i)
        Next i
    
        frmMirage.Caption = Trim$(GAME_NAME) '& " - Powered by Konfuze ORPG Creation (www.Konfuze.com)"
        App.Title = GAME_NAME
        Loading = True
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Multi-Serveur ::
    ' :::::::::::::::::::
    If LCase$(Parse(0)) = "serverresults" Then
        frmServerChooser.lstServers.AddItem ReadINI("SERVER" & Val(Parse(1)), "Name", App.Path & "\Config\Serveur.ini") & " - Ouvert! (" & Val(Parse(2)) & "/" & Val(Parse(3)) & ")"
        CHECK_WAIT = False
        Exit Sub
    End If
        
    ' :::::::::::::::::::
    ' :: Npc hp packet ::
    ' :::::::::::::::::::
    If LCase$(Parse(0)) = "npchp" Then
        n = Val(Parse(1))
 
        MapNpc(n).HP = Val(Parse(2))
        MapNpc(n).MaxHp = Val(Parse(3))
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Npc mp packet ::
    ' :::::::::::::::::::
    If LCase$(Parse(0)) = "npcmp" Then
        n = Val(Parse(1))
 
        MapNpc(n).MP = Val(Parse(2))
        MapNpc(n).MaxMp = Val(Parse(3))
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::::::
    ' :: Alert message packet ::
    ' ::::::::::::::::::::::::::
    If LCase$(Parse(0)) = "alertmsg" Then
        frmMirage.Visible = False
        frmsplash.Visible = False
        frmMainMenu.Visible = True
        DoEvents

        Msg = Parse(1)
        Call MsgBox(Msg, vbOKOnly, GAME_NAME)
        Call GameDestroy
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Plain message packet ::
    ' ::::::::::::::::::::::::::
    If LCase$(Parse(0)) = "plainmsg" Then
        frmsplash.Visible = False
        n = Val(Parse(2))
        If frmMirage.Visible And n <> 6 Then frmMirage.Hide
                
        If n = 1 Then frmMainMenu.fraNewAccount.Visible = True: frmMainMenu.Show
        If n = 2 Then frmMainMenu.fraLogin.Visible = True: frmMainMenu.Show
        If n = 3 Then frmMainMenu.fraLogin.Visible = True: frmMainMenu.Show
        If n = 4 Then frmNewChar.Show
        If n = 5 Then frmMainMenu.fraPers.Visible = True: frmMainMenu.Show
        
        Msg = Parse(1)
        Call MsgBox(Msg, vbOKOnly, GAME_NAME)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: All characters packet ::
    ' :::::::::::::::::::::::::::
    If LCase$(Parse(0)) = "allchars" Then
        n = 1
        
        frmMainMenu.fraPers.Visible = True
        frmsplash.Visible = False
        
        frmMainMenu.lstChars.Clear
        
        charSelectNum = 1
        
        For i = 1 To MAX_CHARS
            name = Parse(n)
            Msg = Parse(n + 1)
            level = Val(Parse(n + 2))
            
            charSelect(i).name = Parse(n)
            charSelect(i).classe = Parse(n + 1)
            charSelect(i).level = Val(Parse(n + 2))
            If charSelect(charSelectNum).name <> "" Then charSelect(i).sprt = Val(Parse(n + 3)) Else charSelect(i).sprt = 0
            
            If Trim$(name) = vbNullString Then frmMainMenu.lstChars.AddItem "Emplacement libre" Else frmMainMenu.lstChars.AddItem name '& " - niveaux " & level & " - " & Msg
            
            n = n + 4
        Next i
        
        For i = 1 To 4
            If i = 1 Then Ending = ".gif"
            If i = 2 Then Ending = ".jpg"
            If i = 3 Then Ending = ".png"
            If i = 4 Then Ending = ".bmp"
            
            If FileExiste("GFX/Sprites/Sprites" & charSelect(charSelectNum).sprt & Ending) Then
                frmMainMenu.PicChar.Picture = LoadPNG(App.Path & "/GFX/Sprites/Sprites" & charSelect(charSelectNum).sprt & Ending)
            End If
        Next i
        frmMainMenu.PicChar.Height = frmMainMenu.PicChar.Height / 4
        frmMainMenu.PicChar.Width = frmMainMenu.PicChar.Width / 4
        If frmMainMenu.PicChar.Width > 960 Then
            frmMainMenu.PicChar.Width = 960
        End If
        If frmMainMenu.PicChar.Height > 960 Then
            frmMainMenu.PicChar.Height = 960
        End If
        If frmMainMenu.PicChar.Width > 480 Then
            frmMainMenu.PicChar.Left = 840 - frmMainMenu.PicChar.Width + 480
        Else
            frmMainMenu.PicChar.Left = 840
        End If
        If charSelect(charSelectNum).name <> "" Then
            frmMainMenu.lblCharNom.Caption = charSelect(charSelectNum).name
            frmMainMenu.lblCharLvl.Caption = "Niv. " & charSelect(charSelectNum).level
            frmMainMenu.lblCharClasse.Caption = charSelect(charSelectNum).classe
        Else
            frmMainMenu.lblCharNom.Caption = "Slot Libre"
            frmMainMenu.lblCharLvl.Caption = ""
            frmMainMenu.lblCharClasse.Caption = ""
        End If
        
        frmMainMenu.lstChars.ListIndex = 0
        frmMainMenu.lstChars.SetFocus
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::
    ' :: Login was successful packet ::
    ' :::::::::::::::::::::::::::::::::
    If LCase$(Parse(0)) = "loginok" Then
        ' Now we can receive game data
        MyIndex = Val(Parse(1))
        
        frmsplash.Visible = True
        frmsplash.Shape1.Width = 255
        frmMainMenu.fraPers.Visible = False
        
        DoEvents
        
        Call SetStatus("Réception des données en cours...")
        DoEvents
        frmsplash.Shape1.Width = frmsplash.Shape1.Width + 2000
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: New character classes data packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If LCase$(Parse(0)) = "newcharclasses" Then
        n = 1
        
        ' Max classes
        Max_Classes = Val(Parse(n))
        ReDim Class(0 To Max_Classes) As ClassRec
        
        n = n + 1

        For i = 0 To Max_Classes
            With Class(i)
            .name = Parse(n)
            
            .HP = Val(Parse(n + 1))
            .MP = Val(Parse(n + 2))
            .SP = Val(Parse(n + 3))
            
            .STR = Val(Parse(n + 4))
            .DEF = Val(Parse(n + 5))
            .speed = Val(Parse(n + 6))
            .MAGI = Val(Parse(n + 7))
            '.INTEL = Val(Parse(n + 8))
            .MaleSprite = Val(Parse(n + 8))
            .FemaleSprite = Val(Parse(n + 9))
            .Locked = Val(Parse(n + 10))
            End With
        n = n + 11
        Next i
        
        ' Used for if the player is creating a new character
        frmNewChar.Visible = True
        frmsplash.Visible = False

        frmNewChar.cmbClass.Clear
        For i = 0 To Max_Classes
            If Class(i).Locked = 0 Then
                frmNewChar.cmbClass.AddItem Trim$(Class(i).name)
                frmNewChar.cmbClass.ItemData(frmNewChar.cmbClass.ListCount - 1) = i
            End If
        Next i
        
        With frmNewChar
            .cmbClass.ListIndex = 0
            .lblHP.Caption = STR$(Class(0).HP)
            .lblMP.Caption = STR$(Class(0).MP)
            .lblSP.Caption = STR$(Class(0).SP)
        
            .lblSTR.Caption = STR$(Class(0).STR)
            .lblDEF.Caption = STR$(Class(0).DEF)
            .lblSPEED.Caption = STR$(Class(0).speed)
            .lblMAGI.Caption = STR$(Class(0).MAGI)
            .Picpic.Height = (PIC_Y + (PIC_Y / 2))
            .Picture4.Height = (PIC_Y + (PIC_Y / 2)) + 4
        End With
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Classes data packet ::
    ' :::::::::::::::::::::::::
    If LCase$(Parse(0)) = "classesdata" Then
        n = 1
        
        ' Max classes
        Max_Classes = Val(Parse(n))
        ReDim Class(0 To Max_Classes) As ClassRec
        
        n = n + 1
        
        For i = 0 To Max_Classes
            With Class(i)
                .name = Parse(n)
                
                .HP = Val(Parse(n + 1))
                .MP = Val(Parse(n + 2))
                .SP = Val(Parse(n + 3))
                
                .STR = Val(Parse(n + 4))
                .DEF = Val(Parse(n + 5))
                .speed = Val(Parse(n + 6))
                .MAGI = Val(Parse(n + 7))
                
                .Locked = Val(Parse(n + 8))
            End With
            n = n + 9
        Next i
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: In game packet ::
    ' ::::::::::::::::::::
    If LCase$(Parse(0)) = "ingame" Then
        InGame = True
        Call GameInit
        Call GameLoop
        If Parse(1) = END_CHAR Then MsgBox ("here"): End
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player inventory packet ::
    ' :::::::::::::::::::::::::::::
    If LCase$(Parse(0)) = "playerinv" Then
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
    If LCase$(Parse(0)) = "playerinvupdate" Then
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
    If LCase$(Parse(0)) = "playerworneq" Then
        z = Val(Parse(1))
        If z <= 0 Then Exit Sub
        
        Call SetPlayerArmorSlot(z, Val(Parse(2)))
        Call SetPlayerWeaponSlot(z, Val(Parse(3)))
        Call SetPlayerHelmetSlot(z, Val(Parse(4)))
        Call SetPlayerShieldSlot(z, Val(Parse(5)))
        Call SetPlayerPetSlot(z, Val(Parse(6)))
        'PAPERDOLL
        'Player(z).Casque = Val(Parse(6))
        'Player(z).Armure = Val(Parse(7))
        'Player(z).Arme = Val(Parse(8))
        'Player(z).Bouclier = Val(Parse(9))
        'FIN PAPERDOLL
        
        If z = MyIndex Then Call UpdateVisInv
        Exit Sub
    End If
    
    If (LCase$(Parse(0)) = "playerpet") Then
        i = Val(Parse(1))
        Player(i).pet.Dir = Val(Parse(2))
        Player(i).pet.x = Val(Parse(3))
        Player(i).pet.y = Val(Parse(4))

        'If Player(i).pet.x = Player(i).x And Player(i).pet.y = Player(i).y Then Exit Sub

        If Val(Parse(5)) <> 1 Then Exit Sub

        Select Case Player(i).pet.Dir
            Case DIR_UP
                Player(i).pet.YOffset = PIC_Y
            Case DIR_DOWN
                Player(i).pet.YOffset = PIC_Y * -1
            Case DIR_LEFT
                Player(i).pet.XOffset = PIC_X
            Case DIR_RIGHT
                Player(i).pet.XOffset = PIC_X * -1
        End Select
        If IsPlaying(i) And i < MAX_PLAYERS And i > 0 Then If Player(i).pet.Anim = 0 Then Player(i).pet.Anim = 2 Else Player(i).pet.Anim = 0
        Exit Sub
    End If
    
    If (LCase$(Parse(0)) = "updatepet") Then
        n = Val(Parse(1))
        Pets(n).nom = Parse(2)
        Pets(n).sprite = Parse(3)
        Pets(n).addForce = Parse(4)
        Pets(n).addDefence = Parse(5)
    End If
    
    ' ::::::::::::
    ' :: Metier ::
    ' ::::::::::::
    If (LCase$(Parse(0)) = "updatemetier") Then
        n = Val(Parse(1))
        Metier(n).nom = Parse(2)
        Metier(n).Type = Val(Parse(3))
        Metier(n).desc = Parse(4)
        x = 5
        For i = 0 To MAX_DATA_METIER
            For z = 0 To 1
                Metier(n).data(i, z) = Val(Parse(x))
                x = x + 1
            Next z
        Next i
        Exit Sub
    End If
    
    If (LCase$(Parse(0)) = "playermetier") Then
        n = Val(Parse(1))
        Player(n).Metier = Val(Parse(2))
        Player(n).MetierLvl = Val(Parse(3))
        Player(n).MetierExp = Val(Parse(4))
    End If
    
    If (LCase$(Parse(0)) = "metier") Then
        If Player(MyIndex).Metier > 0 Then
            frmMirage.pictMetier.Visible = True
            frmMirage.lblmetier(0).Caption = "Metier: " + Metier(Player(MyIndex).Metier).nom
            frmMirage.lblmetier(1).Caption = "Niveau: " + CStr(Player(MyIndex).MetierLvl)
            frmMirage.lblmetier(2).Caption = "Exp: " + CStr(Player(MyIndex).MetierExp) + "/" + CStr((Player(MyIndex).MetierLvl + 1) * 2)
            frmMirage.lblmetier(3).Caption = "Description: " + Metier(Player(MyIndex).Metier).desc
        End If
    End If
    
    ' :::::::::::::
    ' :: recette ::
    ' :::::::::::::
    If (LCase$(Parse(0)) = "updaterecette") Then
        n = Val(Parse(1))
        recette(n).nom = Parse(2)
        x = 3
        For i = 0 To 9
            For z = 0 To 1
                recette(n).InCraft(i, z) = Val(Parse(x))
                x = x + 1
            Next z
        Next i
        For z = 0 To 1
            recette(n).craft(z) = Val(Parse(x))
            x = x + 1
        Next z
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Player points packet ::
    ' ::::::::::::::::::::::::::
    If LCase$(Parse(0)) = "playerpoints" Then
        Player(MyIndex).POINTS = Val(Parse(1))
        frmMirage.lblPoints.Caption = Val(Parse(1))
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::
    ' :: Player hp packet ::
    ' ::::::::::::::::::::::
    If LCase$(Parse(0)) = "playerhp" Then
        Player(MyIndex).MaxHp = Val(Parse(1))
        Call SetPlayerHP(MyIndex, Val(Parse(2)))
        If GetPlayerMaxHP(MyIndex) > 0 Then frmMirage.svie.Width = (((GetPlayerHP(MyIndex) / 1425) / (GetPlayerMaxHP(MyIndex) / 1425)) * 1425): frmMirage.lvie.Caption = "PV : " & GetPlayerHP(MyIndex) & " / " & GetPlayerMaxHP(MyIndex)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Party Hp Packet ::
    ' :::::::::::::::::::::
    If LCase$(Parse(0)) = "partyhp" Then
        n = Val(Parse(1))
 
        Player(n).PartyIndex = Val(Parse(2))
        Player(n).MaxHp = Val(Parse(3))
        Player(n).HP = Val(Parse(4))
        Player(n).MaxMp = Val(Parse(5))
        Player(n).MP = Val(Parse(6))
        'With frmMirage
        '    For i = 0 To 2
        '        If Val(.lblPPName(i).Tag) = n Then
        '            .shpPPLife(i).Width = Player(n).HP / Player(n).MaxHp * .backPPLife(i).Width
        '            .shpPPMana(i).Width = Player(n).MP / Player(n).MaxMp * .backPPMana(i).Width
        '        End If
        '    Next
        'End With
        Exit Sub
    End If
    If LCase$(Parse(0)) = "partyindex" Then
        n = Val(Parse(1))
 
        Player(n).PartyIndex = Val(Parse(2))
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Player mp packet ::
    ' ::::::::::::::::::::::
    If LCase$(Parse(0)) = "playermp" Then
        Player(MyIndex).MaxMp = Val(Parse(1))
        Call SetPlayerMP(MyIndex, Val(Parse(2)))
        If GetPlayerMaxMP(MyIndex) > 0 Then frmMirage.smana.Width = (((GetPlayerMP(MyIndex) / 1425) / (GetPlayerMaxMP(MyIndex) / 1425)) * 1425): frmMirage.lmana.Caption = "PM : " & GetPlayerMP(MyIndex) & " / " & GetPlayerMaxMP(MyIndex)
        Exit Sub
    End If
    
    ' speech bubble parse
    If (LCase$(Parse(0)) = "mapmsg2") Then
        Bubble(Val(Parse(2))).Text = Parse(1)
        Bubble(Val(Parse(2))).Created = GetTickCount()
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Player sp packet ::
    ' ::::::::::::::::::::::
    'If LCase$(Parse(0)) = "playersp" Then
       ' Player(MyIndex).MaxSP = Val(Parse(1))
        'Call SetPlayerSP(MyIndex, Val(Parse(2)))
        'If GetPlayerMaxSP(MyIndex) > 0 Then
            'frmMirage.lblSP.Caption = Int(GetPlayerSP(MyIndex) / GetPlayerMaxSP(MyIndex) * 100) & "%"
        'End If
'        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::
    ' :: Player Stats Packet ::
    ' :::::::::::::::::::::::::
    If (LCase$(Parse(0)) = "playerstatspacket") Then
        Dim SubStr As Long, SubDef As Long, SubMagi As Long, SubSpeed As Long
        Dim apf As Byte, apd As Byte
        SubStr = 0
        SubDef = 0
        SubMagi = 0
        SubSpeed = 0
        apf = 0
        apd = 0
        If GetPlayerWeaponSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerArmorSlot(MyIndex) > 0 Then
        If GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex)) > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).Type = ITEM_TYPE_MONTURE Then GoTo mont:
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddSpeed
        End If
        End If
        
mont:
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
        If GetPlayerPetSlot(MyIndex) > 0 Then
            apf = apf + Pets(Item(GetPlayerInvItemNum(MyIndex, GetPlayerPetSlot(MyIndex))).Data1).addForce
            apd = apd + Pets(Item(GetPlayerInvItemNum(MyIndex, GetPlayerPetSlot(MyIndex))).Data1).addDefence
        End If

        If SubStr > 0 Or apf > 0 Then frmMirage.lblSTR.Caption = Val(Parse(1)) - SubStr & " (+" & SubStr + apf & ")" Else If SubStr < 0 Then frmMirage.lblSTR.Caption = Val(Parse(1)) - SubStr & " (" & apf - SubStr & ")" Else frmMirage.lblSTR.Caption = Val(Parse(1))
        If SubDef > 0 Or apd > 0 Then frmMirage.lblDEF.Caption = Val(Parse(2)) - SubDef & " (+" & SubDef + apd & ")" Else If SubDef < 0 Then frmMirage.lblDEF.Caption = Val(Parse(2)) - SubDef & " (" & apd - SubDef & ")" Else frmMirage.lblDEF.Caption = Val(Parse(2))
        If SubMagi > 0 Then frmMirage.lblMAGI.Caption = Val(Parse(4)) - SubMagi & " (+" & SubMagi & ")" Else If SubMagi < 0 Then frmMirage.lblMAGI.Caption = Val(Parse(4)) - SubMagi & " (" & SubMagi & ")" Else frmMirage.lblMAGI.Caption = Val(Parse(4))
        If SubSpeed > 0 Then frmMirage.lblSPEED.Caption = Val(Parse(3)) - SubSpeed & " (+" & SubSpeed & ")" Else If SubSpeed < 0 Then frmMirage.lblSPEED.Caption = Val(Parse(3)) - SubSpeed & " (" & SubSpeed & ")" Else frmMirage.lblSPEED.Caption = Val(Parse(3))
        Call SetPlayerExp(MyIndex, Val(Parse(6)))
        nelvl = Val(Parse(5))
        frmMirage.lexp.Caption = "EXP : " & Val(Parse(6)) & " / " & Val(Parse(5))
        frmMirage.sexp.Width = (((Val(Parse(6)) / 1425) / (Val(Parse(5)) / 1425)) * 1425)
        frmMirage.monnom.Caption = Trim$(Player(MyIndex).name) & " - " & Trim$(Class(Player(MyIndex).Class).name) & " - Niv" & Val(Parse(7))
        Player(MyIndex).level = Val(Parse(7))
        Exit Sub
    End If
                

    ' ::::::::::::::::::::::::
    ' :: Player data packet ::
    ' ::::::::::::::::::::::::
    If LCase$(Parse(0)) = "playerdata" Then
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
        Player(i).PartyIndex = Val(Parse(14))
        
        ' Make sure they aren't walking
        Player(i).Moving = 0
        Player(i).XOffset = 0
        Player(i).YOffset = 0
        
        ' Check if the player is the client player, and if so reset directions
        If i = MyIndex Then DirUp = False: DirDown = False: DirLeft = False: DirRight = False

        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Player movement packet ::
    ' ::::::::::::::::::::::::::::
    If (LCase$(Parse(0)) = "playermove") Then
        i = Val(Parse(1))
        x = Val(Parse(2))
        y = Val(Parse(3))
        Dir = Val(Parse(4))
        n = Val(Parse(5))

        If Dir < DIR_DOWN Or Dir > DIR_UP Then Exit Sub

        Call SetPlayerX(i, x)
        Call SetPlayerY(i, y)
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
        If IsPlaying(i) And i < MAX_PLAYERS And i > 0 Then If Player(i).Anim = 0 Then Player(i).Anim = 2 Else Player(i).Anim = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Npc movement packet ::
    ' :::::::::::::::::::::::::
    If (LCase$(Parse(0)) = "npcmove") Then
        i = Val(Parse(1))
        x = Val(Parse(2))
        y = Val(Parse(3))
        Dir = Val(Parse(4))
        n = Val(Parse(5))

        MapNpc(i).x = x
        MapNpc(i).y = y
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
        If i < MAX_MAP_NPCS And i > 0 Then If PNJAnim(i) = 0 Then PNJAnim(i) = 2 Else PNJAnim(i) = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player direction packet ::
    ' :::::::::::::::::::::::::::::
    If (LCase$(Parse(0)) = "playerdir") Then
        i = Val(Parse(1))
        Dir = Val(Parse(2))
        
        If Dir < DIR_DOWN Or Dir > DIR_UP Then Exit Sub

        Call SetPlayerDir(i, Dir)
        
        Player(i).XOffset = 0
        Player(i).YOffset = 0
        Player(i).Moving = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: NPC direction packet ::
    ' ::::::::::::::::::::::::::
    If (LCase$(Parse(0)) = "npcdir") Then
        i = Val(Parse(1))
        Dir = Val(Parse(2))
        
        If Dir < DIR_DOWN Or Dir > DIR_UP Then Exit Sub
        
        MapNpc(i).Dir = Dir
        
        MapNpc(i).XOffset = 0
        MapNpc(i).YOffset = 0
        MapNpc(i).Moving = 0
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::::::::::
    ' :: Player XY location packet ::
    ' :::::::::::::::::::::::::::::::
    If (LCase$(Parse(0)) = "playerxy") Then
        x = Val(Parse(1))
        y = Val(Parse(2))
        
        If x > MAX_MAPX Or x < 0 Then Exit Sub
        If y > MAX_MAPY Or y < 0 Then Exit Sub
        
        Call SetPlayerX(MyIndex, x)
        Call SetPlayerY(MyIndex, y)
        
        ' Make sure they aren't walking
        Player(MyIndex).Moving = 0
        Player(MyIndex).XOffset = 0
        Player(MyIndex).YOffset = 0
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
    If (LCase$(Parse(0)) = "attack") Then
        i = Val(Parse(1))
        
        ' Set player to attacking
        Player(i).Attacking = 1
        Player(i).AttackTimer = GetTickCount
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' ::Player controle packet::
    ' ::::::::::::::::::::::::::
    If (LCase$(Parse(0)) = "conoff") Then
        If ConOff = False Then ConOff = True Else ConOff = False
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: NPC attack packet ::
    ' :::::::::::::::::::::::
    If (LCase$(Parse(0)) = "npcattack") Then
        i = Val(Parse(1))
        
        ' Set player to attacking
        MapNpc(i).Attacking = 1
        MapNpc(i).AttackTimer = GetTickCount
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Check for map packet ::
    ' ::::::::::::::::::::::::::
    If (LCase$(Parse(0)) = "checkformap") Then
        ' Erase all players except self
        For i = 1 To MAX_PLAYERS
            If i <> MyIndex Then Call SetPlayerMap(i, 0)
        Next i
        
        ' Erase all temporary tile values
        Call ClearTempTile

        ' Get map num
        x = Val(Parse(1))
        
        ' Get revision
        y = Val(Parse(2))
        
        If FileExiste("maps\map" & x & ".fcc") Then
            ' Check to see if the revisions match
            If GetMapRevision(x) = y Then
                ' We do so we dont need the map
                
                ' Load the map
                'Call LoadMap(X)
                
                Call SendData("needmap" & SEP_CHAR & "no" & SEP_CHAR & END_CHAR)
                
                Call InitPano(x)
                Call InitNightAndFog(x)
                Exit Sub
            End If
        End If
        
        ' Either the revisions didn't match or we dont have the map, so we need it
        OldMap = GetPlayerMap(MyIndex)
        Call SendData("needmap" & SEP_CHAR & "yes" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    If LCase$(Parse(0)) = "mapdown" Then
    Dim URL As String
    Dim rep As String
        z = Val(Parse(1))
        URL = Trim$(Parse(2))
        rep = Trim$(Parse(3))
              
        If z <= 0 Or z > MAX_MAPS Then Exit Sub
        If Mid(URL, Len(URL)) <> "/" And Mid(rep, 1, 1) <> "/" Then URL = URL & "/"
        If rep <> "/" Then If Mid(rep, Len(rep)) <> "/" Then rep = rep & "/"
        If Mid(URL, Len(URL)) = "/" And Mid(rep, 1, 1) = "/" Then rep = Mid(rep, 2)
                
        If Mid(URL, Len(URL)) = "/" And rep = "/" Then
            If FileExiste("Maps\map" & z & ".fcc") Then Call DeleteUrlCacheEntry(URL & "map" & z & ".fcc")
            Call URLDownloadToFile(0, URL & "map" & z & ".fcc", App.Path & "\Maps\map" & z & ".fcc", 0, 0)
        ElseIf Mid(URL, Len(URL)) <> "/" And Mid(rep, 1, 1) = "/" Then
            If FileExiste("Maps\map" & z & ".fcc") Then Call DeleteUrlCacheEntry(URL & rep & "map" & z & ".fcc")
            Call URLDownloadToFile(0, URL & rep & "map" & z & ".fcc", App.Path & "\Maps\map" & z & ".fcc", &O10, 0)
        Else
            If FileExiste("Maps\map" & z & ".fcc") Then Call DeleteUrlCacheEntry(URL & rep & "map" & z & ".fcc")
            Call URLDownloadToFile(0, URL & rep & "map" & z & ".fcc", App.Path & "\Maps\map" & z & ".fcc", &O10, 0)
        End If
        Call LoadMap(z)
        Call InitPano(x)
        Call InitNightAndFog(x)
    End If
    
    If LCase$(Parse(0)) = "notwarp" Then If (GetPlayerY(MyIndex) <= 0 Or GetPlayerY(MyIndex) >= MAX_MAPY Or GetPlayerX(MyIndex) <= 0 Or GetPlayerX(MyIndex) >= MAX_MAPX) And GettingMap Then GettingMap = False
    
    ' :::::::::::::::::::::
    ' :: Map data packet ::
    ' :::::::::::::::::::::
    
    If LCase$(Parse(0)) = "mapdatas" Then
        
        n = 1
        MapNumS = Val(Parse(1))
        With Map(Val(Parse(1)))
            .name = Parse(n + 1)
            .Revision = Val(Parse(n + 2))
            .Moral = Val(Parse(n + 3))
            .Up = Val(Parse(n + 4))
            .Down = Val(Parse(n + 5))
            .Left = Val(Parse(n + 6))
            .Right = Val(Parse(n + 7))
            .Music = Parse(n + 8)
            .BootMap = Val(Parse(n + 9))
            .BootX = Val(Parse(n + 10))
            .BootY = Val(Parse(n + 11))
            .Indoors = Val(Parse(n + 12))
            .PanoInf = Parse(n + 13)
            .TranInf = Val(Parse(n + 14))
            .PanoSup = Parse(n + 15)
            .TranSup = Val(Parse(n + 16))
            .Fog = Val(Parse(n + 17))
            .FogAlpha = Val(Parse(n + 18))
            .guildSoloView = Parse(n + 19)
            .petView = Parse(n + 20)
            .traversable = Parse(n + 21)
        End With
        
        GettingMap = True
        Exit Sub
    End If
    
    If LCase$(Parse(0)) = "maptiles" Then
        n = 1
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                With Map(MapNumS).Tile(x, y)
                    .Ground = Val(Parse(n))
                    .Mask = Val(Parse(n + 1))
                    .Anim = Val(Parse(n + 2))
                    .Mask2 = Val(Parse(n + 3))
                    .M2Anim = Val(Parse(n + 4))
                    .Mask3 = Val(Parse(n + 32))
                    .M3Anim = Val(Parse(n + 30))
                    .Fringe = Val(Parse(n + 5))
                    .FAnim = Val(Parse(n + 6))
                    .Fringe2 = Val(Parse(n + 7))
                    .F2Anim = Val(Parse(n + 8))
                    .Fringe3 = Val(Parse(n + 26))
                    .F3Anim = Val(Parse(n + 27))
                    .Type = Val(Parse(n + 9))
                    .Data1 = Val(Parse(n + 10))
                    .Data2 = Val(Parse(n + 11))
                    .Data3 = Val(Parse(n + 12))
                    .String1 = Parse(n + 13)
                    .String2 = Parse(n + 14)
                    .String3 = Parse(n + 15)
                    .Light = Val(Parse(n + 16))
                    .GroundSet = Val(Parse(n + 17))
                    .MaskSet = Val(Parse(n + 18))
                    .AnimSet = Val(Parse(n + 19))
                    .Mask2Set = Val(Parse(n + 20))
                    .M2AnimSet = Val(Parse(n + 21))
                    .Mask3Set = Val(Parse(n + 33))
                    .M3AnimSet = Val(Parse(n + 31))
                    .FringeSet = Val(Parse(n + 22))
                    .FAnimSet = Val(Parse(n + 23))
                    .Fringe2Set = Val(Parse(n + 24))
                    .F2AnimSet = Val(Parse(n + 25))
                    .Fringe3Set = Val(Parse(n + 28))
                    .F3AnimSet = Val(Parse(n + 29))
                End With
                
                n = n + 34
            Next x
        Next y
        'GettingMap = True
        Exit Sub
    End If
    
    If LCase$(Parse(0)) = "mapnpcs" Then
        n = 1
        For x = 1 To MAX_MAP_NPCS
            With Map(MapNumS)
                .Npc(x) = Val(Parse(n))
                n = n + 1
                .Npcs(x).x = Val(Parse(n))
                n = n + 1
                .Npcs(x).y = Val(Parse(n))
                n = n + 1
                .Npcs(x).x1 = Val(Parse(n))
                n = n + 1
                .Npcs(x).y1 = Val(Parse(n))
                n = n + 1
                .Npcs(x).x2 = Val(Parse(n))
                n = n + 1
                .Npcs(x).y2 = Val(Parse(n))
                n = n + 1
                .Npcs(x).Hasardm = Val(Parse(n))
                n = n + 1
                .Npcs(x).Hasardp = Val(Parse(n))
                n = n + 1
                .Npcs(x).boucle = Val(Parse(n))
                n = n + 1
                .Npcs(x).Imobile = Val(Parse(n))
                n = n + 1
            End With
        Next x
                
        ' Save the map
        Call SaveLocalMap(MapNumS)
                
        'GettingMap = True
        Exit Sub
    End If
            
    ' :::::::::::::::::::::::::::
    ' :: Map items data packet ::
    ' :::::::::::::::::::::::::::
    If LCase$(Parse(0)) = "mapitemdata" Then
        n = 1
        
        For i = 1 To MAX_MAP_ITEMS
            SaveMapItem(i).num = Val(Parse(n))
            SaveMapItem(i).Value = Val(Parse(n + 1))
            SaveMapItem(i).dur = Val(Parse(n + 2))
            SaveMapItem(i).x = Val(Parse(n + 3))
            SaveMapItem(i).y = Val(Parse(n + 4))
            
            n = n + 5
        Next i
        
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::::
    ' :: Map npc data packet ::
    ' :::::::::::::::::::::::::
    If LCase$(Parse(0)) = "mapnpcdata" Then
        n = 1
        
        For i = 1 To MAX_MAP_NPCS
            SaveMapNpc(i).num = Val(Parse(n))
            SaveMapNpc(i).x = Val(Parse(n + 1))
            SaveMapNpc(i).y = Val(Parse(n + 2))
            SaveMapNpc(i).Dir = Val(Parse(n + 3))
            
            n = n + 4
        Next i
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Map send completed packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase$(Parse(0)) = "mapdone" Then
        'Map = SaveMap
        
        Call InitPano(GetPlayerMap(MyIndex))
        Call InitNightAndFog(GetPlayerMap(MyIndex))
        
        For i = 1 To MAX_MAP_ITEMS
            MapItem(i) = SaveMapItem(i)
        Next i
        
        For i = 1 To MAX_MAP_NPCS
            MapNpc(i) = SaveMapNpc(i)
        Next i
       
        GettingMap = False
        
        ' Play music
        If OldMap > 0 Then
            If Trim$(Map(GetPlayerMap(MyIndex)).Music) = Trim$(Map(OldMap).Music) Then OldMap = GetPlayerMap(MyIndex): Exit Sub
            If Trim$(Map(GetPlayerMap(MyIndex)).Music) <> "Aucune" Then Call PlayMidi(Trim$(Map(GetPlayerMap(MyIndex)).Music)) Else Call StopMidi
            OldMap = GetPlayerMap(MyIndex)
        Else
            If Trim$(Map(GetPlayerMap(MyIndex)).Music) <> "Aucune" Then Call PlayMidi(Trim$(Map(GetPlayerMap(MyIndex)).Music)) Else Call StopMidi
            If OldMap <= 0 Then OldMap = GetPlayerMap(MyIndex)
        End If
        GettingMap = False
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    If (LCase$(Parse(0)) = "saymsg") Or (LCase$(Parse(0)) = "broadcastmsg") Or (LCase$(Parse(0)) = "globalmsg") Or (LCase$(Parse(0)) = "playermsg") Or (LCase$(Parse(0)) = "mapmsg") Or (LCase$(Parse(0)) = "adminmsg") Then
        If Len(Parse(1)) > 50 Then
            For i = 0 To ((Len(Parse(1)) \ 50))
                If i > 0 Then Call AddText(Mid$(Parse(1), (50 * i) + 1, 50), Val(Parse(2))) Else Call AddText(Mid$(Parse(1), 1, 50), Val(Parse(2)))
            Next i
        Else
            Call AddText(Parse(1), Val(Parse(2)))
        End If
        Exit Sub
    End If
    
    If (LCase$(Parse(0)) = "bank") Then
        frmbank.Show
        frmMirage.TxtQ2.Text = Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).String1
        frmMirage.txtQ.Visible = True
        Exit Sub
    End If
    
    If (LCase$(Parse(0)) = "craft") Then
        RecetteSelect = Parse(1)
        If Player(MyIndex).Metier > 0 Then
            If Metier(Player(MyIndex).Metier).Type = METIER_CRAFT Then
                frmcraft.Show
                frmcraft.lblMetierNom.Caption = Metier(RecetteSelect).nom
                frmcraft.lblNom.Caption = recette(Metier(RecetteSelect).data(0, 0)).nom
                frmcraft.scrlRecettes.Value = 0
                If Metier(RecetteSelect).data(0, 0) > 0 Then
                n = Metier(RecetteSelect).data(0, 0)
                    For i = 0 To 9
                        If recette(n).InCraft(i, 0) > 0 Then
                            frmcraft.lblNeedItem(i).Caption = Item(recette(n).InCraft(i, 0)).name & " (*" & recette(n).InCraft(i, 1) & ")"
                        Else
                            frmcraft.lblNeedItem(i).Caption = "Pas d'objet"
                        End If
                    Next i
                End If
                If recette(n).craft(0) > 0 Then
                    frmcraft.lblObtenu.Caption = Item(recette(n).craft(0)).name & " (*" & recette(n).craft(1) & ")"
                Else
                    frmcraft.lblObtenu.Caption = "Pas de craft"
                End If
            End If
        End If
        Exit Sub
    End If
    
    If (LCase$(Parse(0)) = "newmetier") Then
        i = MsgBox("Voulez vous apprendre se métier? " & Metier(Val(Parse(1))).nom, vbYesNo, GAME_NAME)
        If i = vbYes Then SendData ("newmetier" & SEP_CHAR & Val(Parse(1)) & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    If (LCase$(Parse(0)) = "remplacemetier") Then
        i = MsgBox("Voulez vous oublier votre métier et apprendre se métier? " & Metier(Val(Parse(1))).nom, vbYesNo, GAME_NAME)
        If i = vbYes Then SendData ("remplacemetier" & SEP_CHAR & Val(Parse(1)) & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    If (LCase$(Parse(0)) = "qmsg") Then
        frmMirage.txtQ.Visible = True
        frmMirage.TxtQ2.Text = Parse(1)
        Exit Sub
    End If
    
    If (LCase$(Parse(0)) = "lance") Then Call ShellExecute(frmMirage.hwnd, "open", Parse(1), "", App.Path, 1): Exit Sub
    
    ' ::::::::::::
    ' :: Guilde ::
    ' ::::::::::::
    If (LCase$(Parse(0)) = "guildtraineevbyesno") Then
        i = MsgBox("Voulez vous rentré dans la guilde? " & GetPlayerGuild(Val(Parse(2))), vbYesNo, GAME_NAME)
        If i = vbYes Then SendData ("guildtrainee" & SEP_CHAR & Parse(1) & SEP_CHAR & Val(Parse(2)) & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Item spawn packet ::
    ' :::::::::::::::::::::::
    If LCase$(Parse(0)) = "spawnitem" Then
        n = Val(Parse(1))
        
        MapItem(n).num = Val(Parse(2))
        MapItem(n).Value = Val(Parse(3))
        MapItem(n).dur = Val(Parse(4))
        MapItem(n).x = Val(Parse(5))
        MapItem(n).y = Val(Parse(6))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update item packet ::
    ' ::::::::::::::::::::::::
    If (LCase$(Parse(0)) = "updateitem") Then
        n = Val(Parse(1))
        
        ' Update the item
        With Item(n)
            .name = Parse(2)
            .Pic = Val(Parse(3))
            .Type = Val(Parse(4))
            .Data1 = Val(Parse(5))
            .Data2 = Val(Parse(6))
            .Data3 = Val(Parse(7))
            .StrReq = Val(Parse(8))
            .DefReq = Val(Parse(9))
            .SpeedReq = Val(Parse(10))
            .ClassReq = Val(Parse(11))
            .AccessReq = Val(Parse(12))
        
            .AddHP = Val(Parse(13))
            .AddMP = Val(Parse(14))
            .AddSP = Val(Parse(15))
            .AddStr = Val(Parse(16))
            .AddDef = Val(Parse(17))
            .AddMagi = Val(Parse(18))
            .AddSpeed = Val(Parse(19))
            .AddEXP = Val(Parse(20))
            .desc = Parse(21)
            .AttackSpeed = Val(Parse(22))
        
            .NCoul = Val(Parse(23))
            
            .paperdoll = Val(Parse(24))
            .paperdollPic = Val(Parse(25))
            .Empilable = Val(Parse(26))
            .tArme = Val(Parse(27))
        End With
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
    If LCase$(Parse(0)) = "spawnnpc" Then
        n = Val(Parse(1))
        
        With MapNpc(n)
            .num = Val(Parse(2))
            .x = Val(Parse(3))
            .y = Val(Parse(4))
            .Dir = Val(Parse(5))
        
            ' Client use only
            .XOffset = 0
            .YOffset = 0
            .Moving = 0
        End With
        Exit Sub
    End If
        
    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    If LCase$(Parse(0)) = "npcdead" Then
        n = Val(Parse(1))
        
        With MapNpc(n)
            .num = 0
            .x = 0
            .y = 0
            .Dir = 0
        
            ' Client use only
            .XOffset = 0
            .YOffset = 0
            .Moving = 0
        End With
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Update npc packet ::
    ' :::::::::::::::::::::::
    If (LCase$(Parse(0)) = "updatenpc") Then
        n = Val(Parse(1))
        
        With Npc(n)
        ' Update the item
            .name = Parse(2)
            .AttackSay = vbNullString
            .sprite = Val(Parse(3))
            .SpawnSecs = 0
            .Behavior = Val(Parse(6))
            .Range = 0
        For i = 1 To MAX_NPC_DROPS
            With .ItemNPC(i)
                .Chance = 0
                .ItemNum = 0
                .ItemValue = 0
            End With
        Next i
            .STR = 0
            .DEF = 0
            .speed = 0
            .MAGI = 0
            .MaxHp = Val(Parse(4))
            .exp = 0
            .QueteNum = Val(Parse(5))
            .Inv = Val(Parse(7))
            .Vol = Val(Parse(8))
        End With
        Exit Sub
    End If

    ' ::::::::::::::::::::
    ' :: Map key packet ::
    ' ::::::::::::::::::::
    If (LCase$(Parse(0)) = "mapkey") Then
        x = Val(Parse(1))
        y = Val(Parse(2))
        n = Val(Parse(3))
                
        TempTile(x, y).DoorOpen = n
        
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::::
    ' :: Update shop packet ::
    ' ::::::::::::::::::::::::
    If (LCase$(Parse(0)) = "updateshop") Then
        n = Val(Parse(1))
        
        ' Update the shop name
        Shop(n).name = Parse(2)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' ::    quete packet    ::
    ' ::::::::::::::::::::::::
    If (LCase$(Parse(0)) = "setquetecour") Then
        Player(MyIndex).QueteEnCour = Val(Parse(1))
        Call SendData("DEMAREQUETE" & SEP_CHAR & Player(MyIndex).QueteEnCour & SEP_CHAR & END_CHAR)
    End If
    
    If (LCase$(Parse(0)) = "quetecour") Then
        Player(MyIndex).QueteEnCour = Val(Parse(1))
        If Val(Parse(1)) = 0 Then Accepter = False
        Exit Sub
    End If
    
    If (LCase$(Parse(0)) = "playerquete") Then
        With Player(MyIndex)
            .QueteEnCour = Val(Parse(1))
            .Quetep.Data1 = Val(Parse(2))
            .Quetep.Data2 = Val(Parse(3))
            .Quetep.Data3 = Val(Parse(4))
            .Quetep.String1 = Val(Parse(5))
            n = 5
        
            For i = 1 To 15
                With .Quetep.indexe(i)
                    n = n + 1
                    .Data1 = Val(Parse(n))
                    n = n + 1
                    .Data2 = Val(Parse(n))
                    n = n + 1
                    .Data3 = Val(Parse(n))
                    n = n + 1
                    .String1 = Val(Parse(n))
                End With
            Next i
        End With
       
        Exit Sub
    End If
    
    If (LCase$(Parse(0)) = "finquete") Then
        Call MsgBox("Bravo!! vous venez de finir la quete : " & Trim$(quete(Player(MyIndex).QueteEnCour).nom) & " retourner voir celui qui vous la donnez pour avoir vos récompenses")
        frmMirage.quetetimersec.Enabled = False
        frmMirage.tmpsquete.Visible = False
        Exit Sub
    End If
    
    If (LCase$(Parse(0)) = "terminequete") Then Call ClearPlayerQuete(MyIndex): Exit Sub
    
    If (LCase$(Parse(0)) = "tempsquete") Then
        If Val(Parse(1)) > 0 Then
            frmMirage.quetetimersec.Interval = 1000
            Seco = Val(Parse(1)) - ((Val(Parse(1)) \ 60) * 60)
            Minu = Val(Parse(1)) \ 60
            frmMirage.tmpsquete.Visible = True
            If Len(STR$(Minu)) > 2 Then frmMirage.minute.Caption = Minu & ":" Else frmMirage.minute.Caption = "0" & Minu & ":"
            If Len(STR$(Seco)) > 2 Then frmMirage.seconde.Caption = Seco Else frmMirage.seconde.Caption = "0" & Seco
            frmMirage.quetetimersec.Enabled = True
            Exit Sub
        End If
        Exit Sub
    End If
    
    If (LCase$(Parse(0)) = "tuerquete") Then
        If Player(MyIndex).QueteEnCour <= 0 Then Exit Sub
        frmMirage.qt.Caption = "Monstres tuer :"
        n = 0
        For i = 1 To 15
            n = n + quete(Player(MyIndex).QueteEnCour).indexe(i).Data2
        Next i
        If frmMirage.av.Caption = " " Then frmMirage.av.Caption = "1/" & n Else frmMirage.av.Caption = Val(Mid(frmMirage.av.Caption, 1, 1)) + 1 & "/" & n
    End If
    
    If (LCase$(Parse(0)) = "xpquete") Then
        n = Val(Parse(1))
        If Player(MyIndex).QueteEnCour <= 0 Then Exit Sub
        frmMirage.qt.Caption = "Points gagnés :"
        If n > Val(quete(Player(MyIndex).QueteEnCour).Data1) Then n = Val(quete(Player(MyIndex).QueteEnCour).Data1)
        frmMirage.av.Caption = n & "/" & quete(Player(MyIndex).QueteEnCour).Data1
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update quete packet::
    ' ::::::::::::::::::::::::
    If (LCase$(Parse(0)) = "updatequete") Then
        n = Val(Parse(1))
        With quete(n)
        ' Update the quete
            .nom = Parse(2)
            .Data1 = Val(Parse(3))
            .Data2 = Val(Parse(4))
            .Data3 = Val(Parse(5))
            .description = Parse(6)
            .reponse = Parse(7)
            .String1 = Parse(8)
            .Temps = Val(Parse(9))
            .Type = Val(Parse(10))
        
        Dim l As Long
        i = 10
        For l = 1 To 15
            With .indexe(l)
                i = i + 1
                .Data1 = Val(Parse(i))
                i = i + 1
                .Data2 = Val(Parse(i))
                i = i + 1
                .Data3 = Val(Parse(i))
                i = i + 1
                .String1 = Parse(i)
            End With
        Next l
            .Recompence.exp = Val(Parse(i + 1))
            .Recompence.objn1 = Val(Parse(i + 2))
            .Recompence.objn2 = Val(Parse(i + 3))
            .Recompence.objn3 = Val(Parse(i + 4))
            .Recompence.objq1 = Val(Parse(i + 5))
            .Recompence.objq2 = Val(Parse(i + 6))
            .Recompence.objq3 = Val(Parse(i + 7))
            .Case = Val(Parse(i + 8))
        End With
        Exit Sub
    End If

    
    ' ::::::::::::::::::::::::
    ' :: Update spell packet ::
    ' ::::::::::::::::::::::::
    If (LCase$(Parse(0)) = "updatespell") Then
        n = Val(Parse(1))
        
        ' Update the spell name
        Spell(n).name = Parse(2)
        Spell(n).Big = Parse(3)
        Spell(n).SpellIco = Parse(4)
        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    If (LCase$(Parse(0)) = "trade") Then
        ShopNum = Val(Parse(1))
        If Val(Parse(2)) = 1 And Val(Parse(3)) > 0 Then frmTrade.picFixItems.Tag = Val(Parse(3)) Else frmTrade.picFixItems.Tag = 0
        n = 4
        For z = 1 To 6
            For i = 1 To MAX_TRADES
                GiveItem = Val(Parse(n))
                GiveValue = Val(Parse(n + 1))
                GetItem = Val(Parse(n + 2))
                GetValue = Val(Parse(n + 3))
                
                With Trade(z).Items(i)
                    .ItemGetNum = GetItem
                    .ItemGiveNum = GiveItem
                    .ItemGetVal = GetValue
                    .ItemGiveVal = GiveValue
                End With
                
                n = n + 4
            Next i
        Next z
        
        Dim xx As Long
        For xx = 1 To 6
            Trade(xx).Selected = NO
        Next xx
        
        Trade(1).Selected = YES
                    
        frmTrade.shopType.Top = frmTrade.label(1).Top
        frmTrade.shopType.Left = frmTrade.label(1).Left
        frmTrade.shopType.Height = frmTrade.label(1).Height
        frmTrade.shopType.Width = frmTrade.label(1).Width
        Trade(1).SelectedItem = 1
        
        NumShop = ShopNum
        
        Call ItemSelected(1, 1)
            
        frmTrade.Show vbModeless, frmMirage
        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If (LCase$(Parse(0)) = "spells") Then
        
        frmMirage.picPlayerSpells.Visible = True
        
        ' Put spells known in player record
        For i = 1 To MAX_PLAYER_SPELLS
            Player(MyIndex).Spell(i) = Val(Parse(i))
        Next i
        
        Call Affspell
        Call initRac
    End If

    ' ::::::::::::::::::::
    ' :: Weather packet ::
    ' ::::::::::::::::::::
    If (LCase$(Parse(0)) = "weather") Then
        If Val(Parse(1)) = WEATHER_RAINING And GameWeather <> WEATHER_RAINING Then Call AddText("La pluie commence à tomber!", BrightGreen)
        If Val(Parse(1)) = WEATHER_THUNDER And GameWeather <> WEATHER_THUNDER Then Call AddText("Le tonnerre commence à gronder!", BrightGreen)
        If Val(Parse(1)) = WEATHER_SNOWING And GameWeather <> WEATHER_SNOWING Then Call AddText("La neige commence à tomber", BrightGreen)
                
        If Val(Parse(1)) = WEATHER_NONE Then
            If GameWeather = WEATHER_RAINING Then
                Call AddText("La pluie se calme.", BrightGreen)
            ElseIf GameWeather = WEATHER_SNOWING Then
                Call AddText("La neige se calme.", BrightGreen)
            ElseIf GameWeather = WEATHER_THUNDER Then
                Call AddText("Les éclaires commence à disparaitre.", BrightGreen)
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
    'If (LCase$(Parse(0)) = "time") Then GameTime = Val(Parse(1)): Call InitNightAndFog(GetPlayerMap(MyIndex)): If GameTime = TIME_DAY Then Call AddText("Le jour se lève sur ce Royaume.", White) Else Call AddText("La nuit tombe dans ce Royaume.", White): Exit Sub
    
    ' ::::::::::::::::::::::::::
    ' :: Get Online List ::
    ' ::::::::::::::::::::::::::
    If LCase$(Parse(0)) = "onlinelist" Then
        frmMirage.lstOnline.Clear
    
        n = 2
        z = Val(Parse(1))
        For x = n To (z + 1)
            frmMirage.lstOnline.AddItem Trim$(Parse(n))
            n = n + 2
        Next x
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Blit Player Damage ::
    ' ::::::::::::::::::::::::
    If LCase$(Parse(0)) = "blitplayerdmg" Then
        DmgDamage = Val(Parse(1))
        NPCWho = Val(Parse(2))
        DmgAddRem = Val(Parse(3))
        DmgTime = GetTickCount
        iii = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Blit NPC Damage ::
    ' :::::::::::::::::::::
    If LCase$(Parse(0)) = "blitnpcdmg" Then
        NPCDmgDamage = Val(Parse(1))
        NPCDmgAddRem = Val(Parse(2))
        NPCDmgTime = GetTickCount
        ii = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::
    ' :: Retrieve the player's inventory ::
    ' :::::::::::::::::::::::::::::::::::::
    If LCase$(Parse(0)) = "pptrading" Then
        frmPlayerTrade.Items1.Clear
        frmPlayerTrade.Items2.Clear
        For i = 1 To MAX_PLAYER_TRADES
            Trading(i).InvNum = 0
            Trading(i).InvName = vbNullString
            Trading2(i).InvNum = 0
            Trading2(i).InvName = vbNullString
            frmPlayerTrade.Items1.AddItem i & ": <Aucun>"
            frmPlayerTrade.Items2.AddItem i & ": <Aucun>"
        Next i
        
        frmPlayerTrade.Items1.ListIndex = 0
        frmPlayerTrade.Etat.Caption = "En cours..."
        frmPlayerTrade.Etat.ForeColor = &H0&
        
        Call UpdateTradeInventory
        frmPlayerTrade.Show vbModeless, frmMirage
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If LCase$(Parse(0)) = "qtrade" Then
        For i = 1 To MAX_PLAYER_TRADES
            Trading(i).InvNum = 0
            Trading(i).InvName = vbNullString
            Trading2(i).InvNum = 0
            Trading2(i).InvName = vbNullString
        Next i
        frmPlayerTrade.Etat.Caption = "Refuser"
        frmPlayerTrade.Etat.ForeColor = &HFF&
        
        frmPlayerTrade.Visible = False
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If LCase$(Parse(0)) = "updatetradeitem" Then
            n = Val(Parse(1))
            
            Trading2(n).InvNum = Val(Parse(2))
            Trading2(n).InvName = Parse(3)
            Trading2(n).InvVal = Val(Parse(4))
            If Val(Trading2(n).InvNum) <= 0 Then frmPlayerTrade.Items2.List(n - 1) = n & ": <Aucun>" Else If Val(Trading2(n).InvVal) > 0 Then frmPlayerTrade.Items2.List(n - 1) = n & " : " & Trim$(Trading2(n).InvName) & "(" & Trading2(n).InvVal & ")" Else frmPlayerTrade.Items2.List(n - 1) = n & " : " & Trim$(Trading2(n).InvName)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If LCase$(Parse(0)) = "trading" Then n = Val(Parse(1)): If n = 1 Then frmPlayerTrade.Etat.Caption = "Accepter": frmPlayerTrade.Etat.ForeColor = &HFF00& Else frmPlayerTrade.Etat.Caption = "En cours...": frmPlayerTrade.Etat.ForeColor = 0: Exit Sub
    
' :::::::::::::::::::::::::
' :: Chat System Packets ::
' :::::::::::::::::::::::::
    If LCase$(Parse(0)) = "ppchatting" Then
        Call TransRegion(frmPlayerChat, 200, &H8000000F)
        frmPlayerChat.txtChat.Text = vbNullString
        frmPlayerChat.txtSay.Text = vbNullString
        frmPlayerChat.Label1.Caption = Trim$(Player(Val(Parse(1))).name)

        frmPlayerChat.Show vbModeless, frmMirage
        Exit Sub
    End If
    
    If LCase$(Parse(0)) = "qchat" Then
        frmPlayerChat.txtChat.Text = vbNullString
        frmPlayerChat.txtSay.Text = vbNullString
        frmPlayerChat.Visible = False
        Exit Sub
    End If
    
    If LCase$(Parse(0)) = "sendchat" Then
        Dim s As String
  
        s = vbNewLine & GetPlayerName(Val(Parse(2))) & "> " & Trim$(Parse(1))
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
    If LCase$(Parse(0)) = "sound" Then
        s = LCase$(Parse(1))
        Select Case Trim$(s)
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
                Call PlaySound(Trim$(Parse(2)))
        End Select
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: Sprite Change Confirmation Packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If LCase$(Parse(0)) = "spritechange" Then
        If Val(Parse(1)) = 1 Then
            i = MsgBox("Êtes-vous sur de vouloir acheter ce sprite?", 4, "Achat de Sprite")
            If i = 6 Then Call SendData("buysprite" & SEP_CHAR & END_CHAR)
        Else
            Call SendData("buysprite" & SEP_CHAR & END_CHAR)
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::::
    ' :: Change Player Direction Packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If LCase$(Parse(0)) = "changedir" Then Player(Val(Parse(2))).Dir = Val(Parse(1)): Exit Sub
    
    ' ::::::::::::::::::::::::::::::
    ' :: Flash Movie Event Packet ::
    ' ::::::::::::::::::::::::::::::
    If LCase$(Parse(0)) = "flashevent" Then
        If LCase$(Mid$(Trim$(Parse(1)), 1, 7)) = "http://" Then
            WriteINI "CONFIG", "Music", 0, App.Path & "\Config\Account.ini"
            WriteINI "CONFIG", "Sound", 0, App.Path & "\Config\Account.ini"
            Call StopMidi
            Call StopSound
            frmFlash.Flash.LoadMovie 0, Trim$(Parse(1))
            frmFlash.Flash.Play
            frmFlash.Check.Enabled = True
            frmFlash.Show vbModeless, frmMirage
        ElseIf FileExiste("Flashs\" & Trim$(Parse(1))) = True Then
            WriteINI "CONFIG", "Music", 0, App.Path & "\Config\Account.ini"
            WriteINI "CONFIG", "Sound", 0, App.Path & "\Config\Account.ini"
            Call StopMidi
            Call StopSound
            frmFlash.Flash.LoadMovie 0, App.Path & "\Flashs\" & Trim$(Parse(1))
            frmFlash.Flash.Play
            frmFlash.Check.Enabled = True
            frmFlash.Show vbModeless, frmMirage
        End If
        Exit Sub
    End If
    

    
    ' :::::::::::::::::::
    ' :: Prompt Packet ::
    ' :::::::::::::::::::
    If LCase$(Parse(0)) = "prompt" Then i = MsgBox(Trim$(Parse(1)), vbYesNo): Call SendData("prompt" & SEP_CHAR & i & SEP_CHAR & Val(Parse(2)) & SEP_CHAR & END_CHAR): Exit Sub
    
    If (LCase$(Parse(0)) = "updateemoticon") Then
        n = Val(Parse(1))
        
        Emoticons(n).Command = Parse(2)
        Emoticons(n).Pic = Val(Parse(3))
        Exit Sub
    End If
    
    If (LCase$(Parse(0)) = "updateemoticon") Then
        n = Val(Parse(1))
        
        Emoticons(n).Command = Parse(2)
        Emoticons(n).Pic = Val(Parse(3))
        Exit Sub
    End If
    
    If (LCase$(Parse(0)) = "updatearrow") Then
        n = Val(Parse(1))
        
        Arrows(n).name = Parse(2)
        Arrows(n).Pic = Val(Parse(3))
        Arrows(n).Range = Val(Parse(4))
        Exit Sub
    End If

    
    If (LCase$(Parse(0)) = "updatearrow") Then
        n = Val(Parse(1))
        
        Arrows(n).name = Parse(2)
        Arrows(n).Pic = Val(Parse(3))
        Arrows(n).Range = Val(Parse(4))
        Exit Sub
    End If

    If (LCase$(Parse(0)) = "checkarrows") Then
        n = Val(Parse(1))
        z = Val(Parse(2))
        i = Val(Parse(3))
        
        For x = 1 To MAX_PLAYER_ARROWS
            With Player(n).Arrow(x)
                If .Arrow = 0 Then
                    .Arrow = 1
                    .ArrowNum = z
                    .ArrowAnim = Arrows(z).Pic
                    .ArrowTime = GetTickCount
                    .ArrowVarX = 0
                    .ArrowVarY = 0
                    .ArrowY = GetPlayerY(n)
                    .ArrowX = GetPlayerX(n)
                
                    If i = DIR_DOWN Then
                        .ArrowY = GetPlayerY(n) + 1
                        .ArrowPosition = 0
                    If .ArrowY - 1 > MAX_MAPY Then .Arrow = 0: Exit Sub
                    End If
                    If i = DIR_UP Then
                        .ArrowY = GetPlayerY(n) - 1
                        .ArrowPosition = 1
                    If .ArrowY + 1 < 0 Then .Arrow = 0: Exit Sub
                    End If
                    If i = DIR_RIGHT Then
                        .ArrowX = GetPlayerX(n) + 1
                        .ArrowPosition = 2
                    If .ArrowX - 1 > MAX_MAPX Then .Arrow = 0: Exit Sub
                    End If
                    If i = DIR_LEFT Then
                        .ArrowX = GetPlayerX(n) - 1
                        .ArrowPosition = 3
                    If .ArrowX + 1 < 0 Then .Arrow = 0: Exit Sub
                    End If
                    Exit For
                End If
            End With
        Next x
        Exit Sub
    End If

    If (LCase$(Parse(0)) = "checksprite") Then Player(Val(Parse(1))).sprite = Val(Parse(2)): Exit Sub
    
    If (LCase$(Parse(0)) = "mapreport") Then
        n = 1
        
        frmMapReport.lstIndex.Clear
        For i = 1 To MAX_MAPS
            frmMapReport.lstIndex.AddItem i & ": " & Trim$(Parse(n))
            n = n + 1
        Next i
        
        frmMapReport.Show vbModeless, frmMirage
        Exit Sub
    End If
    
    ' :::::::::::::::::
    ' :: Time packet ::
    ' :::::::::::::::::
    If (LCase$(Parse(0)) = "time") Then
        GameTime = Val(Parse(1))
        Call InitNightAndFog(GetPlayerMap(MyIndex))
        If GameTime = TIME_DAY Then Call AddText("Le jour se lève sur le Royaume.", White) Else Call AddText("La nuit tombe dans ce royaume..", White)
        Exit Sub
    End If
    
    If (LCase$(Parse(0)) = "bloodanim") Then
        Player(Val(Parse(1))).BloodAnim.Target = Val(Parse(2))
        Player(Val(Parse(1))).BloodAnim.TargetType = Val(Parse(3))
        Player(Val(Parse(1))).BloodAnim.SpellDone = Int(Rnd * Int(DDSD_Blood.lHeight / PIC_Y + 1))
        Player(Val(Parse(1))).BloodAnim.CastedSpell = YES
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Spell anim packet ::
    ' :::::::::::::::::::::::
    If (LCase$(Parse(0)) = "spellanim") Then
        Dim SpellNum As Long
        
        'Vérifier si le sort à une cible
        If Val(Parse(7)) <= -1 Then Exit Sub
        
        SpellNum = Val(Parse(1))
        Spell(SpellNum).SpellAnim = Val(Parse(2))
        Spell(SpellNum).SpellTime = Val(Parse(3))
        Spell(SpellNum).SpellDone = Val(Parse(4))
        
        Player(Val(Parse(5))).SpellNum = SpellNum
        
        For i = 1 To MAX_SPELL_ANIM
            With Player(Val(Parse(5))).SpellAnim(i)
                If .CastedSpell = NO Then
                    .SpellDone = 0
                    .SpellVar = 0
                    .SpellTime = GetTickCount
                    .TargetType = Val(Parse(6))
                    .Target = Val(Parse(7))
                    .CastedSpell = YES
                    Exit For
                End If
            End With
        Next i
        Exit Sub
    End If
    
    If (LCase$(Parse(0)) = "checkemoticons") Then
        n = Val(Parse(1))
        
        Player(n).EmoticonNum = Val(Parse(2))
        Player(n).EmoticonTime = GetTickCount
        Player(n).EmoticonVar = 0
        Exit Sub
    End If
    
    If LCase$(Parse(0)) = "levelup" Then Player(Val(Parse(1))).LevelUpT = GetTickCount: Player(Val(Parse(1))).LevelUp = 1: Exit Sub
    
    If LCase$(Parse(0)) = "damagedisplay" Then
        For i = 1 To MAX_BLT_LINE
            If Val(Parse(1)) = 0 Then
                If BattlePMsg(i).Index <= 0 Then
                    BattlePMsg(i).Index = 1
                    BattlePMsg(i).Msg = Parse(2)
                    BattlePMsg(i).Color = Val(Parse(3))
                    BattlePMsg(i).Time = GetTickCount
                    BattlePMsg(i).Done = 1
                    BattlePMsg(i).y = 0
                    Exit Sub
                Else
                    BattlePMsg(i).y = BattlePMsg(i).y - 15
                End If
            Else
                If BattleMMsg(i).Index <= 0 Then
                    BattleMMsg(i).Index = 1
                    BattleMMsg(i).Msg = Parse(2)
                    BattleMMsg(i).Color = Val(Parse(3))
                    BattleMMsg(i).Time = GetTickCount
                    BattleMMsg(i).Done = 1
                    BattleMMsg(i).y = 0
                    Exit Sub
                Else
                    BattleMMsg(i).y = BattleMMsg(i).y - 15
                End If
            End If
        Next i
        
        z = 1
        If Val(Parse(1)) = 0 Then
            For i = 1 To MAX_BLT_LINE
                If i < MAX_BLT_LINE Then If BattlePMsg(i).y < BattlePMsg(i + 1).y Then z = i Else If BattlePMsg(i).y < BattlePMsg(1).y Then z = i
            Next i
                        
            BattlePMsg(z).Index = 1
            BattlePMsg(z).Msg = Parse(2)
            BattlePMsg(z).Color = Val(Parse(3))
            BattlePMsg(z).Time = GetTickCount
            BattlePMsg(z).Done = 1
            BattlePMsg(z).y = 0
        Else
            For i = 1 To MAX_BLT_LINE
                If i < MAX_BLT_LINE Then If BattleMMsg(i).y < BattleMMsg(i + 1).y Then z = i Else If BattleMMsg(i).y < BattleMMsg(1).y Then z = i
            Next i
            BattleMMsg(z).Index = 1
            BattleMMsg(z).Msg = Parse(2)
            BattleMMsg(z).Color = Val(Parse(3))
            BattleMMsg(z).Time = GetTickCount
            BattleMMsg(z).Done = 1
            BattleMMsg(z).y = 0
        End If
        Exit Sub
    End If
    
    If LCase$(Parse(0)) = "itembreak" Then
        ItemDur(Val(Parse(1))).Item = Val(Parse(2))
        ItemDur(Val(Parse(1))).dur = Val(Parse(3))
        ItemDur(Val(Parse(1))).Done = 1
        Exit Sub
    End If
    
    If LCase$(Parse(0)) = "playeranimstart" Then
        PlayerAnim(Parse(1), 0) = Val(Parse(2)) + 1
        PlayerAnim(Parse(1), 1) = GetTickCount
        PlayerAnim(Parse(1), 2) = 0
        PlayerAnim(Parse(1), 3) = 0
        Exit Sub
    End If
    
    If LCase$(Parse(0)) = "playeranimstop" Then
        PlayerAnim(Parse(1), 0) = 0
        PlayerAnim(Parse(1), 1) = GetTickCount
        PlayerAnim(Parse(1), 2) = 0
        PlayerAnim(Parse(1), 3) = 0
        PlayerAnim(Parse(1), 4) = 0
        Exit Sub
    End If
    
    If LCase$(Parse(0)) = "playeranim" Then
        PlayerAnim(Parse(1), 0) = Val(Parse(2)) + 1
        PlayerAnim(Parse(1), 1) = GetTickCount
        PlayerAnim(Parse(1), 2) = 0
        PlayerAnim(Parse(1), 3) = GetTickCount + Val(Parse(3))
        PlayerAnim(Parse(1), 4) = 0
        Exit Sub
    End If
    
    If LCase$(Parse(0)) = "playeranimrt" Then
        PlayerAnim(Parse(1), 0) = Val(Parse(2)) + 1
        PlayerAnim(Parse(1), 1) = GetTickCount
        PlayerAnim(Parse(1), 2) = 0
        PlayerAnim(Parse(1), 3) = GetTickCount + Val(Parse(3))
        PlayerAnim(Parse(1), 4) = Val(Parse(4)) + 1
        Exit Sub
    End If
Exit Sub

erreur:
Call MsgBox("Une erreur de réception du serveur c'est produite(Numéros de l'erreur : " & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "). Si le probléme pérsiste veulliez contacter un administrateur.", vbCritical, "Erreur")
Call GameDestroy
End Sub

Function ConnectToServer() As Boolean
Dim Wait As Long

    ' Check to see if we are already connected, if so just exit
    If IsConnected Then ConnectToServer = True: Exit Function
    
    Wait = GetTickCount
    frmMirage.Socket.Close
    frmMirage.Socket.Connect
    
    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsConnected) And (GetTickCount <= Wait + 3000)
        DoEvents
        Sleep 1
    Loop
    
    If IsConnected Then ConnectToServer = True Else ConnectToServer = False
End Function

Function IsConnected() As Boolean
    If frmMirage.Socket.State = sckConnected Then IsConnected = True Else IsConnected = False
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If GetPlayerName(Index) <> vbNullString Then IsPlaying = True Else IsPlaying = False
End Function

Sub SendData(ByVal data As String)
    If IsConnected Then frmMirage.Socket.SendData data: DoEvents
End Sub

Sub SendNewAccount(ByVal name As String, ByVal Password As String)
Dim Packet As String

    Packet = "newfaccountied" & SEP_CHAR & Trim$(name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDelAccount(ByVal name As String, ByVal Password As String)
Dim Packet As String
    
    Packet = "delimaccounted" & SEP_CHAR & Trim$(name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendLogin(ByVal name As String, ByVal Password As String)
Dim Packet As String

    Packet = "logination" & SEP_CHAR & Trim$(name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & App.Major & SEP_CHAR & App.Minor & SEP_CHAR & App.Revision & SEP_CHAR & SEC_CODE1 & SEP_CHAR & SEC_CODE2 & SEP_CHAR & SEC_CODE3 & SEP_CHAR & SEC_CODE4 & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendAddChar(ByVal name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Slot As Long)
Dim Packet As String

    Packet = "addachara" & SEP_CHAR & Trim$(name) & SEP_CHAR & Sex & SEP_CHAR & ClassNum & SEP_CHAR & Slot & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDelChar(ByVal Slot As Long)
Dim Packet As String
    
    Packet = "delimbocharu" & SEP_CHAR & Slot & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendGetClasses()
Dim Packet As String

    Packet = "gatglasses" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendUseChar(ByVal CharSlot As Long)
Dim Packet As String

    Packet = "usagakarim" & SEP_CHAR & CharSlot & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SayMsg(ByVal Text As String)
Dim Packet As String

    Packet = "saymsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub GlobalMsg(ByVal Text As String)
Dim Packet As String

    Packet = "globalmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub BroadcastMsg(ByVal Text As String)
Dim Packet As String

    Packet = "broadcastmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub EmoteMsg(ByVal Text As String)
Dim Packet As String

    Packet = "emotemsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub GuildeMsg(ByVal Text As String)
Dim Packet As String

   Packet = "guildemsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
   Call SendData(Packet)
End Sub

Sub MapMsg(ByVal Text As String)
Dim Packet As String

    Packet = "mapmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
Dim Packet As String

    Packet = "playermsg" & SEP_CHAR & MsgTo & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub AdminMsg(ByVal Text As String)
Dim Packet As String

    Packet = "adminmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerMove()
Dim Packet As String

    Packet = "playermove" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & Player(MyIndex).Moving & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub Sendplayerdir()
Dim Packet As String

    Packet = "playerdir" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerRequestNewMap()
Dim Packet As String
    
    Packet = "requestnewmap" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub WarpMeTo(ByVal name As String)
Dim Packet As String

    OldMap = GetPlayerMap(MyIndex)
    Packet = "WARPMETO" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub WarpToMe(ByVal name As String)
Dim Packet As String

    OldMap = GetPlayerMap(MyIndex)
    Packet = "WARPTOME" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub WarpTo(ByVal MapNum As Long)
Dim Packet As String
    
    OldMap = GetPlayerMap(MyIndex)
    Packet = "WARPTO" & SEP_CHAR & MapNum & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetAccess(ByVal name As String, ByVal Access As Byte)
Dim Packet As String

    Packet = "SETACCESS" & SEP_CHAR & name & SEP_CHAR & Access & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetSprite(ByVal SpriteNum As Long)
Dim Packet As String

    Packet = "SETSPRITE" & SEP_CHAR & SpriteNum & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetName(ByVal nom As String)
Dim Packet As String

    Packet = "SETNAME" & SEP_CHAR & nom & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendGetStats()
Dim Packet As String

    Packet = "GETSTATS" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendGetOtherStats(ByVal name As String)
Dim Packet As String

    Packet = "GETOTHERSTATS" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerInfoRequest(ByVal name As String)
Dim Packet As String

    Packet = "PLAYERINFOREQUEST" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendKick(ByVal name As String)
Dim Packet As String

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

Sub SendSetPlayerstr(ByVal name As String, ByVal num As Long)
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
