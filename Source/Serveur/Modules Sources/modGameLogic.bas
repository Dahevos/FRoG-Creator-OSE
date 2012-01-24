Attribute VB_Name = "modGameLogic"
Option Explicit

Function GetPlayerDamage(ByVal index As Long) As Long
Dim WeaponSlot As Long

    GetPlayerDamage = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then Exit Function
            
    If GetPlayerPetSlot(index) > 0 Then
        GetPlayerDamage = ((GetPlayerStr(index) \ 2) + (Pets(item(GetPlayerInvItemNum(index, GetPlayerPetSlot(index))).data1).addForce) \ 2)
    Else
        GetPlayerDamage = (GetPlayerStr(index) \ 2)
    End If
    
    If GetPlayerDamage <= 0 Then GetPlayerDamage = 1
    
    If GetPlayerWeaponSlot(index) > 0 Then
        WeaponSlot = GetPlayerWeaponSlot(index)
        
        GetPlayerDamage = GetPlayerDamage + item(GetPlayerInvItemNum(index, WeaponSlot)).data2
        
        If GetPlayerInvItemDur(index, WeaponSlot) > -1 Then
            Call SetPlayerInvItemDur(index, WeaponSlot, GetPlayerInvItemDur(index, WeaponSlot) - 1)
        
            If GetPlayerInvItemDur(index, WeaponSlot) = 0 Then
                Call BattleMsg(index, "Ton " & Trim$(item(GetPlayerInvItemNum(index, WeaponSlot)).Name) & " a été brisé.", Yellow, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, WeaponSlot), 0)
            Else
                If GetPlayerInvItemDur(index, WeaponSlot) <= 10 Then Call BattleMsg(index, "Ton " & Trim$(item(GetPlayerInvItemNum(index, WeaponSlot)).Name) & " va bientôt se briser! Usure: " & GetPlayerInvItemDur(index, WeaponSlot) & "/" & Trim$(item(GetPlayerInvItemNum(index, WeaponSlot)).data1), Yellow, 0)
            End If
        End If
    End If
    
    If GetPlayerDamage < 0 Then GetPlayerDamage = 0
    
End Function

Function GetPlayerProtection(ByVal index As Long) As Long
Dim ArmorSlot As Long, HelmSlot As Long, ShieldSlot As Long
    
    GetPlayerProtection = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then Exit Function
        
    ArmorSlot = GetPlayerArmorSlot(index)
    HelmSlot = GetPlayerHelmetSlot(index)
    ShieldSlot = GetPlayerShieldSlot(index)
    If GetPlayerPetSlot(index) > 0 Then
        GetPlayerProtection = ((GetPlayerDEF(index) \ 4) + (Pets(item(GetPlayerInvItemNum(index, GetPlayerPetSlot(index))).data1).addDefence) \ 4)
    Else
        GetPlayerProtection = (GetPlayerDEF(index) \ 4)
    End If

    If ArmorSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + item(GetPlayerInvItemNum(index, ArmorSlot)).data2
        If GetPlayerInvItemDur(index, ArmorSlot) > -1 Then
            Call SetPlayerInvItemDur(index, ArmorSlot, GetPlayerInvItemDur(index, ArmorSlot) - 1)
        
            If GetPlayerInvItemDur(index, ArmorSlot) = 0 Then
                Call BattleMsg(index, "Ton " & Trim$(item(GetPlayerInvItemNum(index, ArmorSlot)).Name) & " a été brisé.", Yellow, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, ArmorSlot), 0)
            Else
                If GetPlayerInvItemDur(index, ArmorSlot) <= 10 Then Call BattleMsg(index, "Ton " & Trim$(item(GetPlayerInvItemNum(index, ArmorSlot)).Name) & " est endommagé! Usure: " & GetPlayerInvItemDur(index, ArmorSlot) & "/" & Trim$(item(GetPlayerInvItemNum(index, ArmorSlot)).data1), Yellow, 0)
            End If
        End If
    End If
    
    If HelmSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + item(GetPlayerInvItemNum(index, HelmSlot)).data2
        If GetPlayerInvItemDur(index, HelmSlot) > -1 Then
            Call SetPlayerInvItemDur(index, HelmSlot, GetPlayerInvItemDur(index, HelmSlot) - 1)

            If GetPlayerInvItemDur(index, HelmSlot) <= 0 Then
                Call BattleMsg(index, "Ton " & Trim$(item(GetPlayerInvItemNum(index, HelmSlot)).Name) & " a été brisé.", Yellow, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, HelmSlot), 0)
            Else
                If GetPlayerInvItemDur(index, HelmSlot) <= 10 Then Call BattleMsg(index, "Ton " & Trim$(item(GetPlayerInvItemNum(index, HelmSlot)).Name) & " est endommagé! Usure: " & GetPlayerInvItemDur(index, HelmSlot) & "/" & Trim$(item(GetPlayerInvItemNum(index, HelmSlot)).data1), Yellow, 0)
            End If
        End If
    End If
    
    If ShieldSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + item(GetPlayerInvItemNum(index, ShieldSlot)).data2
        If GetPlayerInvItemDur(index, ShieldSlot) > -1 Then
            Call SetPlayerInvItemDur(index, ShieldSlot, GetPlayerInvItemDur(index, ShieldSlot) - 1)

            If GetPlayerInvItemDur(index, ShieldSlot) <= 0 Then
                Call BattleMsg(index, "Ton " & Trim$(item(GetPlayerInvItemNum(index, ShieldSlot)).Name) & " est brisé.", Yellow, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, ShieldSlot), 0)
            Else
                If GetPlayerInvItemDur(index, ShieldSlot) <= 10 Then Call BattleMsg(index, "Ton " & Trim$(item(GetPlayerInvItemNum(index, ShieldSlot)).Name) & " est endommagé! Usure: " & GetPlayerInvItemDur(index, ShieldSlot) & "/" & Trim$(item(GetPlayerInvItemNum(index, ShieldSlot)).data1), Yellow, 0)
            End If
        End If
    End If
End Function

Function FindOpenPlayerSlot() As Long
Dim i As Long

    FindOpenPlayerSlot = 0
    
    For i = 1 To MAX_PLAYERS
        If Not IsConnected(i) Then FindOpenPlayerSlot = i: Exit Function
    Next i
End Function

Function FindOpenInvSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
Dim i As Long
    
    FindOpenInvSlot = 0
    
    ' Check for subscript out of range
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function
    
    If item(ItemNum).type = ITEM_TYPE_CURRENCY Or item(ItemNum).Empilable <> 0 Then
        ' If currency then check to see if they already have an guildSoloView of the item and add it to that
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = ItemNum Then FindOpenInvSlot = i: Exit Function
        Next i
    End If
    
    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, i) <= 0 Then FindOpenInvSlot = i: Exit Function
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

Function FindOpenSpellSlot(ByVal index As Long) As Long
Dim i As Long

    FindOpenSpellSlot = 0
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) = 0 Then FindOpenSpellSlot = i: Exit Function
    Next i
End Function

Function HasSpell(ByVal index As Long, ByVal SpellNum As Long) As Boolean
Dim i As Long

    HasSpell = False
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) = SpellNum Then HasSpell = True: Exit Function
    Next i
End Function

Function TotalOnlinePlayers() As Long
Dim i As Long
TotalOnlinePlayers = 0

For i = 1 To MAX_PLAYERS
    If IsPlaying(i) Then TotalOnlinePlayers = TotalOnlinePlayers + 1
Next i
End Function

Function FindPlayer(ByVal Name As String) As Long
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    FindPlayer = 0
End Function

Function HasItem(ByVal index As Long, ByVal ItemNum As Long) As Long
Dim i As Long
    
    HasItem = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function
        
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If item(ItemNum).type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(index, i)
            Else
                HasItem = 1
            End If
            Exit Function
        End If
    Next i
End Function

Sub TakeItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long, n As Long
Dim TakeItem As Boolean

    TakeItem = False
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If item(ItemNum).type = ITEM_TYPE_CURRENCY Or item(ItemNum).Empilable <> 0 Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    Call SendInventoryUpdate(index, i)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case item(GetPlayerInvItemNum(index, i)).type
                    Case ITEM_TYPE_WEAPON
                        If GetPlayerWeaponSlot(index) > 0 Then
                            If i = GetPlayerWeaponSlot(index) Then
                                Call SetPlayerWeaponSlot(index, 0)
                                Call SendInventory(index)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index)) Then TakeItem = True
                            End If
                        Else
                            TakeItem = True
                        End If
                
                    Case ITEM_TYPE_ARMOR
                        If GetPlayerArmorSlot(index) > 0 Then
                            If i = GetPlayerArmorSlot(index) Then
                                Call SetPlayerArmorSlot(index, 0)
                                Call SendInventory(index)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerArmorSlot(index)) Then TakeItem = True
                            End If
                        Else
                            TakeItem = True
                        End If
                    
                    Case ITEM_TYPE_HELMET
                        If GetPlayerHelmetSlot(index) > 0 Then
                            If i = GetPlayerHelmetSlot(index) Then
                                Call SetPlayerHelmetSlot(index, 0)
                                Call SendInventory(index)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index)) Then TakeItem = True
                            End If
                        Else
                            TakeItem = True
                        End If
                    
                    Case ITEM_TYPE_SHIELD
                        If GetPlayerShieldSlot(index) > 0 Then
                            If i = GetPlayerShieldSlot(index) Then
                                Call SetPlayerShieldSlot(index, 0)
                                Call SendInventory(index)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerShieldSlot(index)) Then TakeItem = True
                            End If
                        Else
                            TakeItem = True
                        End If
                    Case ITEM_TYPE_PET
                        If GetPlayerPetSlot(index) > 0 Then
                            If i = GetPlayerPetSlot(index) Then
                                Call SetPlayerPetSlot(index, 0)
                                Call SendInventory(index)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerPetSlot(index)) Then TakeItem = True
                            End If
                        Else
                            TakeItem = True
                        End If
                End Select
                
                

                
                n = item(GetPlayerInvItemNum(index, i)).type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (n <> ITEM_TYPE_WEAPON) And (n <> ITEM_TYPE_ARMOR) And (n <> ITEM_TYPE_HELMET) And (n <> ITEM_TYPE_SHIELD) And (n <> ITEM_TYPE_PET) Then TakeItem = True
            End If
                            
            If TakeItem = True Then
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                Call SetPlayerInvItemDur(index, i, 0)
                
                ' Send the inventory update
                Call SendInventoryUpdate(index, i)
                Exit Sub
            End If
        End If
    Next i
End Sub

Sub GiveItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    i = FindOpenInvSlot(index, ItemNum)
    
    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(index, i, ItemNum)
        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)
        
        If (item(ItemNum).type = ITEM_TYPE_ARMOR) Or (item(ItemNum).type = ITEM_TYPE_WEAPON) Or (item(ItemNum).type = ITEM_TYPE_HELMET) Or (item(ItemNum).type = ITEM_TYPE_SHIELD) Then
            Call SetPlayerInvItemDur(index, i, item(ItemNum).data1)
        End If
        
        Call SendInventoryUpdate(index, i)
    Else
        Call PlayerMsg(index, "Votre inventaire est plein.", BrightRed)
    End If
End Sub

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
        
    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    
    Call SpawnItemSlot(i, ItemNum, ItemVal, item(ItemNum).data1, MapNum, X, Y)
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

Sub SpawnAllMapsItems()
Dim i As Long
    
    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next i
End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
Dim X As Long
Dim Y As Long
Dim i As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
        
    ' Spawn what we have
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            ' Check if the tile type is an item or a saved tile incase someone drops something

            If (Map(MapNum).Tile(X, Y).type = TILE_TYPE_ITEM) Then
                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If (item(Map(MapNum).Tile(X, Y).data1).type = ITEM_TYPE_CURRENCY Or item(Map(MapNum).Tile(X, Y).data1).Empilable <> 0) And Map(MapNum).Tile(X, Y).data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).Tile(X, Y).data1, 1, MapNum, X, Y)
                Else
                    Call SpawnItem(Map(MapNum).Tile(X, Y).data1, Map(MapNum).Tile(X, Y).data2, MapNum, X, Y)
                End If
            End If
        Next X
    Next Y
End Sub

Sub PlayerMapGetItem(ByVal index As Long)
Dim i As Long
Dim n As Long
Dim MapNum As Long
Dim Msg As String


    If IsPlaying(index) = False Then Exit Sub
    
    MapNum = GetPlayerMap(index)
    
    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(MapNum, i).num > 0) And (MapItem(MapNum, i).num <= MAX_ITEMS) Then
            ' Check if item is at the same location as the player
            If (MapItem(MapNum, i).X = GetPlayerX(index)) And (MapItem(MapNum, i).Y = GetPlayerY(index)) Then
                ' Find open slot
                n = FindOpenInvSlot(index, MapItem(MapNum, i).num)
                               
                ' Open slot available?
                If n <> 0 Then
                    ' Set item in players inventor
                    Call SetPlayerInvItemNum(index, n, MapItem(MapNum, i).num)
                    
                    If item(GetPlayerInvItemNum(index, n)).type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(index, n)).Empilable <> 0 Then
                        Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(MapNum, i).value)
                        Msg = "Vous ramassez " & MapItem(MapNum, i).value & " " & Trim$(item(GetPlayerInvItemNum(index, n)).Name) & "."
                    Else
                        Call SetPlayerInvItemValue(index, n, 1)
                        Msg = "Vous ramassez un " & Trim$(item(GetPlayerInvItemNum(index, n)).Name) & "."
                    End If
                    
                    If Player(index).Char(Player(index).CharNum).QueteEnCour > 0 Then
                        If quete(Player(index).Char(Player(index).CharNum).QueteEnCour).type = QUETE_TYPE_RECUP Then
                            Call PlayerQueteTypeRecup(index, Player(index).Char(Player(index).CharNum).QueteEnCour, GetPlayerInvItemNum(index, n), GetPlayerInvItemValue(index, n))
                        End If
                    End If
                                            
                    Call SetPlayerInvItemDur(index, n, MapItem(MapNum, i).Dur)
                        
                    ' Erase item from the map
                    MapItem(MapNum, i).num = 0
                    MapItem(MapNum, i).value = 0
                    MapItem(MapNum, i).Dur = 0
                    MapItem(MapNum, i).X = 0
                    MapItem(MapNum, i).Y = 0
                        
                    Call SendInventoryUpdate(index, n)
                    Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                    Call PlayerMsg(index, Msg, Yellow)
                    Exit Sub
                Else
                    Call PlayerMsg(index, "Votre inventaire est plein.", BrightRed)
                    Exit Sub
                End If
            End If
        End If
        
    Next i
End Sub

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
                    If InvNum = GetPlayerArmorSlot(index) Then
                        Call SetPlayerArmorSlot(index, 0)
                        Call SendInventory(index)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                
                Case ITEM_TYPE_WEAPON
                    If InvNum = GetPlayerWeaponSlot(index) Then
                        Call SetPlayerWeaponSlot(index, 0)
                        Call SendInventory(index)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                    
                Case ITEM_TYPE_HELMET
                    If InvNum = GetPlayerHelmetSlot(index) Then
                        Call SetPlayerHelmetSlot(index, 0)
                        Call SendInventory(index)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                                    
                Case ITEM_TYPE_SHIELD
                    If InvNum = GetPlayerShieldSlot(index) Then
                        Call SetPlayerShieldSlot(index, 0)
                        Call SendInventory(index)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                 
                Case ITEM_TYPE_PET
                    If InvNum = GetPlayerPetSlot(index) Then
                        Call SetPlayerPetSlot(index, 0)
                        Call SendInventory(index)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                    
                Case ITEM_TYPE_MONTURE
                    If InvNum = GetPlayerArmorSlot(index) Then
                        Dim s As Long
                        Call SetPlayerArmorSlot(index, 0)
                        Call SendInventory(index)
                        Call SendWornEquipment(index)
                        s = Val(GetVar(App.Path & "\accounts\" & Trim$(Player(index).Login) & ".ini", "CHAR" & Player(index).CharNum, "monture"))
                        Call SetPlayerSprite(index, s)
                        Call SendPlayerData(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
            End Select
                                
            MapItem(GetPlayerMap(index), i).num = GetPlayerInvItemNum(index, InvNum)
            MapItem(GetPlayerMap(index), i).X = GetPlayerX(index)
            MapItem(GetPlayerMap(index), i).Y = GetPlayerY(index)
                        
            If item(GetPlayerInvItemNum(index, InvNum)).type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(index, InvNum)).Empilable <> 0 Then
                ' Check if its more then they have and if so drop it all
                If Amount >= GetPlayerInvItemValue(index, InvNum) Then
                    MapItem(GetPlayerMap(index), i).value = GetPlayerInvItemValue(index, InvNum)
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " dépose un " & GetPlayerInvItemValue(index, InvNum) & " " & Trim$(item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemNum(index, InvNum, 0)
                    Call SetPlayerInvItemValue(index, InvNum, 0)
                    Call SetPlayerInvItemDur(index, InvNum, 0)
                Else
                    MapItem(GetPlayerMap(index), i).value = Amount
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " dépose un " & Amount & " " & Trim$(item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
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
            Call PlayerMsg(index, "Il y a trop d'objets par terre !", BrightRed)
        End If
    End If
End Sub

'Thanks to Gotakk 'cause he's the one who tracked a bug (after almost a year, he's the only one)
'The bug was a forgot to check if we already informed the player about the QuestStatus

Sub PlayerQueteTypeRecup(ByVal index As Long, ByVal Queteec As Long, ByVal Objnum As Long, ByVal Objvalue As Long)
Dim i As Long
Dim n As Long
Dim z As Long

If Not IsPlaying(index) Then Exit Sub
If Queteec <= 0 Then Exit Sub
If Objnum <= 0 Or Objnum > MAX_ITEMS Or Objvalue < 0 Then Exit Sub
If GetPlayerQueteEtat(index, Queteec) Then Exit Sub

For i = 1 To 15
    n = AObjet(index, quete(Queteec).indexe(i).data1)
    If n > 0 Then
        Player(index).Char(Player(index).CharNum).Quetep.indexe(i).data1 = 1
        Player(index).Char(Player(index).CharNum).Quetep.indexe(i).data2 = NbObjet(index, quete(Queteec).indexe(i).data1)
    End If

    If quete(Queteec).indexe(i).data2 <= 0 Then
        Player(index).Char(Player(index).CharNum).Quetep.indexe(i).data1 = 1
        Player(index).Char(Player(index).CharNum).Quetep.indexe(i).data2 = 1
    End If
Next i
    
n = 0
For i = 1 To 15
    If Player(index).Char(Player(index).CharNum).Quetep.indexe(i).data1 = 1 And Player(index).Char(Player(index).CharNum).Quetep.indexe(i).data2 >= Val(quete(Queteec).indexe(i).data2) Then
        n = n + 1
    End If
Next i

If n = 15 And Player(index).Char(Player(index).CharNum).QueteStatut(Queteec) = 0 Then
    Player(index).Char(Player(index).CharNum).QueteStatut(Queteec) = 1
    Call SendDataTo(index, "FINQUETE" & SEP_CHAR & END_CHAR)
End If

End Sub

Sub PlayerQueteTypeAport(ByVal index As Long, ByVal Queteec As Long)
Dim i As Long
Dim n As Long
n = 0

If Not IsPlaying(index) Then Exit Sub
If Queteec <= 0 Then Exit Sub
If quete(Queteec).data1 <= 0 Or quete(Queteec).data1 > MAX_ITEMS Then Exit Sub
If GetPlayerQueteEtat(index, Queteec) Then Exit Sub

For i = 1 To 24
    If Player(index).Char(Player(index).CharNum).Inv(i).num = quete(Queteec).data1 Then
        Call SetPlayerInvItemNum(index, i, 0)
        Call SetPlayerInvItemValue(index, i, 0)
        Call SetPlayerInvItemDur(index, i, 0)
        Call SendInventory(index)
        n = 1
        Exit For
    End If
Next i

If n = 1 And Player(index).Char(Player(index).CharNum).QueteStatut(Queteec) = 0 Then
    Call QueteMsg(index, Trim$(quete(Player(index).Char(Player(index).CharNum).QueteEnCour).String1))
    Player(index).Char(Player(index).CharNum).QueteStatut(Queteec) = 1
    Call SendDataTo(index, "FINQUETE" & SEP_CHAR & END_CHAR)
Else
    Call QueteMsg(index, "Je suis désolé tu n'as pas l'objet que je cherche.")
End If

End Sub

Sub PlayerQueteTypeTuer(ByVal index As Long, ByVal Queteec As Long, ByVal NpcTnum As Long)
Dim i As Long
Dim n As Long

If Not IsPlaying(index) Then Exit Sub
If Queteec <= 0 Then Exit Sub
If NpcTnum <= 0 Or NpcTnum > MAX_NPCS Then Exit Sub
If GetPlayerQueteEtat(index, Queteec) = True Then Exit Sub

For i = 1 To 15
    If NpcTnum = quete(Queteec).indexe(i).data1 And quete(Queteec).indexe(i).data2 > 0 And Player(index).Char(Player(index).CharNum).Quetep.indexe(i).data2 < quete(Queteec).indexe(i).data2 Then
        Player(index).Char(Player(index).CharNum).Quetep.indexe(i).data1 = 1
        Player(index).Char(Player(index).CharNum).Quetep.indexe(i).data2 = Val(Player(index).Char(Player(index).CharNum).Quetep.indexe(i).data2) + 1
        Call SendDataTo(index, "TUERQUETE" & SEP_CHAR & END_CHAR)
    End If
    
    If quete(Queteec).indexe(i).data2 <= 0 Then
        Player(index).Char(Player(index).CharNum).Quetep.indexe(i).data1 = 1
        Player(index).Char(Player(index).CharNum).Quetep.indexe(i).data2 = 0
    End If
Next i

n = 0
For i = 1 To 15
    If Player(index).Char(Player(index).CharNum).Quetep.indexe(i).data1 = 1 And Player(index).Char(Player(index).CharNum).Quetep.indexe(i).data2 >= quete(Queteec).indexe(i).data2 Then
        n = n + 1
    End If
Next i

If n = 15 And Player(index).Char(Player(index).CharNum).QueteStatut(Queteec) = 0 Then
    Player(index).Char(Player(index).CharNum).QueteStatut(Queteec) = 1
    Call SendDataTo(index, "FINQUETE" & SEP_CHAR & END_CHAR)
End If

End Sub

Sub PlayerQueteTypeXp(ByVal index As Long, ByVal Queteec As Long, ByVal Xp As Long)
Dim i As Long
Dim n As Long

If Not IsPlaying(index) Then Exit Sub
If Queteec <= 0 Or Xp <= 0 Then Exit Sub
If GetPlayerQueteEtat(index, Queteec) = True Then Exit Sub

If Xp > GetPlayerExp(index) Then Xp = Xp - GetPlayerExp(index)

Player(index).Char(Player(index).CharNum).Quetep.data1 = Val(Player(index).Char(Player(index).CharNum).Quetep.data1) + Val(Xp)
Call SendDataTo(index, "XPQUETE" & SEP_CHAR & Player(index).Char(Player(index).CharNum).Quetep.data1 & SEP_CHAR & END_CHAR)

If Val(Player(index).Char(Player(index).CharNum).Quetep.data1) >= Val(quete(Queteec).data1) And Player(index).Char(Player(index).CharNum).QueteStatut(Queteec) = 0 Then
    Player(index).Char(Player(index).CharNum).QueteStatut(Queteec) = 1
    Call SendDataTo(index, "FINQUETE" & SEP_CHAR & END_CHAR)
End If

End Sub

Sub TerminerPlayerQuete(ByVal index As Long, ByVal QueteTindex As Long)
Dim Packet As String
Dim i As Long

If Not IsPlaying(index) Then Exit Sub
If QueteTindex <= 0 Then Exit Sub
If GetPlayerQueteEtat(index, QueteTindex) Then Exit Sub

Call ClearPlayerQuete(index)
Call SendPlayerQuete(index)

If GetPlayerLevel(index) = MAX_LEVEL Then
    If quete(QueteTindex).Recompence.Exp > 0 Then
        Call SetPlayerExp(index, experience(MAX_LEVEL))
        Call BattleMsg(index, "Tu ne peux pas gagner plus d'expérience!", BrightBlue, 0)
    End If
Else
    If quete(QueteTindex).Recompence.Exp > 0 Then
        Call SetPlayerExp(index, Player(index).Char(Player(index).CharNum).Exp + quete(QueteTindex).Recompence.Exp)
        Call BattleMsg(index, "Tu as gagné " & quete(QueteTindex).Recompence.Exp & "pts d'expérience.", BrightBlue, 0)
    End If
End If
Call CheckPlayerLevelUp(index)
Call SendPlayerData(index)

If quete(QueteTindex).Recompence.objq1 > 0 And quete(QueteTindex).Recompence.objn1 > 0 Then Call GiveItem(index, quete(QueteTindex).Recompence.objn1, Val(quete(QueteTindex).Recompence.objq1))

If quete(QueteTindex).Recompence.objq2 > 0 And quete(QueteTindex).Recompence.objn2 > 0 Then Call GiveItem(index, quete(QueteTindex).Recompence.objn2, Val(quete(QueteTindex).Recompence.objq2))

If quete(QueteTindex).Recompence.objq3 > 0 And quete(QueteTindex).Recompence.objn3 > 0 Then Call GiveItem(index, quete(QueteTindex).Recompence.objn3, Val(quete(QueteTindex).Recompence.objq3))

Call SendPlayerData(index)
Call SendInventory(index)

If quete(QueteTindex).Case > 0 Then MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & index & "," & Val(quete(QueteTindex).Case)

Player(index).Char(Player(index).CharNum).QueteStatut(QueteTindex) = 2

Packet = "TERMINEQUETE" & SEP_CHAR & END_CHAR
Call SendDataTo(index, Packet)

End Sub

Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
Dim Packet As String
Dim npcnum As Long
Dim i As Long, X As Long, Y As Long
Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
    
    Spawned = False
    
    npcnum = Map(MapNum).Npc(MapNpcNum)

    If npcnum > 0 Then
        If GameTime = TIME_NIGHT Then
            If Npc(npcnum).SpawnTime = 1 Then
                MapNpc(MapNum, MapNpcNum).num = 0
                MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
                MapNpc(MapNum, MapNpcNum).HP = 0
                Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        Else
            If Npc(npcnum).SpawnTime = 2 Then
                MapNpc(MapNum, MapNpcNum).num = 0
                MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
                MapNpc(MapNum, MapNpcNum).HP = 0
                Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        End If
    
        MapNpc(MapNum, MapNpcNum).num = npcnum
        MapNpc(MapNum, MapNpcNum).Target = 0
        
        MapNpc(MapNum, MapNpcNum).HP = GetNpcMaxHP(npcnum)
        MapNpc(MapNum, MapNpcNum).MP = GetNpcMaxMP(npcnum)
        MapNpc(MapNum, MapNpcNum).SP = GetNpcMaxSP(npcnum)
                
        MapNpc(MapNum, MapNpcNum).Dir = Int(Rnd * 4)
        
        ' Well try 100 times to randomly place the sprite
        If Map(MapNum).Npcs(MapNpcNum).Hasardp = 0 Then
            MapNpc(MapNum, MapNpcNum).X = Map(MapNum).Npcs(MapNpcNum).X
            MapNpc(MapNum, MapNpcNum).Y = Map(MapNum).Npcs(MapNpcNum).Y
            If Map(MapNum).Npcs(MapNpcNum).Imobile > 0 Then MapNpc(MapNum, MapNpcNum).Dir = Map(MapNum).Npcs(MapNpcNum).Imobile - 1
            Spawned = True
        Else
            For i = 1 To 100
                X = Int(Rnd * MAX_MAPX)
                Y = Int(Rnd * MAX_MAPY)
                
                ' Check if the tile is walkable
                If Map(MapNum).Tile(X, Y).type = TILE_TYPE_WALKABLE Then
                    MapNpc(MapNum, MapNpcNum).X = X
                    MapNpc(MapNum, MapNpcNum).Y = Y
                    Map(MapNum).Npcs(MapNpcNum).X = X
                    Map(MapNum).Npcs(MapNpcNum).Y = Y
                    Spawned = True
                    Exit For
                End If
            Next i
        End If
            ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    If Map(MapNum).Tile(X, Y).type = TILE_TYPE_WALKABLE Then
                        MapNpc(MapNum, MapNpcNum).X = X
                        MapNpc(MapNum, MapNpcNum).Y = Y
                        Spawned = True
                    End If
                Next X
            Next Y
        End If
             
        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
        End If
    Else
        MapNpc(MapNum, MapNpcNum).num = 0
        
        Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).num & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, Packet)
    End If
    
    'Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(MapNum, MapNpcNum).num) & SEP_CHAR & END_CHAR)
End Sub

Sub SpawnMapNpcs(ByVal MapNum As Long)
Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, MapNum)
    Next i
End Sub

Sub SpawnAllMapNpcs()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next i
End Sub

Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
Dim AttackSpeed As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Then Exit Function
    
    On Error GoTo er:
    
    If GetPlayerWeaponSlot(Attacker) > 0 Then
        AttackSpeed = item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
    Else
        AttackSpeed = 1000
    End If
    
    CanAttackPlayer = False
    
    ' Make sure they have more then 0 hp
    If GetPlayerHP(Victim) <= 0 Then Exit Function
    
    ' Make sure we dont attack the player if they are switching maps
    If Player(Victim).GettingMap = YES Then Exit Function
        
    ' Make sure they are on the same map
    If (GetPlayerMap(Attacker) = GetPlayerMap(Victim)) And (GetTickCount > Player(Attacker).AttackTimer + AttackSpeed) Then
        
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
                If (GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).type <> TILE_TYPE_ARENA Then
                        ' Check if map is attackable
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                            
                            If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                                If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                                    CanAttackPlayer = True
                                Else
                                    Call PlayerMsg(Attacker, "Vous ne pouvez pas attaquer un membre de votre guilde !", BrightRed)
                                End If
                            Else
                                CanAttackPlayer = True
                            End If
                        Else
                            Call PlayerMsg(Attacker, "C'est une safe zone (impossible d'attaquer d'autres joueurs)!", BrightRed)
                        End If
                    ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If

            Case DIR_DOWN
                If (GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).type <> TILE_TYPE_ARENA Then
                        ' Check if map is attackable
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                                                                
                            If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                                If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                                    CanAttackPlayer = True
                                Else
                                    Call PlayerMsg(Attacker, "Vous ne pouvez pas attaquer un membre de votre guilde !", BrightRed)
                                End If
                            Else
                                CanAttackPlayer = True
                            End If
                        Else
                            Call PlayerMsg(Attacker, "C'est une safe zone (impossible d'attaquer d'autres joueurs)!", BrightRed)
                        End If
                        
                    ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If
        
            Case DIR_LEFT
                If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).type <> TILE_TYPE_ARENA Then
                        ' Check if map is attackable
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                            If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                                If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                                    CanAttackPlayer = True
                                Else
                                    Call PlayerMsg(Attacker, "Vous ne pouvez pas attaquer un membre de votre guilde !", BrightRed)
                                End If
                            Else
                                CanAttackPlayer = True
                            End If
                        Else
                            Call PlayerMsg(Attacker, "C'est une safe zone (impossible d'attaquer d'autres joueurs)!", BrightRed)
                        End If
                        
                    ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If
                
            
            Case DIR_RIGHT
                If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).type <> TILE_TYPE_ARENA Then
                        ' Check if map is attackable
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                            If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                                If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                                    CanAttackPlayer = True
                                Else
                                    Call PlayerMsg(Attacker, "Vous ne pouvez pas attaquer un membre de votre guilde !", BrightRed)
                                End If
                            Else
                                CanAttackPlayer = True
                            End If
                        Else
                            Call PlayerMsg(Attacker, "C'est une safe zone (impossible d'attaquer d'autres joueurs)!", BrightRed)
                        End If
                    ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If
                
        End Select
    End If
    
Exit Function
er:
CanAttackPlayer = False
On Error Resume Next
If Attacker < 0 Or Attacker > MAX_PLAYERS Or Victim < 0 Or Victim > MAX_PLAYERS Then Exit Function
Call PlayerMsg(Attacker, "Attaque annulée à cause d'une erreur si le problème persiste contactez un administrateur.", Red)
Call PlayerMsg(Victim, "Attaque (du joueur qui vous attaque) annulée à cause d'une erreur si le problème persiste contactez un administrateur.", Red)
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur dans l'attaque d'un joueur(ATT : " & Player(Attacker).Login & ",VIC : " & Player(Victim).Login & "). Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur dans l'attaque d'un joueur(ATT : " & GetPlayerName(Attacker) & ",VIC : " & GetPlayerName(Victim) & ")", BrightRed, True)
End Function

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
Dim MapNum As Long, npcnum As Long
Dim AttackSpeed As Long
Dim TmpX As Byte, TmpY As Byte



' Check for subscript out of range
If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then Exit Function

On Error GoTo er:

If CLng(Npc(MapNpc(GetPlayerMap(Attacker), MapNpcNum).num).Vol) <> 0 Then CanAttackNpc = False: Exit Function

If GetPlayerWeaponSlot(Attacker) > 0 Then
    AttackSpeed = item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
Else
    AttackSpeed = 1000
End If

CanAttackNpc = False
 
' Check for subscript out of range
If MapNpc(GetPlayerMap(Attacker), MapNpcNum).num <= 0 Or MapNpc(GetPlayerMap(Attacker), MapNpcNum).num > MAX_NPCS Then Exit Function
 
MapNum = GetPlayerMap(Attacker)
npcnum = MapNpc(MapNum, MapNpcNum).num
 
If GetPlayerWeaponSlot(Attacker) > 0 Then
    If Npc(npcnum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or Npc(npcnum).Behavior = NPC_BEHAVIOR_ATTACKWHENATTACKED Then
        If Npc(npcnum).QueteNum > 11 Then
            If Npc(npcnum).QueteNum - 11 <> Player(Attacker).Char(Player(Attacker).CharNum).metier Then
                CanAttackNpc = False
                Exit Function
            End If
        ElseIf Npc(npcnum).QueteNum > 0 Then
            If Npc(npcnum).QueteNum <> item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).tArme Then
                CanAttackNpc = False
                Exit Function
            End If
        End If
    End If
Else
    If Npc(npcnum).QueteNum > 11 Then
            If Npc(npcnum).QueteNum - 11 <> Player(Attacker).Char(Player(Attacker).CharNum).metier Then
                CanAttackNpc = False
                Exit Function
            End If
        ElseIf Npc(npcnum).QueteNum > 0 Then
            CanAttackNpc = False
            Exit Function
        End If
End If
 
' Make sure the npc isn't already dead
If MapNpc(MapNum, MapNpcNum).HP <= 0 And CLng(Npc(npcnum).Inv) = 0 Then Exit Function
 
' Make sure they are on the same map
If IsPlaying(Attacker) Then
    If npcnum > 0 And GetTickCount > Player(Attacker).AttackTimer + AttackSpeed Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
                TmpX = 1
                TmpY = 2
 
            Case DIR_DOWN
                TmpX = 1
                TmpY = 0
                 
            Case DIR_LEFT
                TmpX = 2
                TmpY = 1
                 
            Case DIR_RIGHT
                TmpX = 0
                TmpY = 1
        End Select
        
        If (MapNpc(MapNum, MapNpcNum).Y + (TmpY - 1) = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).X + (TmpX - 1) = GetPlayerX(Attacker)) Then
            If Npc(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(npcnum).Behavior <> NPC_BEHAVIOR_QUETEUR And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SCRIPT Then
                CanAttackNpc = True
                If Val(Scripting) = 1 And IsNumeric(Trim$(Npc(npcnum).AttackSay)) Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & Attacker & "," & Val(Trim$(Npc(npcnum).AttackSay))
                    Exit Function
                End If
            Else
                If Val(Scripting) = 1 And IsNumeric(Trim$(Npc(npcnum).AttackSay)) Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & Attacker & "," & Val(Trim$(Npc(npcnum).AttackSay))
                    Exit Function
                Else
                    If Trim$(Npc(npcnum).AttackSay) > vbNullString And Npc(npcnum).Behavior <> NPC_BEHAVIOR_QUETEUR Then Call QueteMsg(Attacker, Trim$(Npc(npcnum).Name) & " : " & Trim$(Npc(npcnum).AttackSay))
                End If
                
                If Npc(npcnum).Behavior = NPC_BEHAVIOR_QUETEUR Then
                    If Player(Attacker).Char(Player(Attacker).CharNum).QueteStatut(Npc(npcnum).QueteNum) <> 2 Then
                        If Player(Attacker).Char(Player(Attacker).CharNum).QueteStatut(Npc(npcnum).QueteNum) = 1 Then
                            If Val(Scripting) = 1 And IsNumeric(Trim$(quete(Npc(npcnum).QueteNum).reponse)) Then
                                MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & Attacker & "," & Val(Trim$(quete(Npc(npcnum).QueteNum).reponse))
                                Exit Function
                            Else
                                Call QueteMsg(Attacker, Trim$(quete(Npc(npcnum).QueteNum).nom) & " : " & Trim$(quete(Npc(npcnum).QueteNum).reponse))
                            End If
                            Call TerminerPlayerQuete(Attacker, Npc(npcnum).QueteNum)
                        Else
                            Call SendDataTo(Attacker, "QUETECOUR" & SEP_CHAR & Npc(npcnum).QueteNum & SEP_CHAR & END_CHAR)
                            Player(Attacker).Char(Player(Attacker).CharNum).QueteEnCour = Npc(npcnum).QueteNum
                            Call QueteMsg(Attacker, Trim$(quete(Npc(npcnum).QueteNum).nom) & " : " & Trim$(quete(Npc(npcnum).QueteNum).description))
                        End If
                    End If
                ElseIf Npc(npcnum).Behavior = NPC_BEHAVIOR_SHOPKEEPER Then
                    Call QueteMsg(Attacker, Shop(Npc(npcnum).QueteNum).JoinSay)
                    Call SendTrade(Attacker, Npc(npcnum).QueteNum)
                ElseIf Npc(npcnum).Behavior = NPC_BEHAVIOR_SCRIPT Then
                    If Val(Scripting) = 1 Then
                        MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & Attacker & "," & (Npc(npcnum).QueteNum)
                        Exit Function
                    End If
                Else
                    If Player(Attacker).Char(Player(Attacker).CharNum).QueteEnCour > 0 Then
                        If quete(Player(Attacker).Char(Player(Attacker).CharNum).QueteEnCour).type = QUETE_TYPE_APORT And quete(Player(Attacker).Char(Player(Attacker).CharNum).QueteEnCour).data2 = npcnum Then
                            If GetPlayerQueteEtat(Attacker, Player(Attacker).Char(Player(Attacker).CharNum).QueteEnCour) Then Exit Function
                            Call PlayerQueteTypeAport(Attacker, Player(Attacker).Char(Player(Attacker).CharNum).QueteEnCour)
                        End If
                        If quete(Player(Attacker).Char(Player(Attacker).CharNum).QueteEnCour).type = QUETE_TYPE_PARLER And quete(Player(Attacker).Char(Player(Attacker).CharNum).QueteEnCour).data1 = npcnum Then
                            If GetPlayerQueteEtat(Attacker, Player(Attacker).Char(Player(Attacker).CharNum).QueteEnCour) Then Exit Function
                            If Player(Attacker).Char(Player(Attacker).CharNum).QueteStatut(Npc(npcnum).QueteNum) > 0 Then Exit Function
                            Player(Attacker).Char(Player(Attacker).CharNum).QueteStatut(Npc(npcnum).QueteNum) = 1
                            Call QueteMsg(Attacker, Trim$(Npc(npcnum).Name) & " : " & Trim$(Npc(npcnum).AttackSay))
                            Call SendDataTo(Attacker, "FINQUETE" & SEP_CHAR & END_CHAR)
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

Exit Function
er:
CanAttackNpc = False
On Error Resume Next
If Attacker < 0 Or Attacker > MAX_PLAYERS Then Exit Function
Call PlayerMsg(Attacker, "Attaque annulée à cause d'une erreur si le problème persiste contactez un administrateur.", Red)
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur dans l'attaque d'un PNJ(" & npcnum & ") par un joueur(" & Player(Attacker).Login & "). Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur dans l'attaque d'un PNJ(" & npcnum & ") par un joueur(" & GetPlayerName(Attacker) & ")", BrightRed, True)
End Function

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long) As Boolean
Dim MapNum As Long, npcnum As Long
    
    CanNpcAttackPlayer = False
    
    If Not IsPlaying(index) Then Exit Function
    
    On Error GoTo er:
    
    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(index) = False Then Exit Function
        
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(index), MapNpcNum).num <= 0 Or MapNpc(GetPlayerMap(index), MapNpcNum).num > MAX_NPCS Then Exit Function
        
    MapNum = GetPlayerMap(index)
    npcnum = MapNpc(MapNum, MapNpcNum).num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum, MapNpcNum).HP <= 0 And CLng(Npc(npcnum).Inv) = 0 Then Exit Function
        
    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(MapNum, MapNpcNum).AttackTimer + 1000 Then Exit Function
        
    ' Make sure we dont attack the player if they are switching maps
    If Player(index).GettingMap = YES Then Exit Function
    
    MapNpc(MapNum, MapNpcNum).AttackTimer = GetTickCount
    
    ' Make sure they are on the same map
    If IsPlaying(index) Then
        If npcnum > 0 Then
            ' Check if at same coordinates
            If (GetPlayerY(index) + 1 = MapNpc(MapNum, MapNpcNum).Y) And (GetPlayerX(index) = MapNpc(MapNum, MapNpcNum).X) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(index) = MapNpc(MapNum, MapNpcNum).Y + 1) And (GetPlayerX(index) = MapNpc(MapNum, MapNpcNum).X) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(index) = MapNpc(MapNum, MapNpcNum).Y) And (GetPlayerX(index) + 1 = MapNpc(MapNum, MapNpcNum).X) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(index) = MapNpc(MapNum, MapNpcNum).Y) And (GetPlayerX(index) = MapNpc(MapNum, MapNpcNum).X + 1) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If
        End If
    End If

Exit Function
er:
CanNpcAttackPlayer = False
On Error Resume Next
If index < 0 Or index > MAX_PLAYERS Then Exit Function
Call PlayerMsg(index, "Attaque du PNJ annulée à cause d'une erreur si le problème persiste contactez un administrateur.", Red)
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur dans l'attaque d'un joueur(" & Player(index).Login & ")par un PNJ(" & npcnum & "). Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur dans l'attaque d'un joueur(" & GetPlayerName(index) & ")par un PNJ(" & npcnum & ")", BrightRed, True)
End Function

Sub AttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long)
Dim Exp As Long
Dim n As Long
Dim i As Long
    
    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then Exit Sub
    
    On Error GoTo er:
    
    ' Check for weapon
    If GetPlayerWeaponSlot(Attacker) > 0 Then n = GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker)) Else n = 0
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), "ATTACK" & SEP_CHAR & Attacker & SEP_CHAR & END_CHAR)
    
    ' Effet de sang
    Call SendDataToMap(GetPlayerMap(Attacker), "BloodAnim" & SEP_CHAR & Attacker & SEP_CHAR & Victim & SEP_CHAR & TARGET_TYPE_PLAYER & SEP_CHAR & END_CHAR)

If Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).type <> TILE_TYPE_ARENA Then
    If Damage >= GetPlayerHP(Victim) Then
        ' Set HP to nothing
        Call SetPlayerHP(Victim, 0)
        
        ' Check for a weapon and say damage
        Call BattleMsg(Attacker, "Tu frappes " & GetPlayerName(Victim) & " pour " & Damage & " dommages.", White, 1)
        Call BattleMsg(Victim, GetPlayerName(Attacker) & " te frappe pour " & Damage & " dommages.", BrightRed, 1)
    
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " a été tué par " & GetPlayerName(Attacker), BrightRed)
        
        If Map(GetPlayerMap(Victim)).Moral <> MAP_MORAL_NO_PENALTY Then
            If Scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "DropItems " & Victim
            Else
                If GetPlayerWeaponSlot(Victim) > 0 Then Call PlayerMapDropItem(Victim, GetPlayerWeaponSlot(Victim), 0)
                            
                If GetPlayerArmorSlot(Victim) > 0 Then Call PlayerMapDropItem(Victim, GetPlayerArmorSlot(Victim), 0)
                
                If GetPlayerHelmetSlot(Victim) > 0 Then Call PlayerMapDropItem(Victim, GetPlayerHelmetSlot(Victim), 0)
                            
                If GetPlayerShieldSlot(Victim) > 0 Then Call PlayerMapDropItem(Victim, GetPlayerShieldSlot(Victim), 0)
            End If
            
            ' Calculate exp to give attacker
            Exp = (GetPlayerExp(Victim) \ 10) * Val(GetVar(App.Path & "/Data.ini", "RATIO", "Exp_pvp"))
            If Val(GetVar(App.Path & "/Data.ini", "CONFIG", "ExpDynamique")) = 1 Then
                Exp = Int(100 * Rnd() + Exp - (100 * Rnd()))
            End If
            
            ' Make sure we dont get less then 0
            If Exp < 0 Then Exp = 0
                        
            If GetPlayerLevel(Victim) = MAX_LEVEL Then
                Call BattleMsg(Victim, "Tu ne peux pas perdre d'expérience!", BrightRed, 1)
                Call BattleMsg(Attacker, GetPlayerName(Victim) & " est niveau maximum!", BrightBlue, 0)
            Else
                If Exp = 0 Then
                    Call BattleMsg(Victim, "Tu n'as pas perdu d'expérience.", BrightRed, 1)
                    Call BattleMsg(Attacker, "Tu ne reçois pas d'expérience.", BrightBlue, 0)
                Else
                    Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                    Call BattleMsg(Victim, "Tu perds " & Exp & "pts d'expérience.", BrightRed, 1)
                    Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                    Call BattleMsg(Attacker, "Tu gagnes " & Exp & "pts d'expérience pour avoir tué " & GetPlayerName(Victim) & ".", BrightBlue, 0)
                End If
            End If
        End If
        
        ' Warp player away
        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & Victim
        Else
            Call PlayerWarp(Victim, START_MAP, START_X, START_Y)
        End If
        
        ' Restore vitals
        Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
        Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
        Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
        Call SendHP(Victim)
        Call SendMP(Victim)
        Call SendSP(Victim)
                
        ' Check for a level up
        Call CheckPlayerLevelUp(Attacker)
        
        ' Check if target is player who died and if so set target to 0
        If Player(Attacker).TargetType = TARGET_TYPE_PLAYER And Player(Attacker).Target = Victim Then
            Player(Attacker).Target = -1
            Player(Attacker).TargetType = 0
        End If
        
        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " est maintenant un criminel!", BrightRed)
            End If
        Else
            Call SetPlayerPK(Victim, NO)
            Call SendPlayerData(Victim)
            Call GlobalMsg(GetPlayerName(Victim) & " a payé le prix d'être un criminel!", BrightRed)
        End If
    Else
        ' Player not dead, just do the damage
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        Call SendHP(Victim)
        
        ' Check for a weapon and say damage
        Call BattleMsg(Attacker, "Tu frappes " & GetPlayerName(Victim) & " pour " & Damage & " dommages.", White, 1)
        Call BattleMsg(Victim, GetPlayerName(Attacker) & " te frappe pour " & Damage & " dommages.", BrightRed, 1)
    End If
ElseIf Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).type = TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).type = TILE_TYPE_ARENA Then
    If Damage >= GetPlayerHP(Victim) Then
        ' Set HP to nothing
        Call SetPlayerHP(Victim, 0)
        
        ' Check for a weapon and say damage
        Call BattleMsg(Attacker, "Vous frappez " & GetPlayerName(Victim) & " pour " & Damage & " dommages.", White, 0)
        Call BattleMsg(Victim, GetPlayerName(Attacker) & " te frappe pour " & Damage & " dommages.", BrightRed, 1)
            
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " a été tué dans l'arène par " & GetPlayerName(Attacker), BrightRed)
            
        ' Warp player away
        Call PlayerWarp(Victim, Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).data1, Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).data2, Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).data3)
        
        ' Restore vitals
        Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
        Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
        Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
        Call SendHP(Victim)
        Call SendMP(Victim)
        Call SendSP(Victim)
                        
        ' Check if target is player who died and if so set target to 0
        If Player(Attacker).TargetType = TARGET_TYPE_PLAYER And Player(Attacker).Target = Victim Then
            Player(Attacker).Target = -1
            Player(Attacker).TargetType = 0
        End If
    Else
        ' Player not dead, just do the damage
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        Call SendHP(Victim)
        
        ' Check for a weapon and say damage
        Call BattleMsg(Attacker, "Tu frappes " & GetPlayerName(Victim) & " pour " & Damage & " dommages.", White, 1)
        Call BattleMsg(Victim, GetPlayerName(Attacker) & " te frappe pour " & Damage & " dommages.", BrightRed, 1)
    End If
End If
    
    ' Reset timer for attacking
    Player(Attacker).AttackTimer = GetTickCount
    Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "pain" & SEP_CHAR & END_CHAR)

Exit Sub
er:
On Error Resume Next
If Attacker < 0 Or Attacker > MAX_PLAYERS Or Victim < 0 Or Victim > MAX_PLAYERS Then Exit Sub
Call PlayerMsg(Attacker, "Attaque du joueur annulée à cause d'une erreur si le problème persiste contactez un administrateur.", Red)
Call PlayerMsg(Victim, "Attaque (du joueur qui vous attaque) annulée à cause d'une erreur si le problème persiste contactez un administrateur.", Red)
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur dans l'attaque d'un joueur(" & Player(Victim).Login & ")par un autre joueur(" & Player(Attacker).Login & "). Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur dans l'attaque d'un joueur(" & GetPlayerName(Victim) & ")par un autre joueur(" & GetPlayerName(Attacker) & ")", BrightRed, True)
End Sub

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long)
Dim Name As String
Dim Exp As Long
Dim MapNum As Long

    If Not IsPlaying(Victim) Then Exit Sub
    
    On Error GoTo er:
    
    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim), MapNpcNum).num <= 0 Then Exit Sub
            
    ' Send this packet so they can see the person attacking
    Call SendDataToMap(GetPlayerMap(Victim), "NPCATTACK" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
    
    MapNum = GetPlayerMap(Victim)
    
    ':: AUTO TURN ::
    'If Val(GetVar(App.Path & "\Data.ini", "CONFIG", "AutoTurn")) = 1 Then
        'If GetPlayerX(Victim) - 1 = MapNpc(MapNum, MapNpcNum).X Then
            'Call SetPlayerDir(Victim, DIR_LEFT)
        'End If
        'If GetPlayerX(Victim) + 1 = MapNpc(MapNum, MapNpcNum).X Then
            'Call SetPlayerDir(Victim, DIR_RIGHT)
        'End If
        'If GetPlayerY(Victim) - 1 = MapNpc(MapNum, MapNpcNum).Y Then
            'Call SetPlayerDir(Victim, DIR_UP)
        'End If
        'If GetPlayerY(Victim) + 1 = MapNpc(MapNum, MapNpcNum).Y Then
            'Call SetPlayerDir(Victim, DIR_DOWN)
        'End If
        'Call SendDataToMap(GetPlayerMap(Victim), "changedir" & SEP_CHAR & GetPlayerDir(Victim) & SEP_CHAR & Victim & SEP_CHAR & END_CHAR)
    'End If
    ':: END AUTO TURN ::
    
    Name = Trim$(Npc(MapNpc(MapNum, MapNpcNum).num).Name)
    
    If Damage >= GetPlayerHP(Victim) Then
        ' Say damage
        Call BattleMsg(Victim, "Tu frappes pour " & Damage & " dommages.", BrightRed, 1)
                
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " a été tué par " & Name, BrightRed)
        
        If Map(GetPlayerMap(Victim)).Moral <> MAP_MORAL_NO_PENALTY Then
            If Scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "DropItems " & Victim
            Else
                If GetPlayerWeaponSlot(Victim) > 0 Then Call PlayerMapDropItem(Victim, GetPlayerWeaponSlot(Victim), 0)
                            
                If GetPlayerArmorSlot(Victim) > 0 Then Call PlayerMapDropItem(Victim, GetPlayerArmorSlot(Victim), 0)
                                
                If GetPlayerHelmetSlot(Victim) > 0 Then Call PlayerMapDropItem(Victim, GetPlayerHelmetSlot(Victim), 0)
                            
                If GetPlayerShieldSlot(Victim) > 0 Then Call PlayerMapDropItem(Victim, GetPlayerShieldSlot(Victim), 0)
            End If
            
            ' Calculate exp to give attacker
            Exp = (GetPlayerExp(Victim) \ 3)
            
            ' Make sure we dont get less then 0
            If Exp < 0 Then Exp = 0
                        
            If Exp = 0 Then
                Call BattleMsg(Victim, "Tu ne perds pas d'expérience.", BrightRed, 0)
            Else
                Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                Call BattleMsg(Victim, "Tu perds " & Exp & "pts d'expérience.", BrightRed, 0)
            End If
        End If
                
        ' Warp player away
        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & Victim
        Else
            Call PlayerWarp(Victim, START_MAP, START_X, START_Y)
        End If
        
        ' Restore vitals
        Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
        Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
        Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
        Call SendHP(Victim)
        Call SendMP(Victim)
        Call SendSP(Victim)
        
        ' Set NPC target to 0
        MapNpc(MapNum, MapNpcNum).Target = 0
        
        ' If the player the attacker killed was a pk then take it away
        If GetPlayerPK(Victim) = YES Then
            Call SetPlayerPK(Victim, NO)
            Call SendPlayerData(Victim)
        End If
    Else
        ' Player not dead, just do the damage
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        Call SendHP(Victim)
        
        ' Say damage
        Call BattleMsg(Victim, "Vous avez subi " & Damage & " dommage.", BrightRed, 1)
    End If
    
    Call SendDataTo(Victim, "BLITNPCDMG" & SEP_CHAR & Damage & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
    Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "pain" & SEP_CHAR & END_CHAR)

Exit Sub
er:
On Error Resume Next
If Victim < 0 Or Victim > MAX_PLAYERS Then Exit Sub
Call PlayerMsg(Victim, "Attaque (du PNJ qui vous attaque) annulée à cause d'une erreur si le problème persiste contactez un administrateur.", Red)
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur dans l'attaque d'un joueur(" & Player(Victim).Login & ")par un PNJ(" & MapNpc(MapNum, MapNpcNum).num & "). Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur dans l'attaque d'un joueur(" & GetPlayerName(Victim) & ")par un PNJ(" & MapNpc(MapNum, MapNpcNum).num & ")", BrightRed, True)
End Sub

Sub AttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long)
Dim Name As String
Dim Exp As Long, ExpG As Long
Dim n As Long, i As Long, q As Integer, X As Long
Dim STR As Long, def As Long, MapNum As Long, npcnum As Long

    On Error GoTo er:
    
    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then Exit Sub
     
    ' Check for weapon
    If GetPlayerWeaponSlot(Attacker) > 0 Then
        n = GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))
    Else
        n = 0
    End If
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), "ATTACK" & SEP_CHAR & Attacker & SEP_CHAR & END_CHAR)
    
    ' Effet de sang
    Call SendDataToMap(GetPlayerMap(Attacker), "BloodAnim" & SEP_CHAR & Attacker & SEP_CHAR & MapNpcNum & SEP_CHAR & TARGET_TYPE_NPC & SEP_CHAR & END_CHAR)
    
    MapNum = GetPlayerMap(Attacker)
    npcnum = MapNpc(MapNum, MapNpcNum).num
    Name = Trim$(Npc(npcnum).Name)
        
    If npcnum <= 0 Or npcnum > MAX_NPCS Then Exit Sub
    
    If CLng(Npc(npcnum).Inv) = -1 Or CLng(Npc(npcnum).Inv) = 1 Then
        Damage = 0
        If Trim$(Npc(npcnum).AttackSay) <> vbNullString Then Call QueteMsg(Attacker, Trim$(Npc(npcnum).Name) & " : " & Trim$(Npc(npcnum).AttackSay) & "")
        Exit Sub
    End If
    
    If Damage >= MapNpc(MapNum, MapNpcNum).HP And MapNpc(MapNum, MapNpcNum).HP > 0 Then
        ' Check for a weapon and say damage
        
        Call BattleMsg(Attacker, "Tu tues " & Name, BrightRed, 1)

        Dim add As String

        add = 0
        If GetPlayerWeaponSlot(Attacker) > 0 Then add = add + item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AddEXP
        If GetPlayerArmorSlot(Attacker) > 0 Then add = add + item(GetPlayerInvItemNum(Attacker, GetPlayerArmorSlot(Attacker))).AddEXP
        If GetPlayerShieldSlot(Attacker) > 0 Then add = add + item(GetPlayerInvItemNum(Attacker, GetPlayerShieldSlot(Attacker))).AddEXP
        If GetPlayerHelmetSlot(Attacker) > 0 Then add = add + item(GetPlayerInvItemNum(Attacker, GetPlayerHelmetSlot(Attacker))).AddEXP
        
        If add > 0 Then
            If add < 100 Then
                If add < 10 Then
                    add = 0 & ".0" & Right$(add, 2)
                Else
                    add = 0 & "." & Right$(add, 2)
                End If
            Else
                add = Mid$(add, 1, 1) & "." & Right$(add, 2)
            End If
        End If
                                
        ' Metier chasseur
        If Player(Attacker).Char(Player(Attacker).CharNum).metier > 0 Then
            n = Player(Attacker).Char(Player(Attacker).CharNum).metier
            If metier(n).type = METIER_CHASSEUR Then
                If InMetier(n, npcnum) <> 10 Then
                    If Player(Attacker).Char(Player(Attacker).CharNum).MetierLvl < 200 Then
                        Player(Attacker).Char(Player(Attacker).CharNum).MetierExp = Player(Attacker).Char(Player(Attacker).CharNum).MetierExp + metier(n).data(InMetier(n, npcnum), 1)
                        Call BattleMsg(Attacker, "(Metier) Vous avez gagné " & metier(n).data(InMetier(n, npcnum), 1) & " pts d'expérience.", BrightBlue, 0)
                    Else
                        Call BattleMsg(Attacker, "(Metier) Vous ne pouver plus gagnez d'expérience", BrightBlue, 0)
                    End If
                End If
            End If
        End If
        Call checkLvlUpMetier(Attacker)
                                
        ' Calculate exp to give attacker
        If Val(add) > 0 Then
            Exp = (Npc(npcnum).Exp + (Npc(npcnum).Exp * Val(add))) * Val(GetVar(App.Path & "/Data.ini", "RATIO", "Exp_pvm"))
        Else
            Exp = Npc(npcnum).Exp * Val(GetVar(App.Path & "/Data.ini", "RATIO", "Exp_pvm"))
        End If
        If Val(GetVar(App.Path & "/Data.ini", "CONFIG", "ExpDynamique")) = 1 Then
            Exp = Int(100 * Rnd() + Exp - (100 * Rnd()))
        End If

        ' Make sure we dont get less then 0
        If Exp < 0 Then Exp = 1

        ' Check if in party, if so divide the exp up by 2
        If Player(Attacker).InParty = 0 Or Party.ShareExp(Player(Attacker).InParty) = 0 Then
            If GetPlayerLevel(Attacker) = MAX_LEVEL Then
                Call SetPlayerExp(Attacker, experience(MAX_LEVEL))
                Call BattleMsg(Attacker, "Tu ne peux pas gagner plus d'expérience!", BrightBlue, 0)
            Else
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                Call BattleMsg(Attacker, "Tu as gagné " & Exp & " pts d'expérience.", BrightBlue, 0)
            End If
        Else
            q = Party.MemberCount(Player(Attacker).InParty)
            If Party.ShareExp(Player(Attacker).InParty) = 2 Then
                For X = 1 To q
                    n = Party.PlayerIndex(Player(Attacker).InParty, X)
                    i = i + Player(n).Char(Player(n).CharNum).Level
                Next X
            Else
                ExpG = Exp / q
            End If
            
            For X = 1 To q
                n = Party.PlayerIndex(Player(Attacker).InParty, X)
                If Party.ShareExp(Player(Attacker).InParty) = 2 Then ExpG = Exp * (Player(n).Char(Player(n).CharNum).Level / i)
                If GetPlayerLevel(n) = MAX_LEVEL Then
                    Call SetPlayerExp(n, experience(MAX_LEVEL))
                    Call BattleMsg(n, "Vous ne pouvez pas gagner plus d'expérience!", BrightBlue, 0)
                Else
                    Call SetPlayerExp(n, GetPlayerExp(n) + ExpG)
                    Call BattleMsg(n, "Vous avez gagné " & ExpG & " pts d'expérience (groupe).", BrightBlue, 0)
                End If
            Next X
        End If
                      
        For i = 1 To MAX_NPC_DROPS
            ' Drop the goods if they get it
            n = Int(Rnd * Npc(npcnum).ItemNPC(i).chance) + 1
            If n = 1 Then
                'Call SpawnItem(Npc(npcnum).ItemNPC(i).ItemNum, Npc(npcnum).ItemNPC(i).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y)
                If Player(Attacker).Char(Player(Attacker).CharNum).metier > 0 Then
                    n = Player(Attacker).Char(Player(Attacker).CharNum).metier
                    If metier(n).type = METIER_CHASSEUR Then
                        If InMetier(n, npcnum) <> 10 Then
                            Math.Randomize
                            n = Math.Round(Math.Rnd * 100)

                            If n > 0 And n <= DoubleDrop(Attacker) Then
                                Call SpawnItem(Npc(npcnum).ItemNPC(i).ItemNum, Npc(npcnum).ItemNPC(i).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y)
                                Call SpawnItem(Npc(npcnum).ItemNPC(i).ItemNum, Npc(npcnum).ItemNPC(i).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y)
                            Else
                                Call SpawnItem(Npc(npcnum).ItemNPC(i).ItemNum, Npc(npcnum).ItemNPC(i).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y)
                            End If
                        Else
                            Call SpawnItem(Npc(npcnum).ItemNPC(i).ItemNum, Npc(npcnum).ItemNPC(i).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y)
                        End If
                    Else
                        Call SpawnItem(Npc(npcnum).ItemNPC(i).ItemNum, Npc(npcnum).ItemNPC(i).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y)
                    End If
                Else
                    Call SpawnItem(Npc(npcnum).ItemNPC(i).ItemNum, Npc(npcnum).ItemNPC(i).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y)
                End If
            End If
        Next i
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum, MapNpcNum).num = 0
        MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum, MapNpcNum).HP = 0
        Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
        
        If Val(Scripting) = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "KillNPC " & Attacker & "," & npcnum
        End If
        
        If Player(Attacker).Char(Player(Attacker).CharNum).QueteEnCour > 0 Then
            If quete(Player(Attacker).Char(Player(Attacker).CharNum).QueteEnCour).type = QUETE_TYPE_TUER And Player(Attacker).Char(Player(Attacker).CharNum).QueteStatut(Player(Attacker).Char(Player(Attacker).CharNum).QueteEnCour) = 0 Then
                Call PlayerQueteTypeTuer(Attacker, Player(Attacker).Char(Player(Attacker).CharNum).QueteEnCour, npcnum)
            End If
        End If
        
        ' Check for level up
        Call CheckPlayerLevelUp(Attacker)

        ' Check for level up party member
        If Player(Attacker).InParty > 0 Then
            For X = 1 To Party.MemberCount(Player(Attacker).InParty)
                n = Party.PlayerIndex(Player(Attacker).InParty, X)
                Call CheckPlayerLevelUp(n)
            Next X
        End If
    
        ' Check if target is npc that died and if so set target to 0
        If Player(Attacker).TargetType = TARGET_TYPE_NPC And Player(Attacker).Target = MapNpcNum Then
            Player(Attacker).Target = -1
            Player(Attacker).TargetType = 0
        End If
    Else
        ' NPC not dead, just do the damage
        MapNpc(MapNum, MapNpcNum).HP = MapNpc(MapNum, MapNpcNum).HP - Damage
        
        ' Check for a weapon and say damage
        Call BattleMsg(Attacker, "Tu frappes " & Name & " pour " & Damage & " dommages.", White, 1)
               
        ' Check if we should send a message
        If MapNpc(MapNum, MapNpcNum).Target = 0 And MapNpc(MapNum, MapNpcNum).Target <> Attacker Then
            If Trim$(Npc(npcnum).AttackSay) <> vbNullString Then
                Call QueteMsg(Attacker, Trim$(Npc(npcnum).Name) & " : " & Trim$(Npc(npcnum).AttackSay) & "")
            End If
        End If
        
        ' Set the NPC target to the player
        MapNpc(MapNum, MapNpcNum).Target = Attacker
        
        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum, MapNpcNum).num).Behavior = NPC_BEHAVIOR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum, i).num = MapNpc(MapNum, MapNpcNum).num Then
                    MapNpc(MapNum, i).Target = Attacker
                End If
            Next i
        End If
    End If
    
    ' Reset attack timer
    Player(Attacker).AttackTimer = GetTickCount

Exit Sub
er:
On Error Resume Next
If Attacker < 0 Or Attacker > MAX_PLAYERS Then Exit Sub
Call PlayerMsg(Attacker, "Attaque du PNJ annulée à cause d'une erreur si le problème persiste contactez un administrateur.", Red)
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur dans l'attaque d'un PNJ(" & npcnum & ")par un joueur(" & Player(Attacker).Login & "). Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur dans l'attaque d'un PNJ(" & npcnum & ")par un joueur(" & GetPlayerName(Attacker) & ")", BrightRed, True)
End Sub

Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
Dim Packet As String
Dim OldMap As Long

    On Error GoTo er:

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
    'Call SendWornEquipment(Index)
    'PAPERDOLL
    Dim i As Long
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(index) Then
            If index <> i Then Call SendInventory(i)
            Call SendWornEquipment(i)
        End If
    Next i
    'FIN PAPERDOLL
Exit Sub
er:
On Error Resume Next
If index < 0 Or index > MAX_PLAYERS Then Exit Sub
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur pendant la téléportation du joueur : " & GetPlayerName(index) & ",Compte : " & GetPlayerLogin(index) & ",Carte : " & MapNum & "(" & X & "," & Y & "). Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur pendant la téléportation du joueur : " & GetPlayerName(index), BrightRed, True)
Call PlainMsg(index, "Erreur du serveur, relancer SVP!(Pour tous problème récurent visiter " & Trim$(GetVar(App.Path & "\Config\.ini", "CONFIG", "WebSite")) & ").", 3)
End Sub

Function canPetMove(ByVal index As Long, ByVal Dir As Byte) As Boolean
    canPetMove = True
    With Player(index).Char(Player(index).CharNum).pet
        Select Case Dir
            Case DIR_UP
                If Map(GetPlayerMap(index)).Tile(.X, .Y - 1).type = TILE_TYPE_BLOCKED Then canPetMove = False
            Case DIR_DOWN
                If Map(GetPlayerMap(index)).Tile(.X, .Y + 1).type = TILE_TYPE_BLOCKED Then canPetMove = False
            Case DIR_LEFT
                If Map(GetPlayerMap(index)).Tile(.X - 1, .Y).type = TILE_TYPE_BLOCKED Then canPetMove = False
            Case DIR_RIGHT
                If Map(GetPlayerMap(index)).Tile(.X + 1, .Y).type = TILE_TYPE_BLOCKED Then canPetMove = False
        End Select
    End With
End Function

Sub PetMove(ByVal index As Long)
Dim Moved As Byte
    If IsPlaying(index) = False Then Exit Sub
    With Player(index).Char(Player(index).CharNum).pet
        Moved = 0
                
        If GetPlayerX(index) = .X And GetPlayerY(index) = .Y And Moved <> 2 Then Exit Sub
        If GetPlayerX(index) > .X Then
            If canPetMove(index, DIR_RIGHT) And Moved = 0 Then
                Moved = 1
                .X = .X + 1
                .Dir = DIR_RIGHT
            End If
            If .X - GetPlayerX(index) > 2 Then Moved = 2
        ElseIf GetPlayerX(index) < .X Then
            If canPetMove(index, DIR_LEFT) And Moved = 0 Then
                Moved = 1
                .X = .X - 1
                .Dir = DIR_LEFT
            End If
            If .X - GetPlayerX(index) > 2 Then Moved = 2
        End If
        If GetPlayerY(index) > .Y Then
            If canPetMove(index, DIR_DOWN) And Moved = 0 Then
                Moved = 1
                .Y = .Y + 1
                .Dir = DIR_DOWN
            End If
            If GetPlayerY(index) - .Y > 2 Then Moved = 2
        ElseIf GetPlayerY(index) < .Y Then
            
            If canPetMove(index, DIR_UP) And Moved = 0 Then
                Moved = 1
                .Y = .Y - 1
                .Dir = DIR_UP
            End If
            If .Y - GetPlayerY(index) > 2 Then Moved = 2
        End If
           

        If Moved = 2 Then
            .Y = GetPlayerY(index)
            .X = GetPlayerX(index)
            .Dir = GetPlayerDir(index)
            Moved = 0
        End If
        Call SendDataToMap(GetPlayerMap(index), "PLAYERPET" & SEP_CHAR & index & SEP_CHAR & .Dir & SEP_CHAR & .X & SEP_CHAR & .Y & SEP_CHAR & Moved & SEP_CHAR & END_CHAR)
    End With
End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Packet As String
Dim MapNum As Long
Dim X As Long
Dim Y As Long
Dim i As Long
Dim Moved As Byte
    On Error GoTo er:
        
    ' Check for subscript out of range
    If IsPlaying(index) = False Or Dir < DIR_DOWN Or Dir > DIR_UP Or Movement < 1 Or Movement > 2 Then Exit Sub
    Call SetPlayerDir(index, Dir)
    
    Moved = NO
'    Stop
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) > 0 Then
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).type = TILE_TYPE_CBLOCK Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).data1 = Val(GetPlayerClass(index)) Then Moved = YES
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).data2 = Val(GetPlayerClass(index)) Then Moved = YES
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).data3 = Val(GetPlayerClass(index)) Then Moved = YES
                    If Moved = NO Then Moved = NO: Exit Sub
                End If
                
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).type = TILE_TYPE_BLOCK_DIR Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).data1 = Val(GetPlayerDir(index)) Then Moved = YES
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).data2 = Val(GetPlayerDir(index)) Then Moved = YES
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).data3 = Val(GetPlayerDir(index)) Then Moved = YES
                    If Moved = NO Then Moved = NO: Exit Sub
                End If
                
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).type <> TILE_TYPE_KEY Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).type <> TILE_TYPE_DOOR Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).type <> TILE_TYPE_PORTE_CODE) Or ((Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).type = TILE_TYPE_DOOR Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).type = TILE_TYPE_KEY Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).type = TILE_TYPE_PORTE_CODE) And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) - 1) = YES) Then
                        If (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).type = TILE_TYPE_BLOCK_NIVEAUX And GetPlayerLevel(index) < Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).data1) Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).type = TILE_TYPE_BLOCK_MONTURE And AvMonture(index)) Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).type = TILE_TYPE_BLOCK_GUILDE And Trim$(GetPlayerGuild(index)) <> Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).String1)) Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).type = TILE_TYPE_BLOCK_TOIT Then Moved = NO: Exit Sub
                        'Call PlayerPet(Index, 1, OldDir)
                        Call PetMove(index)
                        Call SetPlayerY(index, GetPlayerY(index) - 1)
                                                
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Up > 0 Then
                    'Vérifier si on ne le téléporte pas sur une case bloquer
                    If Map(Map(GetPlayerMap(index)).Up).Tile(GetPlayerX(index), MAX_MAPY).type <> TILE_TYPE_BLOCKED Then
                        Call PlayerWarp(index, Map(GetPlayerMap(index)).Up, GetPlayerX(index), MAX_MAPY)
                        Moved = YES
                    Else
                        Call SendDataTo(index, "NOTWARP" & SEP_CHAR & END_CHAR)
                    End If
                End If
            End If
                    
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) < MAX_MAPY Then
                
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).type = TILE_TYPE_CBLOCK Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).data1 = GetPlayerClass(index) Then Moved = YES
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).data2 = GetPlayerClass(index) Then Moved = YES
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).data3 = GetPlayerClass(index) Then Moved = YES
                    If Moved = NO Then Moved = NO: Exit Sub
                End If
                
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).type = TILE_TYPE_BLOCK_DIR Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).data1 = GetPlayerDir(index) Then Moved = YES
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).data2 = GetPlayerDir(index) Then Moved = YES
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).data3 = GetPlayerDir(index) Then Moved = YES
                    If Moved = NO Then Moved = NO: Exit Sub
                End If
                
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).type <> TILE_TYPE_KEY Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).type <> TILE_TYPE_DOOR Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).type <> TILE_TYPE_PORTE_CODE) Or ((Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).type = TILE_TYPE_DOOR Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).type = TILE_TYPE_KEY Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).type = TILE_TYPE_PORTE_CODE) And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) + 1) = YES) Then
                        If (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).type = TILE_TYPE_BLOCK_NIVEAUX And GetPlayerLevel(index) < Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).data1) Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).type = TILE_TYPE_BLOCK_MONTURE And AvMonture(index)) Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).type = TILE_TYPE_BLOCK_GUILDE And Trim$(GetPlayerGuild(index)) <> Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).String1)) Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).type = TILE_TYPE_BLOCK_TOIT Then Moved = NO: Exit Sub
                        'Call PlayerPet(Index, 1, OldDir)
                        Call PetMove(index)
                        Call SetPlayerY(index, GetPlayerY(index) + 1)
                                                
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Down > 0 Then
                    'Vérifier si on ne le téléporte pas sur une case bloquer
                    If Map(Map(GetPlayerMap(index)).Down).Tile(GetPlayerX(index), 0).type <> TILE_TYPE_BLOCKED Then
                        Call PlayerWarp(index, Map(GetPlayerMap(index)).Down, GetPlayerX(index), 0)
                        Moved = YES
                    Else
                        Call SendDataTo(index, "NOTWARP" & SEP_CHAR & END_CHAR)
                    End If
                End If
            End If
        
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If GetPlayerX(index) > 0 Then
                
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).type = TILE_TYPE_CBLOCK Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).data1 = GetPlayerClass(index) Then Moved = YES
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).data2 = GetPlayerClass(index) Then Moved = YES
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).data3 = GetPlayerClass(index) Then Moved = YES
                    If Moved = NO Then Moved = NO: Exit Sub
                End If
                
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).type = TILE_TYPE_BLOCK_DIR Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).data1 = GetPlayerDir(index) Then Moved = YES
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).data2 = GetPlayerDir(index) Then Moved = YES
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).data3 = GetPlayerDir(index) Then Moved = YES
                    If Moved = NO Then Moved = NO: Exit Sub
                End If
                
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).type <> TILE_TYPE_KEY Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).type <> TILE_TYPE_DOOR Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).type <> TILE_TYPE_PORTE_CODE) Or ((Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).type = TILE_TYPE_DOOR Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).type = TILE_TYPE_KEY Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).type = TILE_TYPE_PORTE_CODE) And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) - 1, GetPlayerY(index)) = YES) Then
                        If (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).type = TILE_TYPE_BLOCK_NIVEAUX And GetPlayerLevel(index) < Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).data1) Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).type = TILE_TYPE_BLOCK_MONTURE And AvMonture(index)) Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).type = TILE_TYPE_BLOCK_GUILDE And Trim$(GetPlayerGuild(index)) <> Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).String1)) Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).type = TILE_TYPE_BLOCK_TOIT Then Moved = NO: Exit Sub
                        'Call PlayerPet(Index, 1, OldDir)
                        Call PetMove(index)
                        Call SetPlayerX(index, GetPlayerX(index) - 1)
                                                                            
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Left > 0 Then
                    'Vérifier si on ne le téléporte pas sur une case bloquer
                    If Map(Map(GetPlayerMap(index)).Left).Tile(MAX_MAPX, GetPlayerY(index)).type <> TILE_TYPE_BLOCKED Then
                        Call PlayerWarp(index, Map(GetPlayerMap(index)).Left, MAX_MAPX, GetPlayerY(index))
                        Moved = YES
                    Else
                        Call SendDataTo(index, "NOTWARP" & SEP_CHAR & END_CHAR)
                    End If
                End If
            End If
        
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If GetPlayerX(index) < MAX_MAPX Then
                
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).type = TILE_TYPE_CBLOCK Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).data1 = GetPlayerClass(index) Then Moved = YES
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).data2 = GetPlayerClass(index) Then Moved = YES
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).data3 = GetPlayerClass(index) Then Moved = YES
                    If Moved = NO Then Moved = NO: Exit Sub
                End If
                
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).type = TILE_TYPE_BLOCK_DIR Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).data1 = GetPlayerDir(index) Then Moved = YES
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).data2 = GetPlayerDir(index) Then Moved = YES
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).data3 = GetPlayerDir(index) Then Moved = YES
                    If Moved = NO Then Moved = NO: Exit Sub
                End If
                
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).type <> TILE_TYPE_KEY Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).type <> TILE_TYPE_DOOR Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).type <> TILE_TYPE_PORTE_CODE) Or ((Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).type = TILE_TYPE_DOOR Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).type = TILE_TYPE_KEY Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).type = TILE_TYPE_PORTE_CODE) And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) + 1, GetPlayerY(index)) = YES) Then
                        If (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).type = TILE_TYPE_BLOCK_NIVEAUX And GetPlayerLevel(index) < Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).data1) Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).type = TILE_TYPE_BLOCK_MONTURE And AvMonture(index)) Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).type = TILE_TYPE_BLOCK_GUILDE And Trim$(GetPlayerGuild(index)) <> Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).String1)) Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).type = TILE_TYPE_BLOCK_TOIT Then Moved = NO: Exit Sub
                        'Call PlayerPet(Index, 1, OldDir)
                        Call PetMove(index)
                        Call SetPlayerX(index, GetPlayerX(index) + 1)
                                                                            
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Right > 0 Then
                    'Vérifier si on ne le téléporte pas sur une case bloquer
                    If Map(Map(GetPlayerMap(index)).Right).Tile(0, GetPlayerY(index)).type <> TILE_TYPE_BLOCKED Then
                        Call PlayerWarp(index, Map(GetPlayerMap(index)).Right, 0, GetPlayerY(index))
                        Moved = YES
                    Else
                        Call SendDataTo(index, "NOTWARP" & SEP_CHAR & END_CHAR)
                    End If
                End If
            End If
    End Select
    
    If GetPlayerX(index) < 0 Or GetPlayerY(index) < 0 Or GetPlayerX(index) > MAX_MAPX Or GetPlayerY(index) > MAX_MAPY Or GetPlayerMap(index) <= 0 Then
        Call HackingAttempt(index, "Joueur en dehors de la carte ou sur aucune carte")
        Exit Sub
    End If
    
    ' verifier si le joueure est bloquer sur une case
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_COFFRE Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_SIGN Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_BLOCK_TOIT Then
        'débloquer le joueur
        Call Debloque(index)
    End If
    
    'healing tiles code
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_HEAL Then
        Call SetPlayerHP(index, GetPlayerMaxHP(index))
        Call SendHP(index)
        Call PlayerMsg(index, "Tu sens ta force revenir peu a peu!", BrightGreen)
    End If
    
    'Check for kill tile, and if so kill them
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_KILL Then
        Call SetPlayerHP(index, 0)
        Call PlayerMsg(index, "Tu sens la mort arriver et tu perds peu a peu tes forces!!", BrightRed)
        
        ' Warp player away
        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & index
        Else
            Call PlayerWarp(index, START_MAP, START_X, START_Y)
        End If
        Call SetPlayerHP(index, GetPlayerMaxHP(index))
        Call SetPlayerMP(index, GetPlayerMaxMP(index))
        Call SetPlayerSP(index, GetPlayerMaxSP(index))
        Call SendHP(index)
        Call SendMP(index)
        Call SendSP(index)
        Moved = YES
    End If

    If GetPlayerX(index) + 1 <= MAX_MAPX Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).type = TILE_TYPE_DOOR Then
            X = GetPlayerX(index) + 1
            Y = GetPlayerY(index)
            
            If TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
    If GetPlayerX(index) - 1 >= 0 Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).type = TILE_TYPE_DOOR Then
            X = GetPlayerX(index) - 1
            Y = GetPlayerY(index)
            
            If TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
    If GetPlayerY(index) - 1 >= 0 Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).type = TILE_TYPE_DOOR Then
            X = GetPlayerX(index)
            Y = GetPlayerY(index) - 1
            
            If TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
    If GetPlayerY(index) + 1 <= MAX_MAPY Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).type = TILE_TYPE_DOOR Then
            X = GetPlayerX(index)
            Y = GetPlayerY(index) + 1
            
            If TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
            
    ' Check to see if the tile is a warp tile, and if so warp them
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_WARP Then
        MapNum = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data1
        X = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data2
        Y = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data3
        Call PlayerWarp(index, MapNum, X, Y)
        'Call PlayerPet(Index, 0, GetPlayerDir(Index))
        Call PetMove(index)
        Moved = YES
    End If
    
    ' Check for key trigger open
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_KEYOPEN Then
        X = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data1
        Y = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data2
        
        If Map(GetPlayerMap(index)).Tile(X, Y).type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = NO Then
            TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = YES
            TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                            
            Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            If Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1) = vbNullString Then
                Call MapMsg(GetPlayerMap(index), "La porte a été ouverte par un mécanisme!", White)
            Else
                Call MapMsg(GetPlayerMap(index), Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1), White)
            End If
            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "key" & SEP_CHAR & END_CHAR)
        End If
    End If
        
    ' Check for shop
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_SHOP Then
       If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data1 > 0 Then
            Call QueteMsg(index, Shop(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data1).JoinSay)
            Call SendTrade(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data1)
        Else
            Call PlayerMsg(index, "Il n'y a pas de magasin ici.", BrightRed)
        End If
    End If
        
    ' Check if player stepped on sprite changing tile
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_SPRITE_CHANGE Then
        If GetPlayerSprite(index) = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data1 Then
            Call PlayerMsg(index, "Tu as déjà ce sprites!", BrightRed)
            Exit Sub
        Else
            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data2 = 0 Then
                Call SendDataTo(index, "spritechange" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            Else
                If item(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data2).type = ITEM_TYPE_CURRENCY Then
                    Call PlayerMsg(index, "Ce sprite vous coûte " & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data3 & " " & Trim$(item(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data2).Name) & "!", Yellow)
                    Call TakeItem(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data2, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data3)
                    Call SendInventory(index)
                Else
                    Call PlayerMsg(index, "Ce sprite vous coûte un " & Trim$(item(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data2).Name) & "!", Yellow)
                    Call TakeItem(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data2, 1)
                    Call SendInventory(index)
                End If
                
                Call SendDataTo(index, "spritechange" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
    
    ' Check if player stepped on class change
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_CLASS_CHANGE Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data2 > -1 Then
            If GetPlayerClass(index) <> Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data2 Then
                Call PlayerMsg(index, "Tu n'as pas la classe requise!", BrightRed)
                Exit Sub
            End If
        End If
        
        If GetPlayerClass(index) = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data1 Then
            Call PlayerMsg(index, "Tu as déjà cette classe!", BrightRed)
        Else
            If Player(index).Char(Player(index).CharNum).Sex = 0 Then
                If GetPlayerSprite(index) = Classe(GetPlayerClass(index)).MaleSprite Then
                    Call SetPlayerSprite(index, Classe(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data1).MaleSprite)
                End If
            Else
                If GetPlayerSprite(index) = Classe(GetPlayerClass(index)).FemaleSprite Then
                    Call SetPlayerSprite(index, Classe(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data1).FemaleSprite)
                End If
            End If
            
            Call SetPlayerStr(index, (Player(index).Char(Player(index).CharNum).STR - Classe(GetPlayerClass(index)).STR))
            Call SetPlayerDEF(index, (Player(index).Char(Player(index).CharNum).def - Classe(GetPlayerClass(index)).def))
            Call SetPlayerMAGI(index, (Player(index).Char(Player(index).CharNum).magi - Classe(GetPlayerClass(index)).magi))
            Call SetPlayerSPEED(index, (Player(index).Char(Player(index).CharNum).Speed - Classe(GetPlayerClass(index)).Speed))
            
            Call SetPlayerClass(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data1)

            Call SetPlayerStr(index, (Player(index).Char(Player(index).CharNum).STR + Classe(GetPlayerClass(index)).STR + GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "CLASSCHANGE", "AddStr")))
            Call SetPlayerDEF(index, (Player(index).Char(Player(index).CharNum).def + Classe(GetPlayerClass(index)).def + GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "CLASSCHANGE", "AddDef")))
            Call SetPlayerMAGI(index, (Player(index).Char(Player(index).CharNum).magi + Classe(GetPlayerClass(index)).magi + GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "CLASSCHANGE", "AddMagi")))
            Call SetPlayerSPEED(index, (Player(index).Char(Player(index).CharNum).Speed + Classe(GetPlayerClass(index)).Speed + GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "CLASSCHANGE", "AddSpeed")))
            
            Dim ItemNum As Long
            ItemNum = Val(GetVar(App.Path & "\" & "Classes\Class" & GetPlayerClass(index) & ".ini", "STARTUP", "Weapon"))
            If item(ItemNum).type = ITEM_TYPE_WEAPON Then
                i = FindOpenInvSlot(index, ItemNum)
                If i > 0 Then
                    Call SetPlayerInvItemNum(index, i, ItemNum)
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + 1)
                    Call SetPlayerInvItemDur(index, i, item(ItemNum).data1)
                    Call SetPlayerWeaponSlot(index, i)
                End If
            End If
            ItemNum = Val(GetVar(App.Path & "\" & "Classes\Class" & GetPlayerClass(index) & ".ini", "STARTUP", "Shield"))
            If item(ItemNum).type = ITEM_TYPE_SHIELD Then
                i = FindOpenInvSlot(index, ItemNum)
                If i > 0 Then
                    Call SetPlayerInvItemNum(index, i, ItemNum)
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + 1)
                    Call SetPlayerInvItemDur(index, i, item(ItemNum).data1)
                    Call SetPlayerShieldSlot(index, i)
                End If
            End If
            ItemNum = Val(GetVar(App.Path & "\" & "Classes\Class" & GetPlayerClass(index) & ".ini", "STARTUP", "Armor"))
            If item(ItemNum).type = ITEM_TYPE_ARMOR Then
                i = FindOpenInvSlot(index, ItemNum)
                If i > 0 Then
                    Call SetPlayerInvItemNum(index, i, ItemNum)
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + 1)
                    Call SetPlayerInvItemDur(index, i, item(ItemNum).data1)
                    Call SetPlayerArmorSlot(index, i)
                End If
            End If
            ItemNum = Val(GetVar(App.Path & "\" & "Classes\Class" & GetPlayerClass(index) & ".ini", "STARTUP", "Helmet"))
            If item(ItemNum).type = ITEM_TYPE_HELMET Then
                i = FindOpenInvSlot(index, ItemNum)
                If i > 0 Then
                    Call SetPlayerInvItemNum(index, i, ItemNum)
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + 1)
                    Call SetPlayerInvItemDur(index, i, item(ItemNum).data1)
                    Call SetPlayerHelmetSlot(index, i)
                End If
            End If
            If item(ItemNum).type = ITEM_TYPE_PET Then
                i = FindOpenInvSlot(index, ItemNum)
                If i > 0 Then
                    Call SetPlayerInvItemNum(index, i, ItemNum)
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + 1)
                    Call SetPlayerInvItemDur(index, i, item(ItemNum).data1)
                    Call SetPlayerPetSlot(index, i)
                End If
            End If
            
            
            Call PlayerMsg(index, "Ta nouvelle classe est " & Trim$(Classe(GetPlayerClass(index)).Name) & "!", BrightGreen)
            
            Call SendStats(index)
            Call SendHP(index)
            Call SendMP(index)
            Call SendSP(index)
            Call SendWornEquipment(index)
            Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & END_CHAR)
        End If
    End If
    
    ' Check if player stepped on notice tile
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_NOTICE Then
        If Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1) <> vbNullString And Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String2) <> vbNullString Then
            Call QueteMsg(index, Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1) & vbCrLf & Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String2))
        ElseIf Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String2) <> vbNullString Then
            Call QueteMsg(index, Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String2))
        End If
        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String3 & SEP_CHAR & END_CHAR)
    End If
    
    ' Check if player stepped on sound tile
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_SOUND Then
        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1 & SEP_CHAR & END_CHAR)
    End If
    
    If Scripting = 1 And Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_SCRIPTED Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & index & "," & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data1
    End If
    
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_CRAFT Then
        If Player(index).Char(Player(index).CharNum).metier > 0 Then
            If metier(Player(index).Char(Player(index).CharNum).metier).type = METIER_CRAFT Then
                Packet = "CRAFT" & SEP_CHAR & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data1 & SEP_CHAR & END_CHAR
                Call SendDataTo(index, Packet)
            Else
                Call PlayerMsg(index, "Votre métier n'est pas un métier de craft!", Red)
            End If
        Else
            Call PlayerMsg(index, "Vous n'avez pas de métier !", Red)
        End If
    End If
    
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_METIER Then
        If Player(index).Char(Player(index).CharNum).metier = 0 Then
            Packet = "NEWMETIER" & SEP_CHAR & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data1 & SEP_CHAR & END_CHAR
        Else
            If Player(index).Char(Player(index).CharNum).metier <> Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data1 Then
                Packet = "REMPLACEMETIER" & SEP_CHAR & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).data1 & SEP_CHAR & END_CHAR
            Else
                Call PlayerMsg(index, "Vous avez déjà ce métier !", Red)
            End If
        End If
        Call SendDataTo(index, Packet)
    End If
    
    ' verifier si le joueure marche sur une case bank
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_BANK Then
        Call QueteMsg(index, Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1))
        Packet = "BANK" & SEP_CHAR & END_CHAR
        Call SendDataTo(index, Packet)
    End If
        
    ' verifier si le joueure marche sur une case de fin de donjon
    If Player(index).Char(Player(index).CharNum).QueteEnCour > 0 Then
        If quete(Player(index).Char(Player(index).CharNum).QueteEnCour).type = QUETE_TYPE_FINIR And Player(index).Char(Player(index).CharNum).QueteStatut(Player(index).Char(Player(index).CharNum).QueteEnCour) = 0 Then
            If GetPlayerMap(index) = quete(Player(index).Char(Player(index).CharNum).QueteEnCour).data3 And GetPlayerX(index) = quete(Player(index).Char(Player(index).CharNum).QueteEnCour).data1 And GetPlayerY(index) = quete(Player(index).Char(Player(index).CharNum).QueteEnCour).data2 Then
                Player(index).Char(Player(index).CharNum).QueteStatut(Player(index).Char(Player(index).CharNum).QueteEnCour) = 1
                Call SendDataTo(index, "FINQUETE" & SEP_CHAR & END_CHAR)
            End If
        End If
    End If

Exit Sub
er:
On Error Resume Next
If index < 0 Or index > MAX_PLAYERS Then Exit Sub
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur pendant le mouvement du joueur : " & GetPlayerName(index) & ",Compte : " & GetPlayerLogin(index) & ",Direction : " & Dir & "(" & Movement & "). Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur pendant le mouvement du joueur : " & GetPlayerName(index), BrightRed, True)
Call PlainMsg(index, "Erreur du serveur, relancer SVP!(Pour tous problème récurent visiter " & Trim$(GetVar(App.Path & "\Config\.ini", "CONFIG", "WebSite")) & ").", 3)
End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir) As Boolean
Dim i As Long, n As Long
Dim X As Long, Y As Long
Dim BX As Long, BY As Long
Dim TmpX As Byte, TmpY As Byte

    On Error GoTo er:
    
    CanNpcMove = False
    
    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_DOWN Or Dir > DIR_UP Then
        Exit Function
    End If
    
    X = MapNpc(MapNum, MapNpcNum).X
    Y = MapNpc(MapNum, MapNpcNum).Y
    
    CanNpcMove = True
    
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If Y > 0 Then
                TmpY = 0
                TmpX = 1
            Else
                CanNpcMove = False
                Exit Function
            End If
                
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If Y < MAX_MAPY Then
                TmpY = 2
                TmpX = 1
            Else
                CanNpcMove = False
                Exit Function
            End If
                
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If X > 0 Then
                TmpY = 1
                TmpX = 0
            Else
                CanNpcMove = False
                Exit Function
            End If
                
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If X < MAX_MAPX Then
                TmpY = 1
                TmpX = 2
            Else
                CanNpcMove = False
                Exit Function
            End If
    End Select
    
    n = Map(MapNum).Tile(X + (TmpX - 1), Y + (TmpY - 1)).type
    
    ' Check to make sure that the tile is walkable
    If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPC_SPAWN And n <> TILE_TYPE_SCRIPTED And n <> TILE_TYPE_TOIT Then
        If n <> TILE_TYPE_NPCAVOID And CLng(Npc(MapNpc(MapNum, MapNpcNum).num).Vol) <> 0 Then CanNpcMove = True: Exit Function
        CanNpcMove = False
        Exit Function
    End If
    
    ' Check to make sure that there is not a player in the way
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X + (TmpX - 1)) And (GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y + (TmpY - 1)) Then
                CanNpcMove = False
                Exit Function
            End If
        End If
    Next i
    
    ' Check to make sure that there is not another npc in the way
    For i = 1 To MAX_MAP_NPCS
        If (i <> MapNpcNum) And (MapNpc(MapNum, i).num > 0) And (MapNpc(MapNum, i).X = MapNpc(MapNum, MapNpcNum).X + (TmpX - 1)) And (MapNpc(MapNum, i).Y = MapNpc(MapNum, MapNpcNum).Y + (TmpY - 1)) Then
            CanNpcMove = False
            Exit Function
        End If
    Next i
Exit Function
er:
CanNpcMove = False
On Error Resume Next
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur du mouvement du PNJ" & MapNpcNum & " sur la carte " & MapNum & ". Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur du mouvement du PNJ" & MapNpcNum & " sur la carte " & MapNum, BrightRed, True)
End Function

Sub NpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Packet As String
Dim X As Long
Dim Y As Long
Dim i As Long
    
    On Error GoTo er:
    
    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_DOWN Or Dir > DIR_UP Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If
    
    MapNpc(MapNum, MapNpcNum).Dir = Dir
    
    Select Case Dir
        Case DIR_UP
            MapNpc(MapNum, MapNpcNum).Y = MapNpc(MapNum, MapNpcNum).Y - 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_DOWN
            MapNpc(MapNum, MapNpcNum).Y = MapNpc(MapNum, MapNpcNum).Y + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_LEFT
            MapNpc(MapNum, MapNpcNum).X = MapNpc(MapNum, MapNpcNum).X - 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_RIGHT
            MapNpc(MapNum, MapNpcNum).X = MapNpc(MapNum, MapNpcNum).X + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    End Select
Exit Sub
er:
On Error Resume Next
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur pendant le mouvement du PNJ" & MapNpcNum & " sur la carte : " & MapNum & ",Direction : " & Dir & "(" & Movement & "). Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur pendant le mouvement du PNJ" & MapNpcNum & " sur la carte : " & MapNum, BrightRed, True)
End Sub

Sub NpcMoveTo(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal Movement As Long, ByVal X As Long, ByVal Y As Long)
Dim Packet As String
Dim i As Long
 
   On Error GoTo er:
        
    'Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Sub
    End If
    
If Map(MapNum).Npcs(MapNpcNum).boucle = 0 Or Map(MapNum).Npcs(MapNpcNum).Axy = True Then

    If X > MapNpc(MapNum, MapNpcNum).X Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X + 1 And GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y And CLng(Npc(MapNpc(MapNum, MapNpcNum).num).Vol) = 0 Then Exit Sub
        Next i
        If Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X + 1, MapNpc(MapNum, MapNpcNum).Y).type = TILE_TYPE_WALKABLE Or Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X + 1, MapNpc(MapNum, MapNpcNum).Y).type = TILE_TYPE_ITEM Or (Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X + 1, MapNpc(MapNum, MapNpcNum).Y).type <> TILE_TYPE_NPCAVOID Or CLng(Npc(MapNpc(MapNum, MapNpcNum).num).Vol) <> 0) Then
            If Not CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then Exit Sub
            MapNpc(MapNum, MapNpcNum).Dir = DIR_RIGHT
            MapNpc(MapNum, MapNpcNum).X = MapNpc(MapNum, MapNpcNum).X + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
            Exit Sub
        End If
    End If
        
    If X < MapNpc(MapNum, MapNpcNum).X Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X - 1 And GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y And CLng(Npc(MapNpc(MapNum, MapNpcNum).num).Vol) = 0 Then Exit Sub
        Next i
        If Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X - 1, MapNpc(MapNum, MapNpcNum).Y).type = TILE_TYPE_WALKABLE Or Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X - 1, MapNpc(MapNum, MapNpcNum).Y).type = TILE_TYPE_ITEM Or (Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X - 1, MapNpc(MapNum, MapNpcNum).Y).type <> TILE_TYPE_NPCAVOID Or CLng(Npc(MapNpc(MapNum, MapNpcNum).num).Vol) <> 0) Then
            If Not CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then Exit Sub
            MapNpc(MapNum, MapNpcNum).Dir = DIR_LEFT
            MapNpc(MapNum, MapNpcNum).X = MapNpc(MapNum, MapNpcNum).X - 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
            Exit Sub
        End If
    End If
    
    If MapNpc(MapNum, MapNpcNum).Y < Y Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X And GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y + 1 And CLng(Npc(MapNpc(MapNum, MapNpcNum).num).Vol) = 0 Then Exit Sub
        Next i
        If Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y + 1).type = TILE_TYPE_WALKABLE Or Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y + 1).type = TILE_TYPE_ITEM Or (Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y + 1).type <> TILE_TYPE_NPCAVOID Or CLng(Npc(MapNpc(MapNum, MapNpcNum).num).Vol) <> 0) Then
            If Not CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then Exit Sub
            MapNpc(MapNum, MapNpcNum).Dir = DIR_DOWN
            MapNpc(MapNum, MapNpcNum).Y = MapNpc(MapNum, MapNpcNum).Y + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
            Exit Sub
        End If
    End If
    
    If MapNpc(MapNum, MapNpcNum).Y > Y Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X And GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y - 1 And CLng(Npc(MapNpc(MapNum, MapNpcNum).num).Vol) = 0 Then Exit Sub
        Next i
        If Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y - 1).type = TILE_TYPE_WALKABLE Or Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y - 1).type = TILE_TYPE_ITEM Or (Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y - 1).type <> TILE_TYPE_NPCAVOID Or CLng(Npc(MapNpc(MapNum, MapNpcNum).num).Vol) <> 0) Then
            If Not CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then Exit Sub
            MapNpc(MapNum, MapNpcNum).Dir = DIR_UP
            MapNpc(MapNum, MapNpcNum).Y = MapNpc(MapNum, MapNpcNum).Y - 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
            Exit Sub
        End If
     End If
Else

    If MapNpc(MapNum, MapNpcNum).Y < Y Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X And GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y + 1 And CLng(Npc(MapNpc(MapNum, MapNpcNum).num).Vol) = 0 Then Exit Sub
        Next i
        If Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y + 1).type = TILE_TYPE_WALKABLE Or Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y + 1).type = TILE_TYPE_ITEM Or (Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y + 1).type <> TILE_TYPE_NPCAVOID Or CLng(Npc(MapNpc(MapNum, MapNpcNum).num).Vol) <> 0) Then
            If Not CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then Exit Sub
            MapNpc(MapNum, MapNpcNum).Dir = DIR_DOWN
            MapNpc(MapNum, MapNpcNum).Y = MapNpc(MapNum, MapNpcNum).Y + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
            Exit Sub
        End If
    End If
    
    If MapNpc(MapNum, MapNpcNum).Y > Y Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X And GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y - 1 And CLng(Npc(MapNpc(MapNum, MapNpcNum).num).Vol) = 0 Then Exit Sub
        Next i
        If Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y - 1).type = TILE_TYPE_WALKABLE Or Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y - 1).type = TILE_TYPE_ITEM Or (Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y - 1).type <> TILE_TYPE_NPCAVOID Or CLng(Npc(MapNpc(MapNum, MapNpcNum).num).Vol) <> 0) Then
            If Not CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then Exit Sub
            MapNpc(MapNum, MapNpcNum).Dir = DIR_UP
            MapNpc(MapNum, MapNpcNum).Y = MapNpc(MapNum, MapNpcNum).Y - 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
            Exit Sub
        End If
    End If
    
    If X > MapNpc(MapNum, MapNpcNum).X Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X + 1 And GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y And CLng(Npc(MapNpc(MapNum, MapNpcNum).num).Vol) = 0 Then Exit Sub
        Next i
        If Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X + 1, MapNpc(MapNum, MapNpcNum).Y).type = TILE_TYPE_WALKABLE Or Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X + 1, MapNpc(MapNum, MapNpcNum).Y).type = TILE_TYPE_ITEM Or (Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X + 1, MapNpc(MapNum, MapNpcNum).Y).type <> TILE_TYPE_NPCAVOID Or CLng(Npc(MapNpc(MapNum, MapNpcNum).num).Vol) <> 0) Then
            If Not CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then Exit Sub
            MapNpc(MapNum, MapNpcNum).Dir = DIR_RIGHT
            MapNpc(MapNum, MapNpcNum).X = MapNpc(MapNum, MapNpcNum).X + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
            Exit Sub
        End If
    End If
    
    If X < MapNpc(MapNum, MapNpcNum).X Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X - 1 And GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y And CLng(Npc(MapNpc(MapNum, MapNpcNum).num).Vol) = 0 Then Exit Sub
        Next i
        If Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X - 1, MapNpc(MapNum, MapNpcNum).Y).type = TILE_TYPE_WALKABLE Or Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X - 1, MapNpc(MapNum, MapNpcNum).Y).type = TILE_TYPE_ITEM Or (Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X - 1, MapNpc(MapNum, MapNpcNum).Y).type <> TILE_TYPE_NPCAVOID Or CLng(Npc(MapNpc(MapNum, MapNpcNum).num).Vol) <> 0) Then
            If Not CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then Exit Sub
            MapNpc(MapNum, MapNpcNum).Dir = DIR_LEFT
            MapNpc(MapNum, MapNpcNum).X = MapNpc(MapNum, MapNpcNum).X - 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
            Exit Sub
        End If
    End If
End If

Exit Sub
er:
On Error Resume Next
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur pendant le mouvement du PNJ" & MapNpcNum & " sur la carte : " & MapNum & ",Direction : " & Dir & "(" & Movement & ")" & ",Vers(X;Y) : " & X & ";" & Y & ". Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur pendant le mouvement du PNJ" & MapNpcNum & " sur la carte : " & MapNum, BrightRed, True)
End Sub

Sub NpcDir(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)
Dim Packet As String

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_DOWN Or Dir > DIR_UP Then
        Exit Sub
    End If
    
    MapNpc(MapNum, MapNpcNum).Dir = Dir
    Packet = "NPCDIR" & SEP_CHAR & MapNpcNum & SEP_CHAR & Dir & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub JoinGame(ByVal index As Long)
Dim MOTD As String
Dim f As Long
    
    On Error GoTo er:
    
    ' Set the flag so we know the person is in the game
    Player(index).InGame = True
    
    ' Send an ok to client to start receiving in game data
    Call SendDataTo(index, "LOGINOK" & SEP_CHAR & index & SEP_CHAR & END_CHAR)
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(index)
    Call SendClasses(index)
    Call SendItems(index)
    Call SendPets(index)
    Call SendMetiers(index)
    Call SendRecettes(index)
    Call SendEmoticons(index)
    Call SendArrows(index)
    Call SendNpcs(index)
    Call SendShops(index)
    Call SendSpells(index)
    Call SendQuetes(index)
    Call SendInventory(index)
    Call SendWornEquipment(index)
    Call SendHP(index)
    Call SendMP(index)
    Call SendSP(index)
    Call SendStats(index)
    Call SendWeatherTo(index)
    Call SendTimeTo(index)
    Call SendOnlineList
    Call SendDataTo(index, "PICVALUE" & SEP_CHAR & PIC_PL & SEP_CHAR & PIC_NPC1 & SEP_CHAR & PIC_NPC2 & SEP_CHAR & AccModo & SEP_CHAR & AccMapeur & SEP_CHAR & AccDevelopeur & SEP_CHAR & AccAdmin & SEP_CHAR & END_CHAR)
    Call LoadPlayerQuete(index)
    Call SendPlayerQuete(index)
    Call SendPlayerSpells(index)
    Call SendPlayerMetier(index)
    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    Call SendPlayerData(index)
    
    If Scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "JoinGame " & index
    Else
        MOTD = GetVar("motd.ini", "MOTD", "Msg")
        
        ' Send a global message that he/she joined
        If GetPlayerAccess(index) <= ADMIN_MONITER Then
            Call GlobalMsg(GetPlayerName(index) & " a rejoin " & GAME_NAME & "!", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(index) & " a rejoin " & GAME_NAME & "!", AdminColor)
            Call IBMsg("L'Admin/Modo : " & GetPlayerName(index) & " a rejoin " & GAME_NAME & "!", IBCAdmin, True)
        End If
    
        ' Send them welcome
        Call PlayerMsg(index, "Bienvenue sur " & GAME_NAME & "!", 15)
        
        ' Send motd
        If Trim$(MOTD) <> vbNullString Then Call PlayerMsg(index, "MOTD: " & MOTD, 11)
    End If
    
    ' Send whos online
    Call SendWhosOnline(index)
    Call ShowPLR(index)
    'PAPERDOLL
    Dim i As Long
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(index) Then
            If index <> i Then Call SendInventory(i)
            Call SendWornEquipment(i)
            'Call PlayerPet(i, 0, GetPlayerDir(i))
            Call PetMove(i)
        End If
    Next i
    'FIN PAPERDOLL

    ' Send the flag so they know they can start doing stuff
    Call SendDataTo(index, "INGAME" & SEP_CHAR & END_CHAR)
Exit Sub
er:
On Error Resume Next
If index < 0 Or index > MAX_PLAYERS Then Exit Sub
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur de connexion au jeu, joueur : " & GetPlayerName(index) & ",Compte : " & GetPlayerLogin(index) & ". Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur de connexion au jeu, joueur : " & GetPlayerName(index), BrightRed, True)
Call PlainMsg(index, "Erreur du serveur, relancer SVP!(Pour tous problème récurent visiter " & Trim$(GetVar(App.Path & "\Config\.ini", "CONFIG", "WebSite")) & ").", 3)
End Sub

Sub LeftGame(ByVal index As Long)
Dim n As Long
    
    On Error GoTo er:
        
    If bouclier(index) Then bouclier(index) = False: BouclierT(index) = 0
    If Para(index) Then Call ContrOnOff(index): Para(index) = False: ParaT(index) = 0
    If Point(index) > 0 And Point(index) < MAX_SPELLS Then
        If Spell(Point(index)).type = SPELL_TYPE_AMELIO Then
            Player(index).Char(Player(index).CharNum).def = Player(index).Char(Player(index).CharNum).def - Val(Spell(Point(index)).data3)
            Player(index).Char(Player(index).CharNum).magi = Player(index).Char(Player(index).CharNum).magi - Val(Spell(Point(index)).data3)
            Player(index).Char(Player(index).CharNum).STR = Player(index).Char(Player(index).CharNum).STR - Val(Spell(Point(index)).data3)
            Player(index).Char(Player(index).CharNum).Speed = Player(index).Char(Player(index).CharNum).Speed - Val(Spell(Point(index)).data3)
            Call SendStats(index)
            Point(index) = 0
            PointT(index) = 0
        ElseIf Spell(Point(index)).type = SPELL_TYPE_DECONC And GetTickCount >= PointT(index) Then
            Player(index).Char(Player(index).CharNum).def = Player(index).Char(Player(index).CharNum).def + Val(Spell(Point(index)).data3)
            Player(index).Char(Player(index).CharNum).magi = Player(index).Char(Player(index).CharNum).magi + Val(Spell(Point(index)).data3)
            Player(index).Char(Player(index).CharNum).STR = Player(index).Char(Player(index).CharNum).STR + Val(Spell(Point(index)).data3)
            Player(index).Char(Player(index).CharNum).Speed = Player(index).Char(Player(index).CharNum).Speed + Val(Spell(Point(index)).data3)
            Call SendStats(index)
            Point(index) = 0
            PointT(index) = 0
        End If
    End If
    
    If Player(index).InGame Then
        
        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(index)) = 0 Then PlayersOnMap(GetPlayerMap(index)) = NO
        
        Player(index).InGame = False
        
        ' Check if the player was in a party, and if so cancel it out so the other player doesn't continue to get half exp
        Party.RemoveMember Player(index).InParty, Player(index).PartyPlayer
        
        ' Check for boot map
        If Map(GetPlayerMap(index)).BootMap > 0 Then
            Call SetPlayerX(index, Map(GetPlayerMap(index)).BootX)
            Call SetPlayerY(index, Map(GetPlayerMap(index)).BootY)
            Call SetPlayerMap(index, Map(GetPlayerMap(index)).BootMap)
        End If
            
        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "LeftGame " & index
        Else
            ' Send a global message that he/she left
            If GetPlayerAccess(index) <= 1 Then
                Call GlobalMsg(GetPlayerName(index) & " a quitté " & GAME_NAME & "!", 7)
            Else
                Call GlobalMsg(GetPlayerName(index) & " a quitté " & GAME_NAME & "!", 15)
            End If
        End If
        'If quete(Player(Index).Char(Player(Index).CharNum).QueteEnCour).temps > 0 Then
        '    Player(Index).Char(Player(Index).CharNum).QueteStatut(Player(Index).Char(Player(Index).CharNum).QueteEnCour) = 0
        'Else
        '    Player(Index).Char(Player(Index).CharNum).QueteEnCour = 0
        'End If
        Call SavePlayer(index)
        
        Call TextAdd(frmServer.txtText(0), GetPlayerName(index) & " est déconnecté de " & GAME_NAME & ".", True)
        Call SendLeftGame(index)
        'Call RemovePLR
        'For N = 1 To MAX_PLAYERS
        '   Call ShowPLR(N)
        'Next N
    End If
    Call ClearPlayer(index)
    Call SendOnlineList
Exit Sub
er:
On Error Resume Next
If index < 0 Or index > MAX_PLAYERS Then Exit Sub
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur de déconnexion au jeu, joueur : " & GetPlayerName(index) & ",Compte : " & GetPlayerLogin(index) & ". Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur de déconnexion au jeu, joueur : " & GetPlayerName(index), BrightRed, True)
Call PlainMsg(index, "Erreur du serveur, relancer SVP!(Pour tous problème récurent visiter " & Trim$(GetVar(App.Path & "\Config\.ini", "CONFIG", "WebSite")) & ").", 3)
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
Dim i As Long, n As Long

    n = 0
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then n = n + 1
    Next i
    
    GetTotalMapPlayers = n
End Function

Function GetNpcMaxHP(ByVal npcnum As Long)

    ' Prevent subscript out of range
    If npcnum <= 0 Or npcnum > MAX_NPCS Then GetNpcMaxHP = 0: Exit Function
        
    GetNpcMaxHP = Npc(npcnum).MaxHp
End Function

Function GetNpcMaxMP(ByVal npcnum As Long)
    ' Prevent subscript out of range
    If npcnum <= 0 Or npcnum > MAX_NPCS Then GetNpcMaxMP = 0: Exit Function
        
    GetNpcMaxMP = Npc(npcnum).magi * 2
End Function

Function GetNpcMaxSP(ByVal npcnum As Long)
    ' Prevent subscript out of range
    If npcnum <= 0 Or npcnum > MAX_NPCS Then GetNpcMaxSP = 0: Exit Function
        
    GetNpcMaxSP = Npc(npcnum).Speed * 2
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

Function GetNpcHPRegen(ByVal npcnum As Long)
Dim i As Long

    'Prevent subscript out of range
    If npcnum <= 0 Or npcnum > MAX_NPCS Then GetNpcHPRegen = 0: Exit Function
    
    i = (Npc(npcnum).def \ 3)
    If i < 1 Then i = 1
    
    GetNpcHPRegen = i
End Function

Sub CheckPlayerLevelUp(ByVal index As Long)
Dim i As Long
Dim D As Long
Dim C As Long
    C = 0
    
    On Error GoTo er:
    
    If GetPlayerExp(index) >= GetPlayerNextLevel(index) Then
        If GetPlayerLevel(index) < MAX_LEVEL Then
            If Scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerLevelUp " & index
            Else
                Do Until GetPlayerExp(index) < GetPlayerNextLevel(index)
                    DoEvents
                    If GetPlayerLevel(index) < MAX_LEVEL Then
                        If GetPlayerExp(index) >= GetPlayerNextLevel(index) Then
                            D = GetPlayerExp(index) - GetPlayerNextLevel(index)
                            Call SetPlayerLevel(index, GetPlayerLevel(index) + 1)
                            i = (GetPlayerSPEED(index) \ 10)
                            If i < 1 Then i = 1
                            If i > 3 Then i = 3
                                
                            Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + i)
                            Call SetPlayerExp(index, D)
                            C = C + 1
                        End If
                    End If
                Loop
                If C > 1 Then
                    Call GlobalMsg(GetPlayerName(index) & " a gagné " & C & " niveaux!", 6)
                Else
                    Call GlobalMsg(GetPlayerName(index) & " a gagné un niveau!", 6)
                End If
                Call BattleMsg(index, "Vous avez " & GetPlayerPOINTS(index) & " points de stats.", 9, 0)
            End If
            Call SendDataToMap(GetPlayerMap(index), "levelup" & SEP_CHAR & index & SEP_CHAR & END_CHAR)
        End If
        
        If GetPlayerLevel(index) = MAX_LEVEL Then
            Call SetPlayerExp(index, experience(MAX_LEVEL))
        End If
    End If
    
    Call SendHP(index)
    Call SendMP(index)
    Call SendSP(index)
    Call SendStats(index)
Exit Sub
er:
On Error Resume Next
If index < 0 Or index > MAX_PLAYERS Then Exit Sub
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur lors de la vérification du niveau du joueur : " & GetPlayerName(index) & ",Compte : " & GetPlayerLogin(index) & ". Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur lors de la vérification du niveau du joueur : " & GetPlayerName(index), BrightRed, True)
End Sub

Sub CastSpell(ByVal index As Long, ByVal SpellSlot As Long)
Dim SpellNum As Long, i As Long, n As Long, Damage As Long
Dim Casted As Boolean

    Casted = False
    
    On Error GoTo er:
    
    ' Prevent subscript out of range
    If SpellSlot <= 0 Or SpellSlot > MAX_PLAYER_SPELLS Then Exit Sub
    If Not IsPlaying(index) Then Exit Sub
        
    SpellNum = GetPlayerSpell(index, SpellSlot)
    
    ' Make sure player has the spell
    If Not HasSpell(index, SpellNum) Then
        Call BattleMsg(index, "Vous n'avez pas ce sort!", BrightRed, 0)
        Exit Sub
    End If
    
    i = GetSpellReqLevel(index, SpellNum)

    ' Check if they have enough MP
    If GetPlayerMP(index) < Spell(SpellNum).MPCost Then
        Call BattleMsg(index, "Pas assez de mana!", BrightRed, 0)
        Exit Sub
    End If
        
    ' Make sure they are the right level
    If i > GetPlayerLevel(index) Then
        Call BattleMsg(index, "Vous devez étre niveau " & i & " pour lancer ce sort.", BrightRed, 0)
        Exit Sub
    End If
    
    ' Check if timer is ok
    If GetTickCount < Player(index).AttackTimer + 1000 Then Exit Sub
    
    ' Check if the spell is a give item and do that instead of a stat modification
    If Spell(SpellNum).type = 15 Then 'SPELL_TYPE_GIVEITEM Then
        n = FindOpenInvSlot(index, Spell(SpellNum).data1)
        
        If n > 0 Then
            Call GiveItem(index, Spell(SpellNum).data1, Spell(SpellNum).data2)
            'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & trim$(Spell(SpellNum).Name) & ".", BrightBlue)
            
            ' Take away the mana points
            Call SetPlayerMP(index, GetPlayerMP(index) - Spell(SpellNum).MPCost)
            Call SendMP(index)
            Casted = True
        Else
            Call PlayerMsg(index, "Votre inventaire est plein!", BrightRed)
        End If
        
        Exit Sub
    End If
        
Dim X As Long, Y As Long

If Spell(SpellNum).AE = 1 Then
    For Y = GetPlayerY(index) - Spell(SpellNum).Range To GetPlayerY(index) + Spell(SpellNum).Range
        For X = GetPlayerX(index) - Spell(SpellNum).Range To GetPlayerX(index) + Spell(SpellNum).Range
            n = -1
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) = True Then
                    If GetPlayerMap(index) = GetPlayerMap(i) Then
                        If GetPlayerX(i) = X And GetPlayerY(i) = Y Then
                            If i = index Then
                                If Spell(SpellNum).type = SPELL_TYPE_ADDHP Or Spell(SpellNum).type = SPELL_TYPE_ADDMP Or Spell(SpellNum).type = SPELL_TYPE_ADDSP Then
                                    Player(index).Target = i
                                    Player(index).TargetType = TARGET_TYPE_PLAYER
                                    n = Player(index).Target
                                End If
                            Else
                                Player(index).Target = i
                                Player(index).TargetType = TARGET_TYPE_PLAYER
                                n = Player(index).Target
                            End If
                        End If
                    End If
                End If
            Next i
            
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(GetPlayerMap(index), i).num > 0 Then
                    If MapNpc(GetPlayerMap(index), i).X = X And MapNpc(GetPlayerMap(index), i).Y = Y Then
                        If Npc(MapNpc(GetPlayerMap(index), i).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(index), i).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(MapNpc(GetPlayerMap(index), i).num).Behavior <> NPC_BEHAVIOR_QUETEUR Then
                            Player(index).Target = i
                            Player(index).TargetType = TARGET_TYPE_NPC
                            n = Player(index).Target
                        End If
                    End If
                End If
            Next i
                
        Casted = False
        If n > 0 Then
            If Player(index).TargetType = TARGET_TYPE_PLAYER Then
                If IsPlaying(n) Then
                    If bouclier(n) = True Then
                        Call BattleMsg(index, "Le sort ne peut pas toucher le joueur car il a un bouclier!", BrightRed, 0)
                        Exit Sub
                    End If
'                    If n <> Index Then
                        Player(index).TargetType = TARGET_TYPE_PLAYER
                        If GetPlayerHP(n) > 0 And GetPlayerMap(index) = GetPlayerMap(n) And GetPlayerLevel(index) >= NOOB_LEVEL And GetPlayerLevel(n) >= NOOB_LEVEL And (Map(GetPlayerMap(index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(n) <= 0 Then
                            Select Case Spell(SpellNum).type
                                
                                Case SPELL_TYPE_SUBHP
                                    Damage = ((GetPlayerMAGI(index) \ 4) + Spell(SpellNum).data1) - GetPlayerProtection(n)
                                    If Damage > 0 Then Call AttackPlayer(index, n, Damage) Else Call BattleMsg(index, "Votre sort n'est pas assez puissant pour blesser " & GetPlayerName(n) & "!", BrightRed, 0)
                                                            
                                Case SPELL_TYPE_SUBMP
                                    Call SetPlayerMP(n, GetPlayerMP(n) - Spell(SpellNum).data1)
                                    Call SendMP(n)
                    
                                Case SPELL_TYPE_SUBSP
                                    Call SetPlayerSP(n, GetPlayerSP(n) - Spell(SpellNum).data1)
                                    Call SendSP(n)
                                    
                                Case SPELL_TYPE_SCRIPT
                                    MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & n & "," & Val(Spell(SpellNum).data1)
                                    
                                Case SPELL_TYPE_PARALY
                                    If Not Para(n) Then Call ContrOnOff(n)
                                    Para(n) = True
                                    ParaT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                                
                                Case SPELL_TYPE_DEFENC
                                    bouclier(n) = True
                                    BouclierT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                                
                                Case SPELL_TYPE_AMELIO
                                    If Point(n) > 0 Or PointT(n) > 0 Then Call PlayerMsg(index, "Le joueur est déjà la cible d'un sort d'amélioration.", BrightRed)
                                    Player(n).Char(Player(n).CharNum).def = Player(n).Char(Player(n).CharNum).def + Val(Spell(SpellNum).data3)
                                    Player(n).Char(Player(n).CharNum).magi = Player(n).Char(Player(n).CharNum).magi + Val(Spell(SpellNum).data3)
                                    Player(n).Char(Player(n).CharNum).STR = Player(n).Char(Player(n).CharNum).STR + Val(Spell(SpellNum).data3)
                                    Player(n).Char(Player(n).CharNum).Speed = Player(n).Char(Player(n).CharNum).Speed + Val(Spell(SpellNum).data3)
                                    Call SendStats(n)
                                    Point(n) = SpellNum
                                    PointT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                                    
                                Case SPELL_TYPE_DECONC
                                    If Point(n) > 0 Or PointT(n) > 0 Then Call PlayerMsg(index, "Le joueur est déjà la cible d'un sort de déconcentration.", BrightRed)
                                    Player(n).Char(Player(n).CharNum).def = Player(n).Char(Player(n).CharNum).def - Val(Spell(SpellNum).data3)
                                    Player(n).Char(Player(n).CharNum).magi = Player(n).Char(Player(n).CharNum).magi - Val(Spell(SpellNum).data3)
                                    Player(n).Char(Player(n).CharNum).STR = Player(n).Char(Player(n).CharNum).STR - Val(Spell(SpellNum).data3)
                                    Player(n).Char(Player(n).CharNum).Speed = Player(n).Char(Player(n).CharNum).Speed - Val(Spell(SpellNum).data3)
                                    Call SendStats(n)
                                    Point(n) = SpellNum
                                    PointT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                                    
                            End Select

                            Casted = True
                        Else
                            If GetPlayerMap(index) = GetPlayerMap(n) And Spell(SpellNum).type >= SPELL_TYPE_ADDHP And Spell(SpellNum).type <= SPELL_TYPE_ADDSP Or Spell(SpellNum).type = SPELL_TYPE_SCRIPT Then
                                Select Case Spell(SpellNum).type
                                
                                    Case SPELL_TYPE_ADDHP
                                        Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).data1)
                                        Call SendDataTo(n, "BLITNPCDMG" & SEP_CHAR & Spell(SpellNum).data1 & SEP_CHAR & 0 & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                                        Call SendHP(n)
                                                
                                    Case SPELL_TYPE_ADDMP
                                        Call SetPlayerMP(n, GetPlayerMP(n) + Spell(SpellNum).data1)
                                        Call SendMP(n)
                                
                                    Case SPELL_TYPE_ADDSP
                                        Call SetPlayerSP(n, GetPlayerSP(n) + Spell(SpellNum).data1)
                                        Call SendMP(n)
                                        
                                    Case SPELL_TYPE_SCRIPT
                                        MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & n & "," & Val(Spell(SpellNum).data1)
                                    
                                    Case SPELL_TYPE_PARALY
                                        If Not Para(n) Then Call ContrOnOff(n)
                                        Para(n) = True
                                        ParaT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                                    
                                    Case SPELL_TYPE_DEFENC
                                        bouclier(n) = True
                                        BouclierT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                                    
                                    Case SPELL_TYPE_AMELIO
                                        If Point(n) > 0 Or PointT(n) > 0 Then Call PlayerMsg(index, "Le joueur est déjà la cible d'un sort d'amélioration.", BrightRed)
                                        Player(n).Char(Player(n).CharNum).def = Player(n).Char(Player(n).CharNum).def + Val(Spell(SpellNum).data3)
                                        Player(n).Char(Player(n).CharNum).magi = Player(n).Char(Player(n).CharNum).magi + Val(Spell(SpellNum).data3)
                                        Player(n).Char(Player(n).CharNum).STR = Player(n).Char(Player(n).CharNum).STR + Val(Spell(SpellNum).data3)
                                        Player(n).Char(Player(n).CharNum).Speed = Player(n).Char(Player(n).CharNum).Speed + Val(Spell(SpellNum).data3)
                                        Call SendStats(n)
                                        Point(n) = SpellNum
                                        PointT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                                    
                                    Case SPELL_TYPE_DECONC
                                        If Point(n) > 0 Or PointT(n) > 0 Then Call PlayerMsg(index, "Le joueur est déjà la cible d'un sort de déconcentration.", BrightRed)
                                        Player(n).Char(Player(n).CharNum).def = Player(n).Char(Player(n).CharNum).def - Val(Spell(SpellNum).data3)
                                        Player(n).Char(Player(n).CharNum).magi = Player(n).Char(Player(n).CharNum).magi - Val(Spell(SpellNum).data3)
                                        Player(n).Char(Player(n).CharNum).STR = Player(n).Char(Player(n).CharNum).STR - Val(Spell(SpellNum).data3)
                                        Player(n).Char(Player(n).CharNum).Speed = Player(n).Char(Player(n).CharNum).Speed - Val(Spell(SpellNum).data3)
                                        Call SendStats(n)
                                        Point(n) = SpellNum
                                        PointT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                                    
                                End Select
                                
                                Casted = True
                            Else
                                Call PlayerMsg(index, "Vous n'avez pas pu envoyer le sort!(la cible n'est pas sur la même carte que vous)", BrightRed)
                            End If
                        End If
                    'Else
                      '  Player(Index).TargetType = TARGET_TYPE_PLAYER
                     '   If GetPlayerHP(n) > 0 And GetPlayerMap(Index) = GetPlayerMap(n) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(n) >= 10 And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(n) <= 0 Then
                       ' Else
                       '     If GetPlayerMap(Index) = GetPlayerMap(n) And Spell(SpellNum).Type >= SPELL_TYPE_ADDHP And Spell(SpellNum).Type <= SPELL_TYPE_ADDSP Or Spell(SpellNum).Type = SPELL_TYPE_SCRIPT Then
                       '         Select Case Spell(SpellNum).Type
                       '
                       '             Case SPELL_TYPE_ADDHP
                       '                 'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                       '                 Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).Data1)
                       '                 Call SendHP(n)
                       '
                       '             Case SPELL_TYPE_ADDMP
                       '                 'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                       '                 Call SetPlayerMP(n, GetPlayerMP(n) + Spell(SpellNum).Data1)
                       '                 Call SendMP(n)
                       '
                       '             Case SPELL_TYPE_ADDSP
                       '                 'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                       '                 Call SetPlayerMP(n, GetPlayerSP(n) + Spell(SpellNum).Data1)
                       '                 Call SendMP(n)
                       '             Case SPELL_TYPE_SCRIPT
                       '             MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & Index & "," & Val(Spell(SpellNum).Data1)
                     '
                      '          End Select

                     '           Casted = True
                     '       Else
                     '           Call BattleMsg(Index, "Vous n'avez pas put envoyer le sort!", BrightRed, 0)
                     '       End If
                     '   End If
                    'End If
                Else
                    Call BattleMsg(index, "Vous n'avez pas put envoyer le sort!(la cible n'est pas/plus en jeu)", BrightRed, 0)
                End If
            Else
                Player(index).TargetType = TARGET_TYPE_NPC
                If Npc(MapNpc(GetPlayerMap(index), n).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(index), n).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(MapNpc(GetPlayerMap(index), n).num).Behavior <> NPC_BEHAVIOR_QUETEUR And Npc(MapNpc(GetPlayerMap(index), n).num).Behavior <> NPC_BEHAVIOR_SCRIPT Then
                    If Spell(SpellNum).type >= SPELL_TYPE_SUBHP And Spell(SpellNum).type <= SPELL_TYPE_SUBSP Or Spell(SpellNum).type = SPELL_TYPE_PARALY Then
                        Select Case Spell(SpellNum).type
                            
                            Case SPELL_TYPE_SUBHP
                                Damage = ((GetPlayerMAGI(index) \ 4) + Spell(SpellNum).data1) - (Npc(MapNpc(GetPlayerMap(index), n).num).def \ 2)
                                If Damage > 0 Then Call AttackNpc(index, n, Damage) Else Call BattleMsg(index, "Votre sort n'est pas assez puissant pour blesser " & Trim$(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & "!", BrightRed, 0)
                            Case SPELL_TYPE_SUBMP
                                MapNpc(GetPlayerMap(index), n).MP = MapNpc(GetPlayerMap(index), n).MP - Spell(SpellNum).data1

                            Case SPELL_TYPE_SUBSP
                                MapNpc(GetPlayerMap(index), n).SP = MapNpc(GetPlayerMap(index), n).SP - Spell(SpellNum).data1
                                
                            Case SPELL_TYPE_PARALY
                                Call PNJOnOff(n, GetPlayerMap(index))
                                ParaN(n, GetPlayerMap(index)) = True
                                ParaNT(n, GetPlayerMap(index)) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                            
                        End Select
                        
                        Casted = True
                    Else
                        Casted = False
                    End If
                Else
                    Call BattleMsg(index, "Vous ne lancez pas le sort!!(PNJ amis)", BrightRed, 0)
                End If
            End If
        End If
        If Casted Then
            Call SendDataToMap(GetPlayerMap(index), "spellanim" & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & index & SEP_CHAR & Player(index).TargetType & SEP_CHAR & Player(index).Target & SEP_CHAR & END_CHAR)
            'Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "magic" & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & END_CHAR)
        End If
        Next X
    Next Y
    
    Call SetPlayerMP(index, GetPlayerMP(index) - Spell(SpellNum).MPCost)
    Call SendMP(index)
Else
    n = Player(index).Target
    If n = -1 Then Call PlayerMsg(index, "Vous n'avez pas pu envoyer le sort!(aucune cible)", BrightRed): Exit Sub
    If Player(index).TargetType = TARGET_TYPE_PLAYER Then
        If IsPlaying(n) Then
            If bouclier(n) = True Then
                Call BattleMsg(index, "Le sort ne peut pas toucher le joueur car il a un bouclier!", BrightRed, 0)
                Exit Sub
            End If

            If GetPlayerName(n) <> GetPlayerName(index) Then
                If CInt(Sqr((GetPlayerX(index) - GetPlayerX(n)) ^ 2 + ((GetPlayerY(index) - GetPlayerY(n)) ^ 2))) > Spell(SpellNum).Range Then
                    Call BattleMsg(index, "Vous êtes trop loin pour toucher la cible.", BrightRed, 0)
                    Exit Sub
                End If
            End If
            Player(index).TargetType = TARGET_TYPE_PLAYER
            
            If GetPlayerHP(n) > 0 And GetPlayerMap(index) = GetPlayerMap(n) And GetPlayerLevel(index) >= NOOB_LEVEL And GetPlayerLevel(n) >= NOOB_LEVEL And (Map(GetPlayerMap(index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(n) <= 0 Then
                                        
                Select Case Spell(SpellNum).type

                    Case SPELL_TYPE_SUBHP
                        Damage = ((GetPlayerMAGI(index) \ 4) + Spell(SpellNum).data1) - GetPlayerProtection(n)
                        If Damage > 0 Then Call AttackPlayer(index, n, Damage) Else Call BattleMsg(index, "Votre sort n'est pas assez puissant pour blesser " & GetPlayerName(n) & "!", BrightRed, 0)
                        
                    Case SPELL_TYPE_SUBMP
                        Call SetPlayerMP(n, GetPlayerMP(n) - Spell(SpellNum).data1)
                        Call SendMP(n)
        
                    Case SPELL_TYPE_SUBSP
                        Call SetPlayerSP(n, GetPlayerSP(n) - Spell(SpellNum).data1)
                        Call SendSP(n)
                        
                    Case SPELL_TYPE_SCRIPT
                        MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & n & "," & Val(Spell(SpellNum).data1)
                        
                    Case SPELL_TYPE_PARALY
                        If Not Para(n) Then Call ContrOnOff(n)
                        Para(n) = True
                        ParaT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                                    
                    Case SPELL_TYPE_DEFENC
                        bouclier(n) = True
                        BouclierT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                    
                    Case SPELL_TYPE_AMELIO
                        If Point(n) > 0 Or PointT(n) > 0 Then Call PlayerMsg(index, "Le joueur est déjà la cible d'un sort d'amélioration.", BrightRed)
                        Player(n).Char(Player(n).CharNum).def = Player(n).Char(Player(n).CharNum).def + Val(Spell(SpellNum).data3)
                        Player(n).Char(Player(n).CharNum).magi = Player(n).Char(Player(n).CharNum).magi + Val(Spell(SpellNum).data3)
                        Player(n).Char(Player(n).CharNum).STR = Player(n).Char(Player(n).CharNum).STR + Val(Spell(SpellNum).data3)
                        Player(n).Char(Player(n).CharNum).Speed = Player(n).Char(Player(n).CharNum).Speed + Val(Spell(SpellNum).data3)
                        Call SendStats(n)
                        Point(n) = SpellNum
                        PointT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                                    
                    Case SPELL_TYPE_DECONC
                        If Point(n) > 0 Or PointT(n) > 0 Then Call PlayerMsg(index, "Le joueur est déjà la cible d'un sort de déconcentration.", BrightRed)
                        Player(n).Char(Player(n).CharNum).def = Player(n).Char(Player(n).CharNum).def - Val(Spell(SpellNum).data3)
                        Player(n).Char(Player(n).CharNum).magi = Player(n).Char(Player(n).CharNum).magi - Val(Spell(SpellNum).data3)
                        Player(n).Char(Player(n).CharNum).STR = Player(n).Char(Player(n).CharNum).STR - Val(Spell(SpellNum).data3)
                        Player(n).Char(Player(n).CharNum).Speed = Player(n).Char(Player(n).CharNum).Speed - Val(Spell(SpellNum).data3)
                        Call SendStats(n)
                        Point(n) = SpellNum
                        PointT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                    
                End Select
            
                ' Take away the mana points
                Call SetPlayerMP(index, GetPlayerMP(index) - Spell(SpellNum).MPCost)
                Call SendMP(index)
                Casted = True
            Else
                If GetPlayerMap(index) = GetPlayerMap(n) And Spell(SpellNum).type >= SPELL_TYPE_ADDHP And Spell(SpellNum).type <= SPELL_TYPE_ADDSP Or Spell(SpellNum).type >= SPELL_TYPE_SCRIPT Then
                    Select Case Spell(SpellNum).type
                    
                        Case SPELL_TYPE_ADDHP
                            Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).data1)
                            Call SendDataTo(n, "BLITNPCDMG" & SEP_CHAR & Spell(SpellNum).data1 & SEP_CHAR & 0 & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                            Call SendHP(n)
                                    
                        Case SPELL_TYPE_ADDMP
                            Call SetPlayerMP(n, GetPlayerMP(n) + Spell(SpellNum).data1)
                            Call SendMP(n)
                    
                        Case SPELL_TYPE_ADDSP
                            Call SetPlayerSP(n, GetPlayerSP(n) + Spell(SpellNum).data1)
                            Call SendSP(n)
                            
                        Case SPELL_TYPE_SCRIPT
                            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & n & "," & Val(Spell(SpellNum).data1)
                        
                        Case SPELL_TYPE_PARALY
                            If Not Para(n) Then Call ContrOnOff(n)
                            Para(n) = True
                            ParaT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                            
                        Case SPELL_TYPE_DEFENC
                            bouclier(n) = True
                            BouclierT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                            
                        Case SPELL_TYPE_AMELIO
                            If Point(n) > 0 Or PointT(n) > 0 Then Call PlayerMsg(index, "Le joueur est déjà la cible d'un sort d'amélioration.", BrightRed)
                            Player(n).Char(Player(n).CharNum).def = Player(n).Char(Player(n).CharNum).def + Val(Spell(SpellNum).data3)
                            Player(n).Char(Player(n).CharNum).magi = Player(n).Char(Player(n).CharNum).magi + Val(Spell(SpellNum).data3)
                            Player(n).Char(Player(n).CharNum).STR = Player(n).Char(Player(n).CharNum).STR + Val(Spell(SpellNum).data3)
                            Player(n).Char(Player(n).CharNum).Speed = Player(n).Char(Player(n).CharNum).Speed + Val(Spell(SpellNum).data3)
                            Call SendStats(n)
                            Point(n) = SpellNum
                            PointT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                                    
                        Case SPELL_TYPE_DECONC
                            If Point(n) > 0 Or PointT(n) > 0 Then Call PlayerMsg(index, "Le joueur est déjà la cible d'un sort de déconcentration", BrightRed)
                            Player(n).Char(Player(n).CharNum).def = Player(n).Char(Player(n).CharNum).def - Val(Spell(SpellNum).data3)
                            Player(n).Char(Player(n).CharNum).magi = Player(n).Char(Player(n).CharNum).magi - Val(Spell(SpellNum).data3)
                            Player(n).Char(Player(n).CharNum).STR = Player(n).Char(Player(n).CharNum).STR - Val(Spell(SpellNum).data3)
                            Player(n).Char(Player(n).CharNum).Speed = Player(n).Char(Player(n).CharNum).Speed - Val(Spell(SpellNum).data3)
                            Call SendStats(n)
                            Point(n) = SpellNum
                            PointT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                                    
                    End Select
                    
                    ' Take away the mana points
                    Call SetPlayerMP(index, GetPlayerMP(index) - Spell(SpellNum).MPCost)
                    Call SendMP(index)
                    Casted = True
                Else
                    Call BattleMsg(index, "Vous n'avez pas put envoyer le sort!", BrightRed, 0)
                End If
            End If
        Else
            Call PlayerMsg(index, "Vous n'avez pas put envoyer le sort!(cible hors ligne)", BrightRed)
        End If
    Else
        If CInt(Sqr((GetPlayerX(index) - MapNpc(GetPlayerMap(index), n).X) ^ 2 + ((GetPlayerY(index) - MapNpc(GetPlayerMap(index), n).Y) ^ 2))) > Spell(SpellNum).Range Then
            Call BattleMsg(index, "Vous êtes trop loin pour toucher la cible.", BrightRed, 0)
            Exit Sub
        End If
        
        Player(index).TargetType = TARGET_TYPE_NPC
        
        If Npc(MapNpc(GetPlayerMap(index), n).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(index), n).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(MapNpc(GetPlayerMap(index), n).num).Behavior <> NPC_BEHAVIOR_QUETEUR And Npc(MapNpc(GetPlayerMap(index), n).num).Behavior <> NPC_BEHAVIOR_SCRIPT Then
            
            Select Case Spell(SpellNum).type
                Case SPELL_TYPE_ADDHP
                    MapNpc(GetPlayerMap(index), n).HP = MapNpc(GetPlayerMap(index), n).HP + Spell(SpellNum).data1
                    Call SendDataTo(n, "BLITPLAYERDMG" & SEP_CHAR & Spell(SpellNum).data1 & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                
                Case SPELL_TYPE_SUBHP
                    Damage = ((GetPlayerMAGI(index) \ 4) + Spell(SpellNum).data1) - (Npc(MapNpc(GetPlayerMap(index), n).num).def \ 2)
                    If Damage > 0 Then Call AttackNpc(index, n, Damage) Else Call BattleMsg(index, "Votre sort n'est pas assez puissant pour blesser " & Trim$(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & "!", BrightRed, 0)
                    
                Case SPELL_TYPE_ADDMP
                    MapNpc(GetPlayerMap(index), n).MP = MapNpc(GetPlayerMap(index), n).MP + Spell(SpellNum).data1
                
                Case SPELL_TYPE_SUBMP
                    MapNpc(GetPlayerMap(index), n).MP = MapNpc(GetPlayerMap(index), n).MP - Spell(SpellNum).data1
            
                Case SPELL_TYPE_ADDSP
                    MapNpc(GetPlayerMap(index), n).SP = MapNpc(GetPlayerMap(index), n).SP + Spell(SpellNum).data1
                
                Case SPELL_TYPE_SUBSP
                    MapNpc(GetPlayerMap(index), n).SP = MapNpc(GetPlayerMap(index), n).SP - Spell(SpellNum).data1
                
                Case SPELL_TYPE_PARALY
                    Call PNJOnOff(n, GetPlayerMap(index))
                    ParaN(n, GetPlayerMap(index)) = True
                    ParaNT(n, GetPlayerMap(index)) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                    
            End Select
        
            ' Take away the mana points
            Call SetPlayerMP(index, GetPlayerMP(index) - Spell(SpellNum).MPCost)
            Call SendMP(index)
            Casted = True
        Else
            Call BattleMsg(index, "Vous n'avez pas pu envoyer le sort!(cible non ennemi)", BrightRed, 0)
        End If
    End If
End If

    If Casted Then
        Player(index).AttackTimer = GetTickCount
        Player(index).CastedSpell = YES
        Call SendDataToMap(GetPlayerMap(index), "spellanim" & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & index & SEP_CHAR & Player(index).TargetType & SEP_CHAR & Player(index).Target & SEP_CHAR & Player(index).CastedSpell & SEP_CHAR & END_CHAR)
        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "magic" & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & END_CHAR)
    End If
Exit Sub
er:
Casted = False
On Error Resume Next
If index < 0 Or index > MAX_PLAYERS Then Exit Sub
Call PlayerMsg(index, "Le sort n'a pas put être lancé.", BrightRed)
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur de lancement d'un sort du joueur : " & GetPlayerName(index) & ",Compte : " & GetPlayerLogin(index) & ",Slot : " & SpellSlot & ". Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur de lancement d'un sort du joueur : " & GetPlayerName(index), BrightRed, True)
End Sub

Function GetSpellReqLevel(ByVal index As Long, ByVal SpellNum As Long)
    GetSpellReqLevel = Spell(SpellNum).LevelReq
End Function

Function CanPlayerCriticalHit(ByVal index As Long) As Boolean
Dim i As Long, n As Long

    CanPlayerCriticalHit = False
        
    If GetPlayerWeaponSlot(index) > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            i = (GetPlayerStr(index) \ 2) + (GetPlayerLevel(index) \ 2)
    
            n = Int(Rnd * 100) + 1
            If n <= i Then CanPlayerCriticalHit = True
        End If
    End If
End Function

Function CanPlayerBlockHit(ByVal index As Long) As Boolean
Dim i As Long, n As Long, ShieldSlot As Long

    CanPlayerBlockHit = False
    
    ShieldSlot = GetPlayerShieldSlot(index)
    
    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            i = (GetPlayerDEF(index) \ 2) + (GetPlayerLevel(index) \ 2)
        
            n = Int(Rnd * 100) + 1
            If n <= i Then CanPlayerBlockHit = True
        End If
    End If
End Function

Function CanPlayerEsquiveHit(ByVal index As Long) As Boolean
Dim i As Long, n As Long

    CanPlayerEsquiveHit = False
    
        n = Int(Rnd * 2)
        If n = 1 Then
            i = Int(GetPlayerSPEED(index) * 0.576)
        
            n = Int(Rnd * 100) + 1
            If n <= i Then CanPlayerEsquiveHit = True
        End If
    
End Function

Sub CheckEquippedItems(ByVal index As Long)
Dim Slot As Long, ItemNum As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    Slot = GetPlayerWeaponSlot(index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, Slot)
        
        If ItemNum > 0 Then
            If item(ItemNum).type <> ITEM_TYPE_WEAPON Then Call SetPlayerWeaponSlot(index, 0)
        Else
            Call SetPlayerWeaponSlot(index, 0)
        End If
    End If

    Slot = GetPlayerArmorSlot(index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, Slot)
        
        If ItemNum > 0 Then
            If item(ItemNum).type <> ITEM_TYPE_ARMOR And item(ItemNum).type <> ITEM_TYPE_MONTURE Then
                Call SetPlayerArmorSlot(index, 0)
            End If
        Else
            Call SetPlayerArmorSlot(index, 0)
        End If
    End If

    Slot = GetPlayerHelmetSlot(index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, Slot)
        
        If ItemNum > 0 Then
            If item(ItemNum).type <> ITEM_TYPE_HELMET Then Call SetPlayerHelmetSlot(index, 0)
        Else
            Call SetPlayerHelmetSlot(index, 0)
        End If
    End If

    Slot = GetPlayerShieldSlot(index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, Slot)
        
        If ItemNum > 0 Then
            If item(ItemNum).type <> ITEM_TYPE_SHIELD Then Call SetPlayerShieldSlot(index, 0)
        Else
            Call SetPlayerShieldSlot(index, 0)
        End If
    End If
    
    Slot = GetPlayerPetSlot(index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, Slot)
        
        If ItemNum > 0 Then
            If item(ItemNum).type <> ITEM_TYPE_PET Then Call SetPlayerPetSlot(index, 0)
        Else
            Call SetPlayerPetSlot(index, 0)
        End If
    End If
End Sub

Public Sub ShowPLR(ByVal index As Long)
Dim ls As ListItem
On Error Resume Next

    If frmServer.lvUsers.ListItems.Count > 0 And IsPlaying(index) = True Then
        frmServer.lvUsers.ListItems.Remove index
    End If
    Set ls = frmServer.lvUsers.ListItems.add(index, , index)
    
    If IsPlaying(index) = False Then
        ls.SubItems(1) = vbNullString
        ls.SubItems(2) = vbNullString
        ls.SubItems(3) = vbNullString
        ls.SubItems(4) = vbNullString
        ls.SubItems(5) = vbNullString
    Else
        ls.SubItems(1) = GetPlayerLogin(index)
        ls.SubItems(2) = GetPlayerName(index)
        ls.SubItems(3) = GetPlayerLevel(index)
        ls.SubItems(4) = GetPlayerSprite(index)
        ls.SubItems(5) = GetPlayerAccess(index)
    End If
End Sub

Public Sub RemovePLR()
    frmServer.lvUsers.ListItems.Clear
End Sub

Function CanAttackNpcWithArrow(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
Dim MapNum As Long, npcnum As Long
Dim AttackSpeed As Long

On Error GoTo er:
If CLng(Npc(MapNpc(GetPlayerMap(Attacker), MapNpcNum).num).Vol) <> 0 Then CanAttackNpcWithArrow = False: Exit Function
   If GetPlayerWeaponSlot(Attacker) > 0 Then
       AttackSpeed = item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
   Else
       AttackSpeed = 1000
   End If

CanAttackNpcWithArrow = False

' Check For subscript out of range
If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then Exit Function

' Check For subscript out of range
If MapNpc(GetPlayerMap(Attacker), MapNpcNum).num <= 0 Then Exit Function

MapNum = GetPlayerMap(Attacker)
npcnum = MapNpc(MapNum, MapNpcNum).num

' Make sure the npc isn't already dead
If MapNpc(MapNum, MapNpcNum).HP <= 0 And CLng(Npc(npcnum).Inv) = 0 Then Exit Function

If Npc(npcnum).Behavior = NPC_BEHAVIOR_FRIENDLY Or Npc(npcnum).Behavior = NPC_BEHAVIOR_SHOPKEEPER Or Npc(npcnum).Behavior = NPC_BEHAVIOR_QUETEUR And Npc(npcnum).Behavior = NPC_BEHAVIOR_SCRIPT Then
    If Npc(npcnum).Behavior = NPC_BEHAVIOR_QUETEUR And ACoter(MapNpcNum, Attacker) Then Call PlayerMsg(Attacker, "Ne pointe pas cette arme sur moi si tu veut me parler!", BrightRed)
    CanAttackNpcWithArrow = False
    Exit Function
End If

' Make sure they are On the same map
If IsPlaying(Attacker) Then
   If npcnum > 0 And GetTickCount > Player(Attacker).AttackTimer + AttackSpeed Then
       ' Check If at same coordinates
       Select Case GetPlayerDir(Attacker)
           Case DIR_UP
                   If Npc(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(npcnum).Behavior <> NPC_BEHAVIOR_QUETEUR And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SCRIPT Then
                       CanAttackNpcWithArrow = True
                   Else
                       Call PlayerMsg(Attacker, Trim$(Npc(npcnum).Name) & " : " & Trim$(Npc(npcnum).AttackSay), Green)
                   End If

           Case DIR_DOWN
                   If Npc(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(npcnum).Behavior <> NPC_BEHAVIOR_QUETEUR And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SCRIPT Then
                       CanAttackNpcWithArrow = True
                   Else
                       Call PlayerMsg(Attacker, Trim$(Npc(npcnum).Name) & " : " & Trim$(Npc(npcnum).AttackSay), Green)
                   End If

           Case DIR_LEFT
                   If Npc(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(npcnum).Behavior <> NPC_BEHAVIOR_QUETEUR And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SCRIPT Then
                       CanAttackNpcWithArrow = True
                   Else
                       Call PlayerMsg(Attacker, Trim$(Npc(npcnum).Name) & " : " & Trim$(Npc(npcnum).AttackSay), Green)
                   End If

           Case DIR_RIGHT
                   If Npc(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(npcnum).Behavior <> NPC_BEHAVIOR_QUETEUR And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SCRIPT Then
                       CanAttackNpcWithArrow = True

                   Else
                       Call PlayerMsg(Attacker, Trim$(Npc(npcnum).Name) & " : " & Trim$(Npc(npcnum).AttackSay), Green)
               End If
       End Select
   End If
End If
Exit Function
er:
CanAttackNpcWithArrow = False
On Error Resume Next
If Attacker < 0 Or Attacker > MAX_PLAYERS Then Exit Function
Call PlayerMsg(Attacker, "Attaque(flêche) annulée à cause d'une erreur si le problème persiste contactez un administrateur.", Red)
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur dans l'attaque d'un PNJ(" & npcnum & ") par un joueur(" & GetPlayerName(Attacker) & ") avec un arc. Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur dans l'attaque d'un PNJ(" & npcnum & ") par un joueur(" & GetPlayerName(Attacker) & ")avec un arc", BrightRed, True)
End Function

Function CanAttackPlayerWithArrow(ByVal Attacker As Long, ByVal Victim As Long) As Boolean

CanAttackPlayerWithArrow = False

On Error GoTo er:

' Check If map Is attackable
If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
' Make sure they are high enough level
    If GetPlayerLevel(Attacker) < PK_LEVEL Then
        Call PlayerMsg(Attacker, "Vous éte en dessous du niveau " & PK_LEVEL & ",vous ne pouvez pas encore attaquer d'autres joueurs!", BrightRed)
    Else
        If GetPlayerLevel(Victim) < NOOB_LEVEL Then
            Call PlayerMsg(Attacker, GetPlayerName(Victim) & " est en dessous du niveau " & NOOB_LEVEL & ",vous ne pouvez pas encore l'attaquer!", BrightRed)
        Else
            If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                    CanAttackPlayerWithArrow = True
                Else
                    Call PlayerMsg(Attacker, "Vous ne pouvez pas attaquer un menbre de votre guilde!", BrightRed)
                End If
            Else
                CanAttackPlayerWithArrow = True
            End If
        End If
    End If
Else
    Call PlayerMsg(Attacker, "La carte est une safe zone(vous ne pouvez pas attaquer d'autres joueurs)!", BrightRed)
End If

Exit Function
er:
CanAttackPlayerWithArrow = False
On Error Resume Next
If Attacker < 0 Or Attacker > MAX_PLAYERS Or Victim < 0 Or Victim > MAX_PLAYERS Then Exit Function
Call PlayerMsg(Attacker, "Attaque(flêche) annulée à cause d'une erreur si le problème persiste contactez un administrateur.", Red)
Call PlayerMsg(Victim, "Attaque(flêche d'un autre joueur) annulée à cause d'une erreur si le problème persiste contactez un administrateur.", Red)
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur dans l'attaque d'un joueur(" & GetPlayerName(Victim) & ") par un joueur(" & GetPlayerName(Attacker) & ") avec un arc. Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur dans l'attaque d'un joueur(" & GetPlayerName(Victim) & ") par un joueur(" & GetPlayerName(Attacker) & ")avec un arc", BrightRed, True)
End Function

Function AvMonture(ByVal index As Long) As Boolean
    If Not IsPlaying(index) Then Exit Function
    
    If GetPlayerArmorSlot(index) > 0 Then If item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).type = ITEM_TYPE_MONTURE Then AvMonture = True Else AvMonture = False
End Function

Sub Debloque(ByVal index As Long)
Dim Packet As String

If Not IsPlaying(index) Then Exit Sub

On Error Resume Next

If GetPlayerX(index) = MAX_MAPX / 2 And GetPlayerY(index) = MAX_MAPY / 2 Then
    If GetPlayerX(index) + 1 < MAX_MAPX Then
        Call SetPlayerX(index, GetPlayerX(index) + 1)
    ElseIf GetPlayerX(index) - 1 > 0 Then
        Call SetPlayerX(index, GetPlayerX(index) - 1)
    ElseIf GetPlayerY(index) + 1 < MAX_MAPY Then
        Call SetPlayerY(index, GetPlayerY(index) + 1)
    ElseIf GetPlayerY(index) - 1 > 0 Then
        Call SetPlayerY(index, GetPlayerY(index) - 1)
    End If
Else
    Call SetPlayerX(index, MAX_MAPX / 2)
    Call SetPlayerY(index, MAX_MAPY / 2)
End If

Packet = "PLAYERXY" & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & END_CHAR
Call SendDataToMap(GetPlayerMap(index), Packet)

End Sub

Function ACoter(ByVal MapNpcNum As Long, ByVal index As Long) As Boolean
On Error Resume Next

ACoter = False
If index < 1 Or index > MAX_PLAYERS Or MapNpcNum < 1 Or MapNpcNum > 15 Then Exit Function

If GetPlayerX(index) - 1 = MapNpc(GetPlayerMap(index), MapNpcNum).X And GetPlayerY(index) = MapNpc(GetPlayerMap(index), MapNpcNum).Y Then ACoter = True: Exit Function
If GetPlayerX(index) = MapNpc(GetPlayerMap(index), MapNpcNum).X And GetPlayerY(index) - 1 = MapNpc(GetPlayerMap(index), MapNpcNum).Y Then ACoter = True: Exit Function
If GetPlayerX(index) = MapNpc(GetPlayerMap(index), MapNpcNum).X And GetPlayerY(index) + 1 = MapNpc(GetPlayerMap(index), MapNpcNum).Y Then ACoter = True: Exit Function
If GetPlayerX(index) + 1 = MapNpc(GetPlayerMap(index), MapNpcNum).X And GetPlayerY(index) = MapNpc(GetPlayerMap(index), MapNpcNum).Y Then ACoter = True: Exit Function
End Function

Sub EnMonture(ByVal index As Long)
Dim s As Long
s = Val(GetVar(App.Path & "\accounts\" & Trim$(Player(index).Login) & ".ini", "CHAR" & Player(index).CharNum, "monture"))
Call SetPlayerSprite(index, s)
Call SendPlayerData(index)
End Sub

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

Function NbObjet(ByVal index As Long, ByVal ItemNum As Long) As Long
Dim i As Long

    NbObjet = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function
        
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If GetPlayerInvItemValue(index, i) <= 0 Then
                NbObjet = NbObjet + 1
            Else
                NbObjet = NbObjet + GetPlayerInvItemValue(index, i)
            End If
        End If
    Next i
End Function

