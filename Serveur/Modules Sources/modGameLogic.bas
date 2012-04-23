Attribute VB_Name = "modGameLogic"
Option Explicit

Function GetPlayerDamage(ByVal Index As Long) As Long
Dim WeaponSlot As Long

    GetPlayerDamage = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then Exit Function
            
    If GetPlayerPetSlot(Index) > 0 Then
        GetPlayerDamage = ((GetPlayerStr(Index) \ 2) + (Pets(item(GetPlayerInvItemNum(Index, GetPlayerPetSlot(Index))).data1).addForce) \ 2)
    Else
        GetPlayerDamage = (GetPlayerStr(Index) \ 2)
    End If
    
    If GetPlayerDamage <= 0 Then GetPlayerDamage = 1
    
    If GetPlayerWeaponSlot(Index) > 0 Then
        WeaponSlot = GetPlayerWeaponSlot(Index)
        
        GetPlayerDamage = GetPlayerDamage + item(GetPlayerInvItemNum(Index, WeaponSlot)).data2
        
        If GetPlayerInvItemDur(Index, WeaponSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, WeaponSlot, GetPlayerInvItemDur(Index, WeaponSlot) - 1)
        
            If GetPlayerInvItemDur(Index, WeaponSlot) = 0 Then
                Call BattleMsg(Index, "Ton " & Trim$(item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " a été brisé.", Yellow, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, WeaponSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, WeaponSlot) <= 10 Then Call BattleMsg(Index, "Ton " & Trim$(item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " va bientôt se briser! Usure: " & GetPlayerInvItemDur(Index, WeaponSlot) & "/" & Trim$(item(GetPlayerInvItemNum(Index, WeaponSlot)).data1), Yellow, 0)
            End If
        End If
    End If
    
    If GetPlayerDamage < 0 Then GetPlayerDamage = 0
    
End Function

Function GetPlayerProtection(ByVal Index As Long) As Long
Dim ArmorSlot As Long, HelmSlot As Long, ShieldSlot As Long
    
    GetPlayerProtection = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then Exit Function
        
    ArmorSlot = GetPlayerArmorSlot(Index)
    HelmSlot = GetPlayerHelmetSlot(Index)
    ShieldSlot = GetPlayerShieldSlot(Index)
    If GetPlayerPetSlot(Index) > 0 Then
        GetPlayerProtection = ((GetPlayerDEF(Index) \ 4) + (Pets(item(GetPlayerInvItemNum(Index, GetPlayerPetSlot(Index))).data1).addDefence) \ 4)
    Else
        GetPlayerProtection = (GetPlayerDEF(Index) \ 4)
    End If

    If ArmorSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + item(GetPlayerInvItemNum(Index, ArmorSlot)).data2
        If GetPlayerInvItemDur(Index, ArmorSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, ArmorSlot, GetPlayerInvItemDur(Index, ArmorSlot) - 1)
        
            If GetPlayerInvItemDur(Index, ArmorSlot) = 0 Then
                Call BattleMsg(Index, "Ton " & Trim$(item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " a été brisé.", Yellow, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, ArmorSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, ArmorSlot) <= 10 Then Call BattleMsg(Index, "Ton " & Trim$(item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " est endommagé! Usure: " & GetPlayerInvItemDur(Index, ArmorSlot) & "/" & Trim$(item(GetPlayerInvItemNum(Index, ArmorSlot)).data1), Yellow, 0)
            End If
        End If
    End If
    
    If HelmSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + item(GetPlayerInvItemNum(Index, HelmSlot)).data2
        If GetPlayerInvItemDur(Index, HelmSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, HelmSlot, GetPlayerInvItemDur(Index, HelmSlot) - 1)

            If GetPlayerInvItemDur(Index, HelmSlot) <= 0 Then
                Call BattleMsg(Index, "Ton " & Trim$(item(GetPlayerInvItemNum(Index, HelmSlot)).Name) & " a été brisé.", Yellow, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, HelmSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, HelmSlot) <= 10 Then Call BattleMsg(Index, "Ton " & Trim$(item(GetPlayerInvItemNum(Index, HelmSlot)).Name) & " est endommagé! Usure: " & GetPlayerInvItemDur(Index, HelmSlot) & "/" & Trim$(item(GetPlayerInvItemNum(Index, HelmSlot)).data1), Yellow, 0)
            End If
        End If
    End If
    
    If ShieldSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + item(GetPlayerInvItemNum(Index, ShieldSlot)).data2
        If GetPlayerInvItemDur(Index, ShieldSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, ShieldSlot, GetPlayerInvItemDur(Index, ShieldSlot) - 1)

            If GetPlayerInvItemDur(Index, ShieldSlot) <= 0 Then
                Call BattleMsg(Index, "Ton " & Trim$(item(GetPlayerInvItemNum(Index, ShieldSlot)).Name) & " est brisé.", Yellow, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, ShieldSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, ShieldSlot) <= 10 Then Call BattleMsg(Index, "Ton " & Trim$(item(GetPlayerInvItemNum(Index, ShieldSlot)).Name) & " est endommagé! Usure: " & GetPlayerInvItemDur(Index, ShieldSlot) & "/" & Trim$(item(GetPlayerInvItemNum(Index, ShieldSlot)).data1), Yellow, 0)
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

Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim i As Long
    
    FindOpenInvSlot = 0
    
    ' Check for subscript out of range
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function
    
    If item(ItemNum).type = ITEM_TYPE_CURRENCY Or item(ItemNum).Empilable <> 0 Then
        ' If currency then check to see if they already have an guildSoloView of the item and add it to that
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, i) = ItemNum Then FindOpenInvSlot = i: Exit Function
        Next i
    End If
    
    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(Index, i) <= 0 Then FindOpenInvSlot = i: Exit Function
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

Function FindOpenSpellSlot(ByVal Index As Long) As Long
Dim i As Long

    FindOpenSpellSlot = 0
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, i) = 0 Then FindOpenSpellSlot = i: Exit Function
    Next i
End Function

Function HasSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
Dim i As Long

    HasSpell = False
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, i) = SpellNum Then HasSpell = True: Exit Function
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

Function HasItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim i As Long
    
    HasItem = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function
        
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If item(ItemNum).type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(Index, i)
            Else
                HasItem = 1
            End If
            Exit Function
        End If
    Next i
End Function

Sub TakeItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long, n As Long
Dim TakeItem As Boolean

    TakeItem = False
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If item(ItemNum).type = ITEM_TYPE_CURRENCY Or item(ItemNum).Empilable <> 0 Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(Index, i) Then
                    TakeItem = True
                Else
                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - ItemVal)
                    Call SendInventoryUpdate(Index, i)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case item(GetPlayerInvItemNum(Index, i)).type
                    Case ITEM_TYPE_WEAPON
                        If GetPlayerWeaponSlot(Index) > 0 Then
                            If i = GetPlayerWeaponSlot(Index) Then
                                Call SetPlayerWeaponSlot(Index, 0)
                                Call SendInventory(Index)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index)) Then TakeItem = True
                            End If
                        Else
                            TakeItem = True
                        End If
                
                    Case ITEM_TYPE_ARMOR
                        If GetPlayerArmorSlot(Index) > 0 Then
                            If i = GetPlayerArmorSlot(Index) Then
                                Call SetPlayerArmorSlot(Index, 0)
                                Call SendInventory(Index)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index)) Then TakeItem = True
                            End If
                        Else
                            TakeItem = True
                        End If
                    
                    Case ITEM_TYPE_HELMET
                        If GetPlayerHelmetSlot(Index) > 0 Then
                            If i = GetPlayerHelmetSlot(Index) Then
                                Call SetPlayerHelmetSlot(Index, 0)
                                Call SendInventory(Index)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index)) Then TakeItem = True
                            End If
                        Else
                            TakeItem = True
                        End If
                    
                    Case ITEM_TYPE_SHIELD
                        If GetPlayerShieldSlot(Index) > 0 Then
                            If i = GetPlayerShieldSlot(Index) Then
                                Call SetPlayerShieldSlot(Index, 0)
                                Call SendInventory(Index)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index)) Then TakeItem = True
                            End If
                        Else
                            TakeItem = True
                        End If
                    Case ITEM_TYPE_PET
                        If GetPlayerPetSlot(Index) > 0 Then
                            If i = GetPlayerPetSlot(Index) Then
                                Call SetPlayerPetSlot(Index, 0)
                                Call SendInventory(Index)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerPetSlot(Index)) Then TakeItem = True
                            End If
                        Else
                            TakeItem = True
                        End If
                End Select
                
                

                
                n = item(GetPlayerInvItemNum(Index, i)).type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (n <> ITEM_TYPE_WEAPON) And (n <> ITEM_TYPE_ARMOR) And (n <> ITEM_TYPE_HELMET) And (n <> ITEM_TYPE_SHIELD) And (n <> ITEM_TYPE_PET) Then TakeItem = True
            End If
                            
            If TakeItem = True Then
                Call SetPlayerInvItemNum(Index, i, 0)
                Call SetPlayerInvItemValue(Index, i, 0)
                Call SetPlayerInvItemDur(Index, i, 0)
                
                ' Send the inventory update
                Call SendInventoryUpdate(Index, i)
                Exit Sub
            End If
        End If
    Next i
End Sub

Sub GiveItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    i = FindOpenInvSlot(Index, ItemNum)
    
    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(Index, i, ItemNum)
        Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + ItemVal)
        
        If (item(ItemNum).type = ITEM_TYPE_ARMOR) Or (item(ItemNum).type = ITEM_TYPE_WEAPON) Or (item(ItemNum).type = ITEM_TYPE_HELMET) Or (item(ItemNum).type = ITEM_TYPE_SHIELD) Then
            Call SetPlayerInvItemDur(Index, i, item(ItemNum).data1)
        End If
        
        Call SendInventoryUpdate(Index, i)
    Else
        Call PlayerMsg(Index, "Votre inventaire est plein.", BrightRed)
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

Sub PlayerMapGetItem(ByVal Index As Long)
Dim i As Long
Dim n As Long
Dim MapNum As Long
Dim Msg As String


    If IsPlaying(Index) = False Then Exit Sub
    
    MapNum = GetPlayerMap(Index)
    
    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(MapNum, i).Num > 0) And (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
            ' Check if item is at the same location as the player
            If (MapItem(MapNum, i).X = GetPlayerX(Index)) And (MapItem(MapNum, i).Y = GetPlayerY(Index)) Then
                ' Find open slot
                n = FindOpenInvSlot(Index, MapItem(MapNum, i).Num)
                               
                ' Open slot available?
                If n <> 0 Then
                    ' Set item in players inventor
                    Call SetPlayerInvItemNum(Index, n, MapItem(MapNum, i).Num)
                    
                    If item(GetPlayerInvItemNum(Index, n)).type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(Index, n)).Empilable <> 0 Then
                        Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + MapItem(MapNum, i).value)
                        Msg = "Vous ramassez " & MapItem(MapNum, i).value & " " & Trim$(item(GetPlayerInvItemNum(Index, n)).Name) & "."
                    Else
                        Call SetPlayerInvItemValue(Index, n, 1)
                        Msg = "Vous ramassez un " & Trim$(item(GetPlayerInvItemNum(Index, n)).Name) & "."
                    End If
                    
                    If Player(Index).Char(Player(Index).CharNum).QueteEnCour > 0 Then
                        If quete(Player(Index).Char(Player(Index).CharNum).QueteEnCour).type = QUETE_TYPE_RECUP Then
                            Call PlayerQueteTypeRecup(Index, Player(Index).Char(Player(Index).CharNum).QueteEnCour, GetPlayerInvItemNum(Index, n), GetPlayerInvItemValue(Index, n))
                        End If
                    End If
                                            
                    Call SetPlayerInvItemDur(Index, n, MapItem(MapNum, i).Dur)
                        
                    ' Erase item from the map
                    MapItem(MapNum, i).Num = 0
                    MapItem(MapNum, i).value = 0
                    MapItem(MapNum, i).Dur = 0
                    MapItem(MapNum, i).X = 0
                    MapItem(MapNum, i).Y = 0
                        
                    Call SendInventoryUpdate(Index, n)
                    Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
                    Call PlayerMsg(Index, Msg, Yellow)
                    Exit Sub
                Else
                    Call PlayerMsg(Index, "Votre inventaire est plein.", BrightRed)
                    Exit Sub
                End If
            End If
        End If
        
    Next i
End Sub

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
                    If InvNum = GetPlayerArmorSlot(Index) Then
                        Call SetPlayerArmorSlot(Index, 0)
                        Call SendInventory(Index)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                
                Case ITEM_TYPE_WEAPON
                    If InvNum = GetPlayerWeaponSlot(Index) Then
                        Call SetPlayerWeaponSlot(Index, 0)
                        Call SendInventory(Index)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                    
                Case ITEM_TYPE_HELMET
                    If InvNum = GetPlayerHelmetSlot(Index) Then
                        Call SetPlayerHelmetSlot(Index, 0)
                        Call SendInventory(Index)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                                    
                Case ITEM_TYPE_SHIELD
                    If InvNum = GetPlayerShieldSlot(Index) Then
                        Call SetPlayerShieldSlot(Index, 0)
                        Call SendInventory(Index)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                 
                Case ITEM_TYPE_PET
                    If InvNum = GetPlayerPetSlot(Index) Then
                        Call SetPlayerPetSlot(Index, 0)
                        Call SendInventory(Index)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                    
                Case ITEM_TYPE_MONTURE
                    If InvNum = GetPlayerArmorSlot(Index) Then
                        Dim s As Long
                        Call SetPlayerArmorSlot(Index, 0)
                        Call SendInventory(Index)
                        Call SendWornEquipment(Index)
                        s = Val(GetVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "monture"))
                        Call SetPlayerSprite(Index, s)
                        Call SendPlayerData(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
            End Select
                                
            MapItem(GetPlayerMap(Index), i).Num = GetPlayerInvItemNum(Index, InvNum)
            MapItem(GetPlayerMap(Index), i).X = GetPlayerX(Index)
            MapItem(GetPlayerMap(Index), i).Y = GetPlayerY(Index)
                        
            If item(GetPlayerInvItemNum(Index, InvNum)).type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(Index, InvNum)).Empilable <> 0 Then
                ' Check if its more then they have and if so drop it all
                If Amount >= GetPlayerInvItemValue(Index, InvNum) Then
                    MapItem(GetPlayerMap(Index), i).value = GetPlayerInvItemValue(Index, InvNum)
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " dépose un " & GetPlayerInvItemValue(Index, InvNum) & " " & Trim$(item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemNum(Index, InvNum, 0)
                    Call SetPlayerInvItemValue(Index, InvNum, 0)
                    Call SetPlayerInvItemDur(Index, InvNum, 0)
                Else
                    MapItem(GetPlayerMap(Index), i).value = Amount
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " dépose un " & Amount & " " & Trim$(item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
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
            Call PlayerMsg(Index, "Il y a trop d'objets par terre !", BrightRed)
        End If
    End If
End Sub

'Thanks to Gotakk 'cause he's the one who tracked a bug (after almost a year, he's the only one)
'The bug was a forgot to check if we already informed the player about the QuestStatus

Sub PlayerQueteTypeRecup(ByVal Index As Long, ByVal Queteec As Long, ByVal Objnum As Long, ByVal Objvalue As Long)
Dim i As Long
Dim n As Long
Dim z As Long

If Not IsPlaying(Index) Then Exit Sub
If Queteec <= 0 Then Exit Sub
If Objnum <= 0 Or Objnum > MAX_ITEMS Or Objvalue < 0 Then Exit Sub
If GetPlayerQueteEtat(Index, Queteec) Then Exit Sub

For i = 1 To 15
    n = AObjet(Index, quete(Queteec).indexe(i).data1)
    If n > 0 Then
        Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data1 = 1
        Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data2 = NbObjet(Index, quete(Queteec).indexe(i).data1)
    End If

    If quete(Queteec).indexe(i).data2 <= 0 Then
        Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data1 = 1
        Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data2 = 1
    End If
Next i
    
n = 0
For i = 1 To 15
    If Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data1 = 1 And Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data2 >= Val(quete(Queteec).indexe(i).data2) Then
        n = n + 1
    End If
Next i

If n = 15 And Player(Index).Char(Player(Index).CharNum).QueteStatut(Queteec) = 0 Then
    Player(Index).Char(Player(Index).CharNum).QueteStatut(Queteec) = 1
    Call SendDataTo(Index, "FINQUETE" & END_CHAR)
End If

End Sub

Sub PlayerQueteTypeAport(ByVal Index As Long, ByVal Queteec As Long)
Dim i As Long
Dim n As Long
n = 0

If Not IsPlaying(Index) Then Exit Sub
If Queteec <= 0 Then Exit Sub
If quete(Queteec).data1 <= 0 Or quete(Queteec).data1 > MAX_ITEMS Then Exit Sub
If GetPlayerQueteEtat(Index, Queteec) Then Exit Sub

For i = 1 To 24
    If Player(Index).Char(Player(Index).CharNum).Inv(i).Num = quete(Queteec).data1 Then
        Call SetPlayerInvItemNum(Index, i, 0)
        Call SetPlayerInvItemValue(Index, i, 0)
        Call SetPlayerInvItemDur(Index, i, 0)
        Call SendInventory(Index)
        n = 1
        Exit For
    End If
Next i

If n = 1 And Player(Index).Char(Player(Index).CharNum).QueteStatut(Queteec) = 0 Then
    Call QueteMsg(Index, Trim$(quete(Player(Index).Char(Player(Index).CharNum).QueteEnCour).String1))
    Player(Index).Char(Player(Index).CharNum).QueteStatut(Queteec) = 1
    Call SendDataTo(Index, "FINQUETE" & END_CHAR)
Else
    Call QueteMsg(Index, "Je suis désolé tu n'as pas l'objet que je cherche.")
End If

End Sub

Sub PlayerQueteTypeTuer(ByVal Index As Long, ByVal Queteec As Long, ByVal NpcTnum As Long)
Dim i As Long
Dim n As Long

If Not IsPlaying(Index) Then Exit Sub
If Queteec <= 0 Then Exit Sub
If NpcTnum <= 0 Or NpcTnum > MAX_NPCS Then Exit Sub
If GetPlayerQueteEtat(Index, Queteec) = True Then Exit Sub

For i = 1 To 15
    If NpcTnum = quete(Queteec).indexe(i).data1 And quete(Queteec).indexe(i).data2 > 0 And Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data2 < quete(Queteec).indexe(i).data2 Then
        Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data1 = 1
        Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data2 = Val(Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data2) + 1
        Call SendDataTo(Index, "TUERQUETE" & END_CHAR)
    End If
    
    If quete(Queteec).indexe(i).data2 <= 0 Then
        Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data1 = 1
        Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data2 = 0
    End If
Next i

n = 0
For i = 1 To 15
    If Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data1 = 1 And Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data2 >= quete(Queteec).indexe(i).data2 Then
        n = n + 1
    End If
Next i

If n = 15 And Player(Index).Char(Player(Index).CharNum).QueteStatut(Queteec) = 0 Then
    Player(Index).Char(Player(Index).CharNum).QueteStatut(Queteec) = 1
    Call SendDataTo(Index, "FINQUETE" & END_CHAR)
End If

End Sub

Sub PlayerQueteTypeXp(ByVal Index As Long, ByVal Queteec As Long, ByVal Xp As Long)
Dim i As Long
Dim n As Long

If Not IsPlaying(Index) Then Exit Sub
If Queteec <= 0 Or Xp <= 0 Then Exit Sub
If GetPlayerQueteEtat(Index, Queteec) = True Then Exit Sub

If Xp > GetPlayerExp(Index) Then Xp = Xp - GetPlayerExp(Index)

Player(Index).Char(Player(Index).CharNum).Quetep.data1 = Val(Player(Index).Char(Player(Index).CharNum).Quetep.data1) + Val(Xp)
Call SendDataTo(Index, "XPQUETE" & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).Quetep.data1 & END_CHAR)

If Val(Player(Index).Char(Player(Index).CharNum).Quetep.data1) >= Val(quete(Queteec).data1) And Player(Index).Char(Player(Index).CharNum).QueteStatut(Queteec) = 0 Then
    Player(Index).Char(Player(Index).CharNum).QueteStatut(Queteec) = 1
    Call SendDataTo(Index, "FINQUETE" & END_CHAR)
End If

End Sub

Sub TerminerPlayerQuete(ByVal Index As Long, ByVal QueteTindex As Long)
Dim Packet As String
Dim i As Long

If Not IsPlaying(Index) Then Exit Sub
If QueteTindex <= 0 Then Exit Sub
If GetPlayerQueteEtat(Index, QueteTindex) Then Exit Sub

Call ClearPlayerQuete(Index)
Call SendPlayerQuete(Index)

If GetPlayerLevel(Index) = MAX_LEVEL Then
    If quete(QueteTindex).Recompence.Exp > 0 Then
        Call SetPlayerExp(Index, experience(MAX_LEVEL))
        Call BattleMsg(Index, "Tu ne peux pas gagner plus d'expérience!", BrightBlue, 0)
    End If
Else
    If quete(QueteTindex).Recompence.Exp > 0 Then
        Call SetPlayerExp(Index, Player(Index).Char(Player(Index).CharNum).Exp + quete(QueteTindex).Recompence.Exp)
        Call BattleMsg(Index, "Tu as gagné " & quete(QueteTindex).Recompence.Exp & "pts d'expérience.", BrightBlue, 0)
    End If
End If
Call CheckPlayerLevelUp(Index)
Call SendPlayerData(Index)

If quete(QueteTindex).Recompence.objq1 > 0 And quete(QueteTindex).Recompence.objn1 > 0 Then Call GiveItem(Index, quete(QueteTindex).Recompence.objn1, Val(quete(QueteTindex).Recompence.objq1))

If quete(QueteTindex).Recompence.objq2 > 0 And quete(QueteTindex).Recompence.objn2 > 0 Then Call GiveItem(Index, quete(QueteTindex).Recompence.objn2, Val(quete(QueteTindex).Recompence.objq2))

If quete(QueteTindex).Recompence.objq3 > 0 And quete(QueteTindex).Recompence.objn3 > 0 Then Call GiveItem(Index, quete(QueteTindex).Recompence.objn3, Val(quete(QueteTindex).Recompence.objq3))

Call SendPlayerData(Index)
Call SendInventory(Index)

If quete(QueteTindex).Case > 0 Then MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & Index & "," & Val(quete(QueteTindex).Case)

Player(Index).Char(Player(Index).CharNum).QueteStatut(QueteTindex) = 2

Packet = "TERMINEQUETE" & END_CHAR
Call SendDataTo(Index, Packet)

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
                MapNpc(MapNum, MapNpcNum).Num = 0
                MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
                MapNpc(MapNum, MapNpcNum).HP = 0
                Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)
                Exit Sub
            End If
        Else
            If Npc(npcnum).SpawnTime = 2 Then
                MapNpc(MapNum, MapNpcNum).Num = 0
                MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
                MapNpc(MapNum, MapNpcNum).HP = 0
                Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)
                Exit Sub
            End If
        End If
    
        MapNpc(MapNum, MapNpcNum).Num = npcnum
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
            Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & END_CHAR
            Call SendDataToMap(MapNum, Packet)
        End If
    Else
        MapNpc(MapNum, MapNpcNum).Num = 0
        
        Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Num & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & END_CHAR
        Call SendDataToMap(MapNum, Packet)
    End If
    
    'Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(MapNum, MapNpcNum).num) & END_CHAR)
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

If CLng(Npc(MapNpc(GetPlayerMap(Attacker), MapNpcNum).Num).Vol) <> 0 Then CanAttackNpc = False: Exit Function

If GetPlayerWeaponSlot(Attacker) > 0 Then
    AttackSpeed = item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
Else
    AttackSpeed = 1000
End If

CanAttackNpc = False
 
' Check for subscript out of range
If MapNpc(GetPlayerMap(Attacker), MapNpcNum).Num <= 0 Or MapNpc(GetPlayerMap(Attacker), MapNpcNum).Num > MAX_NPCS Then Exit Function
 
MapNum = GetPlayerMap(Attacker)
npcnum = MapNpc(MapNum, MapNpcNum).Num
 
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
    If Npc(npcnum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or Npc(npcnum).Behavior = NPC_BEHAVIOR_ATTACKWHENATTACKED Then
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
                            Call SendDataTo(Attacker, "QUETECOUR" & SEP_CHAR & Npc(npcnum).QueteNum & END_CHAR)
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
                            Call SendDataTo(Attacker, "FINQUETE" & END_CHAR)
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

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
Dim MapNum As Long, npcnum As Long
    
    CanNpcAttackPlayer = False
    
    If Not IsPlaying(Index) Then Exit Function
    
    On Error GoTo er:
    
    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Index) = False Then Exit Function
        
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Index), MapNpcNum).Num <= 0 Or MapNpc(GetPlayerMap(Index), MapNpcNum).Num > MAX_NPCS Then Exit Function
        
    MapNum = GetPlayerMap(Index)
    npcnum = MapNpc(MapNum, MapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum, MapNpcNum).HP <= 0 And CLng(Npc(npcnum).Inv) = 0 Then Exit Function
        
    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(MapNum, MapNpcNum).AttackTimer + 1000 Then Exit Function
        
    ' Make sure we dont attack the player if they are switching maps
    If Player(Index).GettingMap = YES Then Exit Function
    
    MapNpc(MapNum, MapNpcNum).AttackTimer = GetTickCount
    
    ' Make sure they are on the same map
    If IsPlaying(Index) Then
        If npcnum > 0 Then
            ' Check if at same coordinates
            If (GetPlayerY(Index) + 1 = MapNpc(MapNum, MapNpcNum).Y) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).X) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).Y + 1) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).X) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).Y) And (GetPlayerX(Index) + 1 = MapNpc(MapNum, MapNpcNum).X) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).Y) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).X + 1) Then
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
If Index < 0 Or Index > MAX_PLAYERS Then Exit Function
Call PlayerMsg(Index, "Attaque du PNJ annulée à cause d'une erreur si le problème persiste contactez un administrateur.", Red)
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur dans l'attaque d'un joueur(" & Player(Index).Login & ")par un PNJ(" & npcnum & "). Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur dans l'attaque d'un joueur(" & GetPlayerName(Index) & ")par un PNJ(" & npcnum & ")", BrightRed, True)
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
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), "ATTACK" & SEP_CHAR & Attacker & END_CHAR)
    
    ' Effet de sang
    Call SendDataToMap(GetPlayerMap(Attacker), "BloodAnim" & SEP_CHAR & Attacker & SEP_CHAR & Victim & SEP_CHAR & TARGET_TYPE_PLAYER & END_CHAR)

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
    Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "pain" & END_CHAR)

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
    If MapNpc(GetPlayerMap(Victim), MapNpcNum).Num <= 0 Then Exit Sub
            
    ' Send this packet so they can see the person attacking
    Call SendDataToMap(GetPlayerMap(Victim), "NPCATTACK" & SEP_CHAR & MapNpcNum & END_CHAR)
    
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
        'Call SendDataToMap(GetPlayerMap(Victim), "changedir" & SEP_CHAR & GetPlayerDir(Victim) & SEP_CHAR & Victim & END_CHAR)
    'End If
    ':: END AUTO TURN ::
    
    Name = Trim$(Npc(MapNpc(MapNum, MapNpcNum).Num).Name)
    
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
    
    Call SendDataTo(Victim, "BLITNPCDMG" & SEP_CHAR & Damage & SEP_CHAR & 0 & END_CHAR)
    Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "pain" & END_CHAR)

Exit Sub
er:
On Error Resume Next
If Victim < 0 Or Victim > MAX_PLAYERS Then Exit Sub
Call PlayerMsg(Victim, "Attaque (du PNJ qui vous attaque) annulée à cause d'une erreur si le problème persiste contactez un administrateur.", Red)
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur dans l'attaque d'un joueur(" & Player(Victim).Login & ")par un PNJ(" & MapNpc(MapNum, MapNpcNum).Num & "). Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur dans l'attaque d'un joueur(" & GetPlayerName(Victim) & ")par un PNJ(" & MapNpc(MapNum, MapNpcNum).Num & ")", BrightRed, True)
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
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), "ATTACK" & SEP_CHAR & Attacker & END_CHAR)
    
    ' Effet de sang
    Call SendDataToMap(GetPlayerMap(Attacker), "BloodAnim" & SEP_CHAR & Attacker & SEP_CHAR & MapNpcNum & SEP_CHAR & TARGET_TYPE_NPC & END_CHAR)
    
    MapNum = GetPlayerMap(Attacker)
    npcnum = MapNpc(MapNum, MapNpcNum).Num
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
        MapNpc(MapNum, MapNpcNum).Num = 0
        MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum, MapNpcNum).HP = 0
        Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)
        
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
        If Npc(MapNpc(MapNum, MapNpcNum).Num).Behavior = NPC_BEHAVIOR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum, i).Num = MapNpc(MapNum, MapNpcNum).Num Then
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

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
Dim Packet As String
Dim OldMap As Long

    On Error GoTo er:

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
    'Call SendWornEquipment(Index)
    'PAPERDOLL
    Dim i As Long
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(Index) Then
            If Index <> i Then Call SendInventory(i)
            Call SendWornEquipment(i)
        End If
    Next i
    'FIN PAPERDOLL
Exit Sub
er:
On Error Resume Next
If Index < 0 Or Index > MAX_PLAYERS Then Exit Sub
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur pendant la téléportation du joueur : " & GetPlayerName(Index) & ",Compte : " & GetPlayerLogin(Index) & ",Carte : " & MapNum & "(" & X & "," & Y & "). Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur pendant la téléportation du joueur : " & GetPlayerName(Index), BrightRed, True)
Call PlainMsg(Index, "Erreur du serveur, relancer SVP!(Pour tous problème récurent visiter " & Trim$(GetVar(App.Path & "\Config\.ini", "CONFIG", "WebSite")) & ").", 3)
End Sub

Function canPetMove(ByVal Index As Long, ByVal Dir As Byte) As Boolean
    canPetMove = True
    With Player(Index).Char(Player(Index).CharNum).pet
        Select Case Dir
            Case DIR_UP
                If Map(GetPlayerMap(Index)).Tile(.X, .Y - 1).type = TILE_TYPE_BLOCKED Then canPetMove = False
            Case DIR_DOWN
                If Map(GetPlayerMap(Index)).Tile(.X, .Y + 1).type = TILE_TYPE_BLOCKED Then canPetMove = False
            Case DIR_LEFT
                If Map(GetPlayerMap(Index)).Tile(.X - 1, .Y).type = TILE_TYPE_BLOCKED Then canPetMove = False
            Case DIR_RIGHT
                If Map(GetPlayerMap(Index)).Tile(.X + 1, .Y).type = TILE_TYPE_BLOCKED Then canPetMove = False
        End Select
    End With
End Function

Sub PetMove(ByVal Index As Long)
Dim Moved As Byte
    If IsPlaying(Index) = False Then Exit Sub
    With Player(Index).Char(Player(Index).CharNum).pet
        Moved = 0
                
        If GetPlayerX(Index) = .X And GetPlayerY(Index) = .Y And Moved <> 2 Then Exit Sub
        If GetPlayerX(Index) > .X Then
            If canPetMove(Index, DIR_RIGHT) And Moved = 0 Then
                Moved = 1
                .X = .X + 1
                .Dir = DIR_RIGHT
            End If
            If .X - GetPlayerX(Index) > 2 Then Moved = 2
        ElseIf GetPlayerX(Index) < .X Then
            If canPetMove(Index, DIR_LEFT) And Moved = 0 Then
                Moved = 1
                .X = .X - 1
                .Dir = DIR_LEFT
            End If
            If .X - GetPlayerX(Index) > 2 Then Moved = 2
        End If
        If GetPlayerY(Index) > .Y Then
            If canPetMove(Index, DIR_DOWN) And Moved = 0 Then
                Moved = 1
                .Y = .Y + 1
                .Dir = DIR_DOWN
            End If
            If GetPlayerY(Index) - .Y > 2 Then Moved = 2
        ElseIf GetPlayerY(Index) < .Y Then
            
            If canPetMove(Index, DIR_UP) And Moved = 0 Then
                Moved = 1
                .Y = .Y - 1
                .Dir = DIR_UP
            End If
            If .Y - GetPlayerY(Index) > 2 Then Moved = 2
        End If
           

        If Moved = 2 Then
            .Y = GetPlayerY(Index)
            .X = GetPlayerX(Index)
            .Dir = GetPlayerDir(Index)
            Moved = 0
        End If
        Call SendDataToMap(GetPlayerMap(Index), "PLAYERPET" & SEP_CHAR & Index & SEP_CHAR & .Dir & SEP_CHAR & .X & SEP_CHAR & .Y & SEP_CHAR & Moved & END_CHAR)
    End With
End Sub

Sub PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Packet As String
Dim MapNum As Long
Dim X As Long
Dim Y As Long
Dim i As Long
Dim Moved As Byte
    On Error GoTo er:
        
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Dir < DIR_DOWN Or Dir > DIR_UP Or Movement < 1 Or Movement > 2 Then Exit Sub
    Call SetPlayerDir(Index, Dir)
    
    Moved = NO
    Player(Index).Char(Player(Index).CharNum).LastX = GetPlayerX(Index)
    Player(Index).Char(Player(Index).CharNum).LastY = GetPlayerY(Index)
'    Stop
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) > 0 Then
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type = TILE_TYPE_CBLOCK Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).data1 = Val(GetPlayerClass(Index)) Then Moved = YES
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).data2 = Val(GetPlayerClass(Index)) Then Moved = YES
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).data3 = Val(GetPlayerClass(Index)) Then Moved = YES
                    If Moved = NO Then Moved = NO: Exit Sub
                End If
                
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type = TILE_TYPE_BLOCK_DIR Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).data1 = Val(GetPlayerDir(Index)) Then Moved = YES
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).data2 = Val(GetPlayerDir(Index)) Then Moved = YES
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).data3 = Val(GetPlayerDir(Index)) Then Moved = YES
                    If Moved = NO Then Moved = NO: Exit Sub
                End If
                
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type <> TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type <> TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type <> TILE_TYPE_PORTE_CODE) Or ((Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type = TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type = TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type = TILE_TYPE_PORTE_CODE) And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) - 1) = YES) Then
                        If (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type = TILE_TYPE_BLOCK_NIVEAUX And GetPlayerLevel(Index) < Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).data1) Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type = TILE_TYPE_BLOCK_MONTURE And AvMonture(Index)) Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type = TILE_TYPE_BLOCK_GUILDE And Trim$(GetPlayerGuild(Index)) <> Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).String1)) Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type = TILE_TYPE_BLOCK_TOIT Then Moved = NO: Exit Sub
                        'Call PlayerPet(Index, 1, OldDir)
                        Call PetMove(Index)
                        Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                                                
                        Packet = "PLAYERMOVE" & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Up > 0 Then
                    'Vérifier si on ne le téléporte pas sur une case bloquer
                    If Map(Map(GetPlayerMap(Index)).Up).Tile(GetPlayerX(Index), MAX_MAPY).type <> TILE_TYPE_BLOCKED Then
                        Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Up, GetPlayerX(Index), MAX_MAPY)
                        Moved = YES
                    Else
                        Call SendDataTo(Index, "NOTWARP" & END_CHAR)
                    End If
                End If
            End If
                    
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) < MAX_MAPY Then
                
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type = TILE_TYPE_CBLOCK Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).data1 = GetPlayerClass(Index) Then Moved = YES
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).data2 = GetPlayerClass(Index) Then Moved = YES
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).data3 = GetPlayerClass(Index) Then Moved = YES
                    If Moved = NO Then Moved = NO: Exit Sub
                End If
                
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type = TILE_TYPE_BLOCK_DIR Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).data1 = GetPlayerDir(Index) Then Moved = YES
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).data2 = GetPlayerDir(Index) Then Moved = YES
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).data3 = GetPlayerDir(Index) Then Moved = YES
                    If Moved = NO Then Moved = NO: Exit Sub
                End If
                
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type <> TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type <> TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type <> TILE_TYPE_PORTE_CODE) Or ((Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type = TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type = TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type = TILE_TYPE_PORTE_CODE) And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) + 1) = YES) Then
                        If (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type = TILE_TYPE_BLOCK_NIVEAUX And GetPlayerLevel(Index) < Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).data1) Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type = TILE_TYPE_BLOCK_MONTURE And AvMonture(Index)) Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type = TILE_TYPE_BLOCK_GUILDE And Trim$(GetPlayerGuild(Index)) <> Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).String1)) Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type = TILE_TYPE_BLOCK_TOIT Then Moved = NO: Exit Sub
                        'Call PlayerPet(Index, 1, OldDir)
                        Call PetMove(Index)
                        Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                                                
                        Packet = "PLAYERMOVE" & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Down > 0 Then
                    'Vérifier si on ne le téléporte pas sur une case bloquer
                    If Map(Map(GetPlayerMap(Index)).Down).Tile(GetPlayerX(Index), 0).type <> TILE_TYPE_BLOCKED Then
                        Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Down, GetPlayerX(Index), 0)
                        Moved = YES
                    Else
                        Call SendDataTo(Index, "NOTWARP" & END_CHAR)
                    End If
                End If
            End If
        
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) > 0 Then
                
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type = TILE_TYPE_CBLOCK Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).data1 = GetPlayerClass(Index) Then Moved = YES
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).data2 = GetPlayerClass(Index) Then Moved = YES
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).data3 = GetPlayerClass(Index) Then Moved = YES
                    If Moved = NO Then Moved = NO: Exit Sub
                End If
                
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type = TILE_TYPE_BLOCK_DIR Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).data1 = GetPlayerDir(Index) Then Moved = YES
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).data2 = GetPlayerDir(Index) Then Moved = YES
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).data3 = GetPlayerDir(Index) Then Moved = YES
                    If Moved = NO Then Moved = NO: Exit Sub
                End If
                
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type <> TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type <> TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type <> TILE_TYPE_PORTE_CODE) Or ((Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type = TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type = TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type = TILE_TYPE_PORTE_CODE) And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) - 1, GetPlayerY(Index)) = YES) Then
                        If (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type = TILE_TYPE_BLOCK_NIVEAUX And GetPlayerLevel(Index) < Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).data1) Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type = TILE_TYPE_BLOCK_MONTURE And AvMonture(Index)) Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type = TILE_TYPE_BLOCK_GUILDE And Trim$(GetPlayerGuild(Index)) <> Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).String1)) Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type = TILE_TYPE_BLOCK_TOIT Then Moved = NO: Exit Sub
                        'Call PlayerPet(Index, 1, OldDir)
                        Call PetMove(Index)
                        Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                                                                            
                        Packet = "PLAYERMOVE" & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Left > 0 Then
                    'Vérifier si on ne le téléporte pas sur une case bloquer
                    If Map(Map(GetPlayerMap(Index)).Left).Tile(MAX_MAPX, GetPlayerY(Index)).type <> TILE_TYPE_BLOCKED Then
                        Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Left, MAX_MAPX, GetPlayerY(Index))
                        Moved = YES
                    Else
                        Call SendDataTo(Index, "NOTWARP" & END_CHAR)
                    End If
                End If
            End If
        
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) < MAX_MAPX Then
                
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type = TILE_TYPE_CBLOCK Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).data1 = GetPlayerClass(Index) Then Moved = YES
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).data2 = GetPlayerClass(Index) Then Moved = YES
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).data3 = GetPlayerClass(Index) Then Moved = YES
                    If Moved = NO Then Moved = NO: Exit Sub
                End If
                
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type = TILE_TYPE_BLOCK_DIR Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).data1 = GetPlayerDir(Index) Then Moved = YES
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).data2 = GetPlayerDir(Index) Then Moved = YES
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).data3 = GetPlayerDir(Index) Then Moved = YES
                    If Moved = NO Then Moved = NO: Exit Sub
                End If
                
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type <> TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type <> TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type <> TILE_TYPE_PORTE_CODE) Or ((Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type = TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type = TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type = TILE_TYPE_PORTE_CODE) And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) + 1, GetPlayerY(Index)) = YES) Then
                        If (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type = TILE_TYPE_BLOCK_NIVEAUX And GetPlayerLevel(Index) < Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).data1) Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type = TILE_TYPE_BLOCK_MONTURE And AvMonture(Index)) Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type = TILE_TYPE_BLOCK_GUILDE And Trim$(GetPlayerGuild(Index)) <> Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).String1)) Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type = TILE_TYPE_BLOCK_TOIT Then Moved = NO: Exit Sub
                        'Call PlayerPet(Index, 1, OldDir)
                        Call PetMove(Index)
                        Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                                                                            
                        Packet = "PLAYERMOVE" & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Right > 0 Then
                    'Vérifier si on ne le téléporte pas sur une case bloquer
                    If Map(Map(GetPlayerMap(Index)).Right).Tile(0, GetPlayerY(Index)).type <> TILE_TYPE_BLOCKED Then
                        Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Right, 0, GetPlayerY(Index))
                        Moved = YES
                    Else
                        Call SendDataTo(Index, "NOTWARP" & END_CHAR)
                    End If
                End If
            End If
    End Select
    
    If GetPlayerX(Index) < 0 Or GetPlayerY(Index) < 0 Or GetPlayerX(Index) > MAX_MAPX Or GetPlayerY(Index) > MAX_MAPY Or GetPlayerMap(Index) <= 0 Then
        Call HackingAttempt(Index, "Joueur en dehors de la carte ou sur aucune carte")
        Exit Sub
    End If
    
    ' verifier si le joueure est bloquer sur une case
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_COFFRE Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_SIGN Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_BLOCK_TOIT Then
        'débloquer le joueur
        Call Debloque(Index)
    End If
    
    'healing tiles code
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_HEAL Then
        Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
        Call SendHP(Index)
        Call PlayerMsg(Index, "Tu sens ta force revenir peu a peu!", BrightGreen)
    End If
    
    'Check for kill tile, and if so kill them
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_KILL Then
        Call SetPlayerHP(Index, 0)
        Call PlayerMsg(Index, "Tu sens la mort arriver et tu perds peu a peu tes forces!!", BrightRed)
        
        ' Warp player away
        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & Index
        Else
            Call PlayerWarp(Index, START_MAP, START_X, START_Y)
        End If
        Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
        Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
        Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
        Call SendHP(Index)
        Call SendMP(Index)
        Call SendSP(Index)
        Moved = YES
    End If

    If GetPlayerX(Index) + 1 <= MAX_MAPX Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type = TILE_TYPE_DOOR Then
            X = GetPlayerX(Index) + 1
            Y = GetPlayerY(Index)
            
            If TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "door" & END_CHAR)
            End If
        End If
    End If
    If GetPlayerX(Index) - 1 >= 0 Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type = TILE_TYPE_DOOR Then
            X = GetPlayerX(Index) - 1
            Y = GetPlayerY(Index)
            
            If TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "door" & END_CHAR)
            End If
        End If
    End If
    If GetPlayerY(Index) - 1 >= 0 Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type = TILE_TYPE_DOOR Then
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) - 1
            
            If TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "door" & END_CHAR)
            End If
        End If
    End If
    If GetPlayerY(Index) + 1 <= MAX_MAPY Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type = TILE_TYPE_DOOR Then
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) + 1
            
            If TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "door" & END_CHAR)
            End If
        End If
    End If
            
    ' Check to see if the tile is a warp tile, and if so warp them
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_WARP Then
        MapNum = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1
        X = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data2
        Y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data3
        Call PlayerWarp(Index, MapNum, X, Y)
        'Call PlayerPet(Index, 0, GetPlayerDir(Index))
        Call PetMove(Index)
        Moved = YES
    End If
    
    ' Check for key trigger open
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_KEYOPEN Then
        X = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1
        Y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data2
        
        If Map(GetPlayerMap(Index)).Tile(X, Y).type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
            TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
            TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                            
            Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
            If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) = vbNullString Then
                Call MapMsg(GetPlayerMap(Index), "La porte a été ouverte par un mécanisme!", White)
            Else
                Call MapMsg(GetPlayerMap(Index), Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), White)
            End If
            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "key" & END_CHAR)
        End If
    End If
        
    ' Check for shop
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_SHOP Then
       If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1 > 0 Then
            If (GetPlayerX(Index) = Player(Index).Char(Player(Index).CharNum).LastX) And (GetPlayerY(Index) <> Player(Index).Char(Player(Index).CharNum).Y) Then
                Call QueteMsg(Index, Shop(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1).JoinSay)
                Call SendTrade(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1)
            End If
        Else
            Call PlayerMsg(Index, "Il n'y a pas de magasin ici.", BrightRed)
        End If
    End If
        
    ' Check if player stepped on sprite changing tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_SPRITE_CHANGE Then
        If GetPlayerSprite(Index) = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1 Then
            Call PlayerMsg(Index, "Tu as déjà ce sprites!", BrightRed)
            Exit Sub
        Else
            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data2 = 0 Then
                Call SendDataTo(Index, "spritechange" & SEP_CHAR & 0 & END_CHAR)
            Else
                If item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data2).type = ITEM_TYPE_CURRENCY Then
                    Call PlayerMsg(Index, "Ce sprite vous coûte " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data3 & " " & Trim$(item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data2).Name) & "!", Yellow)
                    Call TakeItem(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data2, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data3)
                    Call SendInventory(Index)
                Else
                    Call PlayerMsg(Index, "Ce sprite vous coûte un " & Trim$(item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data2).Name) & "!", Yellow)
                    Call TakeItem(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data2, 1)
                    Call SendInventory(Index)
                End If
                
                Call SendDataTo(Index, "spritechange" & SEP_CHAR & 1 & END_CHAR)
            End If
        End If
    End If
    
    ' Check if player stepped on class change
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_CLASS_CHANGE Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data2 > -1 Then
            If GetPlayerClass(Index) <> Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data2 Then
                Call PlayerMsg(Index, "Tu n'as pas la classe requise!", BrightRed)
                Exit Sub
            End If
        End If
        
        If GetPlayerClass(Index) = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1 Then
            Call PlayerMsg(Index, "Tu as déjà cette classe!", BrightRed)
        Else
            If Player(Index).Char(Player(Index).CharNum).Sex = 0 Then
                If GetPlayerSprite(Index) = Classe(GetPlayerClass(Index)).MaleSprite Then
                    Call SetPlayerSprite(Index, Classe(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1).MaleSprite)
                End If
            Else
                If GetPlayerSprite(Index) = Classe(GetPlayerClass(Index)).FemaleSprite Then
                    Call SetPlayerSprite(Index, Classe(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1).FemaleSprite)
                End If
            End If
            
            Call SetPlayerStr(Index, (Player(Index).Char(Player(Index).CharNum).STR - Classe(GetPlayerClass(Index)).STR))
            Call SetPlayerDEF(Index, (Player(Index).Char(Player(Index).CharNum).def - Classe(GetPlayerClass(Index)).def))
            Call SetPlayerMAGI(Index, (Player(Index).Char(Player(Index).CharNum).magi - Classe(GetPlayerClass(Index)).magi))
            Call SetPlayerSPEED(Index, (Player(Index).Char(Player(Index).CharNum).Speed - Classe(GetPlayerClass(Index)).Speed))
            
            Call SetPlayerClass(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1)

            Call SetPlayerStr(Index, (Player(Index).Char(Player(Index).CharNum).STR + Classe(GetPlayerClass(Index)).STR + GetVar("Classes\Class" & GetPlayerClass(Index) & ".ini", "CLASSCHANGE", "AddStr")))
            Call SetPlayerDEF(Index, (Player(Index).Char(Player(Index).CharNum).def + Classe(GetPlayerClass(Index)).def + GetVar("Classes\Class" & GetPlayerClass(Index) & ".ini", "CLASSCHANGE", "AddDef")))
            Call SetPlayerMAGI(Index, (Player(Index).Char(Player(Index).CharNum).magi + Classe(GetPlayerClass(Index)).magi + GetVar("Classes\Class" & GetPlayerClass(Index) & ".ini", "CLASSCHANGE", "AddMagi")))
            Call SetPlayerSPEED(Index, (Player(Index).Char(Player(Index).CharNum).Speed + Classe(GetPlayerClass(Index)).Speed + GetVar("Classes\Class" & GetPlayerClass(Index) & ".ini", "CLASSCHANGE", "AddSpeed")))
            
            Dim ItemNum As Long
            ItemNum = Val(GetVar(App.Path & "\" & "Classes\Class" & GetPlayerClass(Index) & ".ini", "STARTUP", "Weapon"))
            If item(ItemNum).type = ITEM_TYPE_WEAPON Then
                i = FindOpenInvSlot(Index, ItemNum)
                If i > 0 Then
                    Call SetPlayerInvItemNum(Index, i, ItemNum)
                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + 1)
                    Call SetPlayerInvItemDur(Index, i, item(ItemNum).data1)
                    Call SetPlayerWeaponSlot(Index, i)
                End If
            End If
            ItemNum = Val(GetVar(App.Path & "\" & "Classes\Class" & GetPlayerClass(Index) & ".ini", "STARTUP", "Shield"))
            If item(ItemNum).type = ITEM_TYPE_SHIELD Then
                i = FindOpenInvSlot(Index, ItemNum)
                If i > 0 Then
                    Call SetPlayerInvItemNum(Index, i, ItemNum)
                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + 1)
                    Call SetPlayerInvItemDur(Index, i, item(ItemNum).data1)
                    Call SetPlayerShieldSlot(Index, i)
                End If
            End If
            ItemNum = Val(GetVar(App.Path & "\" & "Classes\Class" & GetPlayerClass(Index) & ".ini", "STARTUP", "Armor"))
            If item(ItemNum).type = ITEM_TYPE_ARMOR Then
                i = FindOpenInvSlot(Index, ItemNum)
                If i > 0 Then
                    Call SetPlayerInvItemNum(Index, i, ItemNum)
                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + 1)
                    Call SetPlayerInvItemDur(Index, i, item(ItemNum).data1)
                    Call SetPlayerArmorSlot(Index, i)
                End If
            End If
            ItemNum = Val(GetVar(App.Path & "\" & "Classes\Class" & GetPlayerClass(Index) & ".ini", "STARTUP", "Helmet"))
            If item(ItemNum).type = ITEM_TYPE_HELMET Then
                i = FindOpenInvSlot(Index, ItemNum)
                If i > 0 Then
                    Call SetPlayerInvItemNum(Index, i, ItemNum)
                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + 1)
                    Call SetPlayerInvItemDur(Index, i, item(ItemNum).data1)
                    Call SetPlayerHelmetSlot(Index, i)
                End If
            End If
            If item(ItemNum).type = ITEM_TYPE_PET Then
                i = FindOpenInvSlot(Index, ItemNum)
                If i > 0 Then
                    Call SetPlayerInvItemNum(Index, i, ItemNum)
                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + 1)
                    Call SetPlayerInvItemDur(Index, i, item(ItemNum).data1)
                    Call SetPlayerPetSlot(Index, i)
                End If
            End If
            
            
            Call PlayerMsg(Index, "Ta nouvelle classe est " & Trim$(Classe(GetPlayerClass(Index)).Name) & "!", BrightGreen)
            
            Call SendStats(Index)
            Call SendHP(Index)
            Call SendMP(Index)
            Call SendSP(Index)
            Call SendWornEquipment(Index)
            Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
        End If
    End If
    
    ' Check if player stepped on notice tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_NOTICE Then
        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) <> vbNullString And Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String2) <> vbNullString Then
            Call QueteMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) & vbCrLf & Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String2))
        ElseIf Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String2) <> vbNullString Then
            Call QueteMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String2))
        End If
        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String3 & END_CHAR)
    End If
    
    ' Check if player stepped on sound tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_SOUND Then
        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1 & END_CHAR)
    End If
    
    If Scripting = 1 And Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_SCRIPTED Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & Index & "," & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1
    End If
    
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_CRAFT Then
        If Player(Index).Char(Player(Index).CharNum).metier > 0 Then
            If metier(Player(Index).Char(Player(Index).CharNum).metier).type = METIER_CRAFT Then
                Packet = "CRAFT" & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1 & END_CHAR
                Call SendDataTo(Index, Packet)
            Else
                Call PlayerMsg(Index, "Votre métier n'est pas un métier de craft!", Red)
            End If
        Else
            Call PlayerMsg(Index, "Vous n'avez pas de métier !", Red)
        End If
    End If
    
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_METIER Then
        If Player(Index).Char(Player(Index).CharNum).metier = 0 Then
            Packet = "NEWMETIER" & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1 & END_CHAR
        Else
            If Player(Index).Char(Player(Index).CharNum).metier <> Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1 Then
                Packet = "REMPLACEMETIER" & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1 & END_CHAR
            Else
                Call PlayerMsg(Index, "Vous avez déjà ce métier !", Red)
            End If
        End If
        Call SendDataTo(Index, Packet)
    End If
    
    ' verifier si le joueure marche sur une case bank
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_BANK Then
        Call QueteMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1))
        Packet = "BANK" & END_CHAR
        Call SendDataTo(Index, Packet)
    End If
        
    ' verifier si le joueure marche sur une case de fin de donjon
    If Player(Index).Char(Player(Index).CharNum).QueteEnCour > 0 Then
        If quete(Player(Index).Char(Player(Index).CharNum).QueteEnCour).type = QUETE_TYPE_FINIR And Player(Index).Char(Player(Index).CharNum).QueteStatut(Player(Index).Char(Player(Index).CharNum).QueteEnCour) = 0 Then
            If GetPlayerMap(Index) = quete(Player(Index).Char(Player(Index).CharNum).QueteEnCour).data3 And GetPlayerX(Index) = quete(Player(Index).Char(Player(Index).CharNum).QueteEnCour).data1 And GetPlayerY(Index) = quete(Player(Index).Char(Player(Index).CharNum).QueteEnCour).data2 Then
                Player(Index).Char(Player(Index).CharNum).QueteStatut(Player(Index).Char(Player(Index).CharNum).QueteEnCour) = 1
                Call SendDataTo(Index, "FINQUETE" & END_CHAR)
            End If
        End If
    End If

Exit Sub
er:
On Error Resume Next
If Index < 0 Or Index > MAX_PLAYERS Then Exit Sub
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur pendant le mouvement du joueur : " & GetPlayerName(Index) & ",Compte : " & GetPlayerLogin(Index) & ",Direction : " & Dir & "(" & Movement & "). Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur pendant le mouvement du joueur : " & GetPlayerName(Index), BrightRed, True)
Call PlainMsg(Index, "Erreur du serveur, relancer SVP!(Pour tous problème récurent visiter " & Trim$(GetVar(App.Path & "\Config\.ini", "CONFIG", "WebSite")) & ").", 3)
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
        If n <> TILE_TYPE_NPCAVOID And CLng(Npc(MapNpc(MapNum, MapNpcNum).Num).Vol) <> 0 Then CanNpcMove = True: Exit Function
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
        If (i <> MapNpcNum) And (MapNpc(MapNum, i).Num > 0) And (MapNpc(MapNum, i).X = MapNpc(MapNum, MapNpcNum).X + (TmpX - 1)) And (MapNpc(MapNum, i).Y = MapNpc(MapNum, MapNpcNum).Y + (TmpY - 1)) Then
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
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_DOWN
            MapNpc(MapNum, MapNpcNum).Y = MapNpc(MapNum, MapNpcNum).Y + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_LEFT
            MapNpc(MapNum, MapNpcNum).X = MapNpc(MapNum, MapNpcNum).X - 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_RIGHT
            MapNpc(MapNum, MapNpcNum).X = MapNpc(MapNum, MapNpcNum).X + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
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
            If IsPlaying(i) And GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X + 1 And GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y And CLng(Npc(MapNpc(MapNum, MapNpcNum).Num).Vol) = 0 Then Exit Sub
        Next i
        If Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X + 1, MapNpc(MapNum, MapNpcNum).Y).type = TILE_TYPE_WALKABLE Or Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X + 1, MapNpc(MapNum, MapNpcNum).Y).type = TILE_TYPE_ITEM Or (Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X + 1, MapNpc(MapNum, MapNpcNum).Y).type <> TILE_TYPE_NPCAVOID Or CLng(Npc(MapNpc(MapNum, MapNpcNum).Num).Vol) <> 0) Then
            If Not CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then Exit Sub
            MapNpc(MapNum, MapNpcNum).Dir = DIR_RIGHT
            MapNpc(MapNum, MapNpcNum).X = MapNpc(MapNum, MapNpcNum).X + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, Packet)
            Exit Sub
        End If
    End If
        
    If X < MapNpc(MapNum, MapNpcNum).X Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X - 1 And GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y And CLng(Npc(MapNpc(MapNum, MapNpcNum).Num).Vol) = 0 Then Exit Sub
        Next i
        If Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X - 1, MapNpc(MapNum, MapNpcNum).Y).type = TILE_TYPE_WALKABLE Or Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X - 1, MapNpc(MapNum, MapNpcNum).Y).type = TILE_TYPE_ITEM Or (Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X - 1, MapNpc(MapNum, MapNpcNum).Y).type <> TILE_TYPE_NPCAVOID Or CLng(Npc(MapNpc(MapNum, MapNpcNum).Num).Vol) <> 0) Then
            If Not CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then Exit Sub
            MapNpc(MapNum, MapNpcNum).Dir = DIR_LEFT
            MapNpc(MapNum, MapNpcNum).X = MapNpc(MapNum, MapNpcNum).X - 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, Packet)
            Exit Sub
        End If
    End If
    
    If MapNpc(MapNum, MapNpcNum).Y < Y Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X And GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y + 1 And CLng(Npc(MapNpc(MapNum, MapNpcNum).Num).Vol) = 0 Then Exit Sub
        Next i
        If Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y + 1).type = TILE_TYPE_WALKABLE Or Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y + 1).type = TILE_TYPE_ITEM Or (Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y + 1).type <> TILE_TYPE_NPCAVOID Or CLng(Npc(MapNpc(MapNum, MapNpcNum).Num).Vol) <> 0) Then
            If Not CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then Exit Sub
            MapNpc(MapNum, MapNpcNum).Dir = DIR_DOWN
            MapNpc(MapNum, MapNpcNum).Y = MapNpc(MapNum, MapNpcNum).Y + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, Packet)
            Exit Sub
        End If
    End If
    
    If MapNpc(MapNum, MapNpcNum).Y > Y Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X And GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y - 1 And CLng(Npc(MapNpc(MapNum, MapNpcNum).Num).Vol) = 0 Then Exit Sub
        Next i
        If Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y - 1).type = TILE_TYPE_WALKABLE Or Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y - 1).type = TILE_TYPE_ITEM Or (Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y - 1).type <> TILE_TYPE_NPCAVOID Or CLng(Npc(MapNpc(MapNum, MapNpcNum).Num).Vol) <> 0) Then
            If Not CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then Exit Sub
            MapNpc(MapNum, MapNpcNum).Dir = DIR_UP
            MapNpc(MapNum, MapNpcNum).Y = MapNpc(MapNum, MapNpcNum).Y - 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, Packet)
            Exit Sub
        End If
     End If
Else

    If MapNpc(MapNum, MapNpcNum).Y < Y Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X And GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y + 1 And CLng(Npc(MapNpc(MapNum, MapNpcNum).Num).Vol) = 0 Then Exit Sub
        Next i
        If Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y + 1).type = TILE_TYPE_WALKABLE Or Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y + 1).type = TILE_TYPE_ITEM Or (Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y + 1).type <> TILE_TYPE_NPCAVOID Or CLng(Npc(MapNpc(MapNum, MapNpcNum).Num).Vol) <> 0) Then
            If Not CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then Exit Sub
            MapNpc(MapNum, MapNpcNum).Dir = DIR_DOWN
            MapNpc(MapNum, MapNpcNum).Y = MapNpc(MapNum, MapNpcNum).Y + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, Packet)
            Exit Sub
        End If
    End If
    
    If MapNpc(MapNum, MapNpcNum).Y > Y Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X And GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y - 1 And CLng(Npc(MapNpc(MapNum, MapNpcNum).Num).Vol) = 0 Then Exit Sub
        Next i
        If Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y - 1).type = TILE_TYPE_WALKABLE Or Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y - 1).type = TILE_TYPE_ITEM Or (Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y - 1).type <> TILE_TYPE_NPCAVOID Or CLng(Npc(MapNpc(MapNum, MapNpcNum).Num).Vol) <> 0) Then
            If Not CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then Exit Sub
            MapNpc(MapNum, MapNpcNum).Dir = DIR_UP
            MapNpc(MapNum, MapNpcNum).Y = MapNpc(MapNum, MapNpcNum).Y - 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, Packet)
            Exit Sub
        End If
    End If
    
    If X > MapNpc(MapNum, MapNpcNum).X Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X + 1 And GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y And CLng(Npc(MapNpc(MapNum, MapNpcNum).Num).Vol) = 0 Then Exit Sub
        Next i
        If Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X + 1, MapNpc(MapNum, MapNpcNum).Y).type = TILE_TYPE_WALKABLE Or Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X + 1, MapNpc(MapNum, MapNpcNum).Y).type = TILE_TYPE_ITEM Or (Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X + 1, MapNpc(MapNum, MapNpcNum).Y).type <> TILE_TYPE_NPCAVOID Or CLng(Npc(MapNpc(MapNum, MapNpcNum).Num).Vol) <> 0) Then
            If Not CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then Exit Sub
            MapNpc(MapNum, MapNpcNum).Dir = DIR_RIGHT
            MapNpc(MapNum, MapNpcNum).X = MapNpc(MapNum, MapNpcNum).X + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, Packet)
            Exit Sub
        End If
    End If
    
    If X < MapNpc(MapNum, MapNpcNum).X Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X - 1 And GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y And CLng(Npc(MapNpc(MapNum, MapNpcNum).Num).Vol) = 0 Then Exit Sub
        Next i
        If Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X - 1, MapNpc(MapNum, MapNpcNum).Y).type = TILE_TYPE_WALKABLE Or Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X - 1, MapNpc(MapNum, MapNpcNum).Y).type = TILE_TYPE_ITEM Or (Map(MapNum).Tile(MapNpc(MapNum, MapNpcNum).X - 1, MapNpc(MapNum, MapNpcNum).Y).type <> TILE_TYPE_NPCAVOID Or CLng(Npc(MapNpc(MapNum, MapNpcNum).Num).Vol) <> 0) Then
            If Not CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then Exit Sub
            MapNpc(MapNum, MapNpcNum).Dir = DIR_LEFT
            MapNpc(MapNum, MapNpcNum).X = MapNpc(MapNum, MapNpcNum).X - 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
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
    Packet = "NPCDIR" & SEP_CHAR & MapNpcNum & SEP_CHAR & Dir & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub JoinGame(ByVal Index As Long)
Dim MOTD As String
Dim f As Long
    
    On Error GoTo er:
    
    ' Set the flag so we know the person is in the game
    Player(Index).InGame = True
    
    ' Send an ok to client to start receiving in game data
    Call SendDataTo(Index, "LOGINOK" & SEP_CHAR & Index & END_CHAR)
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(Index)
    Call SendClasses(Index)
    Call SendItems(Index)
    Call SendPets(Index)
    Call SendMetiers(Index)
    Call SendRecettes(Index)
    Call SendEmoticons(Index)
    Call SendArrows(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendQuetes(Index)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
    Call SendStats(Index)
    Call SendWeatherTo(Index)
    Call SendTimeTo(Index)
    Call SendOnlineList
    Call SendDataTo(Index, "PICVALUE" & SEP_CHAR & PIC_PL & SEP_CHAR & PIC_NPC1 & SEP_CHAR & PIC_NPC2 & SEP_CHAR & AccModo & SEP_CHAR & AccMapeur & SEP_CHAR & AccDevelopeur & SEP_CHAR & AccAdmin & END_CHAR)
    Call LoadPlayerQuete(Index)
    Call SendPlayerQuete(Index)
    Call SendPlayerSpells(Index)
    Call SendPlayerMetier(Index)
    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    Call SendPlayerData(Index)
    
    If Scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "JoinGame " & Index
    Else
        MOTD = GetVar("motd.ini", "MOTD", "Msg")
        
        ' Send a global message that he/she joined
        If GetPlayerAccess(Index) <= ADMIN_MONITER Then
            Call GlobalMsg(GetPlayerName(Index) & " a rejoin " & GAME_NAME & "!", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(Index) & " a rejoin " & GAME_NAME & "!", AdminColor)
            Call IBMsg("L'Admin/Modo : " & GetPlayerName(Index) & " a rejoin " & GAME_NAME & "!", IBCAdmin, True)
        End If
    
        ' Send them welcome
        Call PlayerMsg(Index, "Bienvenue sur " & GAME_NAME & "!", 15)
        
        ' Send motd
        If Trim$(MOTD) <> vbNullString Then Call PlayerMsg(Index, "MOTD: " & MOTD, 11)
    End If
    
    ' Send whos online
    Call SendWhosOnline(Index)
    Call ShowPLR(Index)
    'PAPERDOLL
    Dim i As Long
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(Index) Then
            If Index <> i Then Call SendInventory(i)
            Call SendWornEquipment(i)
            'Call PlayerPet(i, 0, GetPlayerDir(i))
            Call PetMove(i)
        End If
    Next i
    'FIN PAPERDOLL

    ' Send the flag so they know they can start doing stuff
    Call SendDataTo(Index, "INGAME" & END_CHAR)
Exit Sub
er:
On Error Resume Next
If Index < 0 Or Index > MAX_PLAYERS Then Exit Sub
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur de connexion au jeu, joueur : " & GetPlayerName(Index) & ",Compte : " & GetPlayerLogin(Index) & ". Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur de connexion au jeu, joueur : " & GetPlayerName(Index), BrightRed, True)
Call PlainMsg(Index, "Erreur du serveur, relancer SVP!(Pour tous problème récurent visiter " & Trim$(GetVar(App.Path & "\Config\.ini", "CONFIG", "WebSite")) & ").", 3)
End Sub

Sub LeftGame(ByVal Index As Long)
Dim n As Long

If Len(Trim$(Player(Index).Login)) <= 1 Then Exit Sub

    On Error GoTo er:
        
    If bouclier(Index) Then bouclier(Index) = False: BouclierT(Index) = 0
    If Para(Index) Then Call ContrOnOff(Index): Para(Index) = False: ParaT(Index) = 0
    If Point(Index) > 0 And Point(Index) < MAX_SPELLS Then
        If Spell(Point(Index)).type = SPELL_TYPE_AMELIO Then
            Player(Index).Char(Player(Index).CharNum).def = Player(Index).Char(Player(Index).CharNum).def - Val(Spell(Point(Index)).data3)
            Player(Index).Char(Player(Index).CharNum).magi = Player(Index).Char(Player(Index).CharNum).magi - Val(Spell(Point(Index)).data3)
            Player(Index).Char(Player(Index).CharNum).STR = Player(Index).Char(Player(Index).CharNum).STR - Val(Spell(Point(Index)).data3)
            Player(Index).Char(Player(Index).CharNum).Speed = Player(Index).Char(Player(Index).CharNum).Speed - Val(Spell(Point(Index)).data3)
            Call SendStats(Index)
            Point(Index) = 0
            PointT(Index) = 0
        ElseIf Spell(Point(Index)).type = SPELL_TYPE_DECONC And GetTickCount >= PointT(Index) Then
            Player(Index).Char(Player(Index).CharNum).def = Player(Index).Char(Player(Index).CharNum).def + Val(Spell(Point(Index)).data3)
            Player(Index).Char(Player(Index).CharNum).magi = Player(Index).Char(Player(Index).CharNum).magi + Val(Spell(Point(Index)).data3)
            Player(Index).Char(Player(Index).CharNum).STR = Player(Index).Char(Player(Index).CharNum).STR + Val(Spell(Point(Index)).data3)
            Player(Index).Char(Player(Index).CharNum).Speed = Player(Index).Char(Player(Index).CharNum).Speed + Val(Spell(Point(Index)).data3)
            Call SendStats(Index)
            Point(Index) = 0
            PointT(Index) = 0
        End If
    End If
    
    If Player(Index).InGame Then
        
        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(Index)) = 0 Then PlayersOnMap(GetPlayerMap(Index)) = NO
        
        Player(Index).InGame = False
        
        ' Check if the player was in a party, and if so cancel it out so the other player doesn't continue to get half exp
        Party.RemoveMember Player(Index).InParty, Player(Index).PartyPlayer
        
        ' Check for boot map
        If Map(GetPlayerMap(Index)).BootMap > 0 Then
            Call SetPlayerX(Index, Map(GetPlayerMap(Index)).BootX)
            Call SetPlayerY(Index, Map(GetPlayerMap(Index)).BootY)
            Call SetPlayerMap(Index, Map(GetPlayerMap(Index)).BootMap)
        End If
            
        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "LeftGame " & Index
        Else
            ' Send a global message that he/she left
            If GetPlayerAccess(Index) <= 1 Then
                Call GlobalMsg(GetPlayerName(Index) & " a quitté " & GAME_NAME & "!", 7)
            Else
                Call GlobalMsg(GetPlayerName(Index) & " a quitté " & GAME_NAME & "!", 15)
            End If
        End If
        'If quete(Player(Index).Char(Player(Index).CharNum).QueteEnCour).temps > 0 Then
        '    Player(Index).Char(Player(Index).CharNum).QueteStatut(Player(Index).Char(Player(Index).CharNum).QueteEnCour) = 0
        'Else
        '    Player(Index).Char(Player(Index).CharNum).QueteEnCour = 0
        'End If
        
        Call SavePlayer(Index)
        
        Call TextAdd(frmServer.txtText(0), GetPlayerName(Index) & " est déconnecté de " & GAME_NAME & ".", True)
        Call SendLeftGame(Index)
        'Call RemovePLR
        For n = 1 To MAX_PLAYERS
           Call ShowPLR(n)
        Next n
    End If
    Call ClearPlayer(Index)
    Call SendOnlineList
Exit Sub
er:
On Error Resume Next
If Index < 0 Or Index > MAX_PLAYERS Then Exit Sub
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur de déconnexion au jeu, joueur : " & GetPlayerName(Index) & ",Compte : " & GetPlayerLogin(Index) & ". Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur de déconnexion au jeu, joueur : " & GetPlayerName(Index), BrightRed, True)
Call PlainMsg(Index, "Erreur du serveur, relancer SVP!(Pour tous problème récurent visiter " & Trim$(GetVar(App.Path & "\Config\.ini", "CONFIG", "WebSite")) & ").", 3)
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

Function GetNpcHPRegen(ByVal npcnum As Long)
Dim i As Long

    'Prevent subscript out of range
    If npcnum <= 0 Or npcnum > MAX_NPCS Then GetNpcHPRegen = 0: Exit Function
    
    i = (Npc(npcnum).def \ 3)
    If i < 1 Then i = 1
    
    GetNpcHPRegen = i
End Function

Sub CheckPlayerLevelUp(ByVal Index As Long)
Dim i As Long
Dim d As Long
Dim c As Long
    c = 0
    
    On Error GoTo er:
    
    If GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then
        If GetPlayerLevel(Index) < MAX_LEVEL Then
            If Scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerLevelUp " & Index
            Else
                Do Until GetPlayerExp(Index) < GetPlayerNextLevel(Index)
                    DoEvents
                    If GetPlayerLevel(Index) < MAX_LEVEL Then
                        If GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then
                            d = GetPlayerExp(Index) - GetPlayerNextLevel(Index)
                            Call SetPlayerLevel(Index, GetPlayerLevel(Index) + 1)
                            i = (GetPlayerSPEED(Index) \ 10)
                            If i < 1 Then i = 1
                            If i > 3 Then i = 3
                                
                            Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + i)
                            Call SetPlayerExp(Index, d)
                            c = c + 1
                        End If
                    End If
                Loop
                If c > 1 Then
                    Call GlobalMsg(GetPlayerName(Index) & " a gagné " & c & " niveaux!", 6)
                Else
                    Call GlobalMsg(GetPlayerName(Index) & " a gagné un niveau!", 6)
                End If
                Call BattleMsg(Index, "Vous avez " & GetPlayerPOINTS(Index) & " points de stats.", 9, 0)
            End If
            Call SendDataToMap(GetPlayerMap(Index), "levelup" & SEP_CHAR & Index & END_CHAR)
        End If
        
        If GetPlayerLevel(Index) = MAX_LEVEL Then
            Call SetPlayerExp(Index, experience(MAX_LEVEL))
        End If
    End If
    
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
    Call SendStats(Index)
Exit Sub
er:
On Error Resume Next
If Index < 0 Or Index > MAX_PLAYERS Then Exit Sub
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur lors de la vérification du niveau du joueur : " & GetPlayerName(Index) & ",Compte : " & GetPlayerLogin(Index) & ". Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur lors de la vérification du niveau du joueur : " & GetPlayerName(Index), BrightRed, True)
End Sub

Sub CastSpell(ByVal Index As Long, ByVal SpellSlot As Long)
Dim SpellNum As Long, i As Long, n As Long, Damage As Long
Dim Casted As Boolean

    Casted = False
    
    On Error GoTo er:
    
    ' Prevent subscript out of range
    If SpellSlot <= 0 Or SpellSlot > MAX_PLAYER_SPELLS Then Exit Sub
    If Not IsPlaying(Index) Then Exit Sub
        
    SpellNum = GetPlayerSpell(Index, SpellSlot)
    
    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then
        Call BattleMsg(Index, "Vous n'avez pas ce sort!", BrightRed, 0)
        Exit Sub
    End If
    
    i = GetSpellReqLevel(Index, SpellNum)

    ' Check if they have enough MP
    If GetPlayerMP(Index) < Spell(SpellNum).MPCost Then
        Call BattleMsg(Index, "Pas assez de mana!", BrightRed, 0)
        Exit Sub
    End If
        
    ' Make sure they are the right level
    If i > GetPlayerLevel(Index) Then
        Call BattleMsg(Index, "Vous devez étre niveau " & i & " pour lancer ce sort.", BrightRed, 0)
        Exit Sub
    End If
    
    ' Check if timer is ok
    If GetTickCount < Player(Index).AttackTimer + 1000 Then Exit Sub
    
    ' Check if the spell is a give item and do that instead of a stat modification
    If Spell(SpellNum).type = 15 Then 'SPELL_TYPE_GIVEITEM Then
        n = FindOpenInvSlot(Index, Spell(SpellNum).data1)
        
        If n > 0 Then
            Call GiveItem(Index, Spell(SpellNum).data1, Spell(SpellNum).data2)
            'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & trim$(Spell(SpellNum).Name) & ".", BrightBlue)
            
            ' Take away the mana points
            Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
            Call SendMP(Index)
            Casted = True
        Else
            Call PlayerMsg(Index, "Votre inventaire est plein!", BrightRed)
        End If
        
        Exit Sub
    End If
        
Dim X As Long, Y As Long

If Spell(SpellNum).AE = 1 Then
    For Y = GetPlayerY(Index) - Spell(SpellNum).Range To GetPlayerY(Index) + Spell(SpellNum).Range
        For X = GetPlayerX(Index) - Spell(SpellNum).Range To GetPlayerX(Index) + Spell(SpellNum).Range
            n = -1
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) = True Then
                    If GetPlayerMap(Index) = GetPlayerMap(i) Then
                        If GetPlayerX(i) = X And GetPlayerY(i) = Y Then
                            If i = Index Then
                                If Spell(SpellNum).type = SPELL_TYPE_ADDHP Or Spell(SpellNum).type = SPELL_TYPE_ADDMP Or Spell(SpellNum).type = SPELL_TYPE_ADDSP Then
                                    Player(Index).Target = i
                                    Player(Index).TargetType = TARGET_TYPE_PLAYER
                                    n = Player(Index).Target
                                End If
                            Else
                                Player(Index).Target = i
                                Player(Index).TargetType = TARGET_TYPE_PLAYER
                                n = Player(Index).Target
                            End If
                        End If
                    End If
                End If
            Next i
            
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(GetPlayerMap(Index), i).Num > 0 Then
                    If MapNpc(GetPlayerMap(Index), i).X = X And MapNpc(GetPlayerMap(Index), i).Y = Y Then
                        If Npc(MapNpc(GetPlayerMap(Index), i).Num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(Index), i).Num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(MapNpc(GetPlayerMap(Index), i).Num).Behavior <> NPC_BEHAVIOR_QUETEUR Then
                            Player(Index).Target = i
                            Player(Index).TargetType = TARGET_TYPE_NPC
                            n = Player(Index).Target
                        End If
                    End If
                End If
            Next i
                
        Casted = False
        If n > 0 Then
            If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
                If IsPlaying(n) Then
                    If bouclier(n) = True Then
                        Call BattleMsg(Index, "Le sort ne peut pas toucher le joueur car il a un bouclier!", BrightRed, 0)
                        Exit Sub
                    End If
'                    If n <> Index Then
                        Player(Index).TargetType = TARGET_TYPE_PLAYER
                        If GetPlayerHP(n) > 0 And GetPlayerMap(Index) = GetPlayerMap(n) And GetPlayerLevel(Index) >= NOOB_LEVEL And GetPlayerLevel(n) >= NOOB_LEVEL And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(n) <= 0 Then
                            Select Case Spell(SpellNum).type
                                
                                Case SPELL_TYPE_SUBHP
                                    Damage = ((GetPlayerMAGI(Index) \ 4) + Spell(SpellNum).data1) - GetPlayerProtection(n)
                                    If Damage > 0 Then Call AttackPlayer(Index, n, Damage) Else Call BattleMsg(Index, "Votre sort n'est pas assez puissant pour blesser " & GetPlayerName(n) & "!", BrightRed, 0)
                                                            
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
                                    If Point(n) > 0 Or PointT(n) > 0 Then Call PlayerMsg(Index, "Le joueur est déjà la cible d'un sort d'amélioration.", BrightRed)
                                    Player(n).Char(Player(n).CharNum).def = Player(n).Char(Player(n).CharNum).def + Val(Spell(SpellNum).data3)
                                    Player(n).Char(Player(n).CharNum).magi = Player(n).Char(Player(n).CharNum).magi + Val(Spell(SpellNum).data3)
                                    Player(n).Char(Player(n).CharNum).STR = Player(n).Char(Player(n).CharNum).STR + Val(Spell(SpellNum).data3)
                                    Player(n).Char(Player(n).CharNum).Speed = Player(n).Char(Player(n).CharNum).Speed + Val(Spell(SpellNum).data3)
                                    Call SendStats(n)
                                    Point(n) = SpellNum
                                    PointT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                                    
                                Case SPELL_TYPE_DECONC
                                    If Point(n) > 0 Or PointT(n) > 0 Then Call PlayerMsg(Index, "Le joueur est déjà la cible d'un sort de déconcentration.", BrightRed)
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
                            If GetPlayerMap(Index) = GetPlayerMap(n) And Spell(SpellNum).type >= SPELL_TYPE_ADDHP And Spell(SpellNum).type <= SPELL_TYPE_ADDSP Or Spell(SpellNum).type = SPELL_TYPE_SCRIPT Then
                                Select Case Spell(SpellNum).type
                                
                                    Case SPELL_TYPE_ADDHP
                                        Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).data1)
                                        Call SendDataTo(n, "BLITNPCDMG" & SEP_CHAR & Spell(SpellNum).data1 & SEP_CHAR & 0 & SEP_CHAR & 1 & END_CHAR)
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
                                        If Point(n) > 0 Or PointT(n) > 0 Then Call PlayerMsg(Index, "Le joueur est déjà la cible d'un sort d'amélioration.", BrightRed)
                                        Player(n).Char(Player(n).CharNum).def = Player(n).Char(Player(n).CharNum).def + Val(Spell(SpellNum).data3)
                                        Player(n).Char(Player(n).CharNum).magi = Player(n).Char(Player(n).CharNum).magi + Val(Spell(SpellNum).data3)
                                        Player(n).Char(Player(n).CharNum).STR = Player(n).Char(Player(n).CharNum).STR + Val(Spell(SpellNum).data3)
                                        Player(n).Char(Player(n).CharNum).Speed = Player(n).Char(Player(n).CharNum).Speed + Val(Spell(SpellNum).data3)
                                        Call SendStats(n)
                                        Point(n) = SpellNum
                                        PointT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                                    
                                    Case SPELL_TYPE_DECONC
                                        If Point(n) > 0 Or PointT(n) > 0 Then Call PlayerMsg(Index, "Le joueur est déjà la cible d'un sort de déconcentration.", BrightRed)
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
                                Call PlayerMsg(Index, "Vous n'avez pas pu envoyer le sort!(la cible n'est pas sur la même carte que vous)", BrightRed)
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
                    Call BattleMsg(Index, "Vous n'avez pas put envoyer le sort!(la cible n'est pas/plus en jeu)", BrightRed, 0)
                End If
            Else
                Player(Index).TargetType = TARGET_TYPE_NPC
                If Npc(MapNpc(GetPlayerMap(Index), n).Num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(Index), n).Num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(MapNpc(GetPlayerMap(Index), n).Num).Behavior <> NPC_BEHAVIOR_QUETEUR And Npc(MapNpc(GetPlayerMap(Index), n).Num).Behavior <> NPC_BEHAVIOR_SCRIPT Then
                    If Spell(SpellNum).type >= SPELL_TYPE_SUBHP And Spell(SpellNum).type <= SPELL_TYPE_SUBSP Or Spell(SpellNum).type = SPELL_TYPE_PARALY Then
                        Select Case Spell(SpellNum).type
                            
                            Case SPELL_TYPE_SUBHP
                                Damage = ((GetPlayerMAGI(Index) \ 4) + Spell(SpellNum).data1) - (Npc(MapNpc(GetPlayerMap(Index), n).Num).def \ 2)
                                If Damage > 0 Then Call AttackNpc(Index, n, Damage) Else Call BattleMsg(Index, "Votre sort n'est pas assez puissant pour blesser " & Trim$(Npc(MapNpc(GetPlayerMap(Index), n).Num).Name) & "!", BrightRed, 0)
                            Case SPELL_TYPE_SUBMP
                                MapNpc(GetPlayerMap(Index), n).MP = MapNpc(GetPlayerMap(Index), n).MP - Spell(SpellNum).data1

                            Case SPELL_TYPE_SUBSP
                                MapNpc(GetPlayerMap(Index), n).SP = MapNpc(GetPlayerMap(Index), n).SP - Spell(SpellNum).data1
                                
                            Case SPELL_TYPE_PARALY
                                Call PNJOnOff(n, GetPlayerMap(Index))
                                ParaN(n, GetPlayerMap(Index)) = True
                                ParaNT(n, GetPlayerMap(Index)) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                            
                        End Select
                        
                        Casted = True
                    Else
                        Casted = False
                    End If
                Else
                    Call BattleMsg(Index, "Vous ne lancez pas le sort!!(PNJ amis)", BrightRed, 0)
                End If
            End If
        End If
        If Casted Then
            Call SendDataToMap(GetPlayerMap(Index), "spellanim" & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & Player(Index).Target & END_CHAR)
            'Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "magic" & SEP_CHAR & Spell(SpellNum).Sound & END_CHAR)
        End If
        Next X
    Next Y
    
    Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
    Call SendMP(Index)
Else
    n = Player(Index).Target
    If n = -1 Then Call PlayerMsg(Index, "Vous n'avez pas pu envoyer le sort!(aucune cible)", BrightRed): Exit Sub
    If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
        If IsPlaying(n) Then
            If bouclier(n) = True Then
                Call BattleMsg(Index, "Le sort ne peut pas toucher le joueur car il a un bouclier!", BrightRed, 0)
                Exit Sub
            End If

            If GetPlayerName(n) <> GetPlayerName(Index) Then
                If CInt(Sqr((GetPlayerX(Index) - GetPlayerX(n)) ^ 2 + ((GetPlayerY(Index) - GetPlayerY(n)) ^ 2))) > Spell(SpellNum).Range Then
                    Call BattleMsg(Index, "Vous êtes trop loin pour toucher la cible.", BrightRed, 0)
                    Exit Sub
                End If
            End If
            Player(Index).TargetType = TARGET_TYPE_PLAYER
            
            If GetPlayerHP(n) > 0 And GetPlayerMap(Index) = GetPlayerMap(n) And GetPlayerLevel(Index) >= NOOB_LEVEL And GetPlayerLevel(n) >= NOOB_LEVEL And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(n) <= 0 Then
                                        
                Select Case Spell(SpellNum).type

                    Case SPELL_TYPE_SUBHP
                        Damage = ((GetPlayerMAGI(Index) \ 4) + Spell(SpellNum).data1) - GetPlayerProtection(n)
                        If Damage > 0 Then Call AttackPlayer(Index, n, Damage) Else Call BattleMsg(Index, "Votre sort n'est pas assez puissant pour blesser " & GetPlayerName(n) & "!", BrightRed, 0)
                        
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
                        If Point(n) > 0 Or PointT(n) > 0 Then Call PlayerMsg(Index, "Le joueur est déjà la cible d'un sort d'amélioration.", BrightRed)
                        Player(n).Char(Player(n).CharNum).def = Player(n).Char(Player(n).CharNum).def + Val(Spell(SpellNum).data3)
                        Player(n).Char(Player(n).CharNum).magi = Player(n).Char(Player(n).CharNum).magi + Val(Spell(SpellNum).data3)
                        Player(n).Char(Player(n).CharNum).STR = Player(n).Char(Player(n).CharNum).STR + Val(Spell(SpellNum).data3)
                        Player(n).Char(Player(n).CharNum).Speed = Player(n).Char(Player(n).CharNum).Speed + Val(Spell(SpellNum).data3)
                        Call SendStats(n)
                        Point(n) = SpellNum
                        PointT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                                    
                    Case SPELL_TYPE_DECONC
                        If Point(n) > 0 Or PointT(n) > 0 Then Call PlayerMsg(Index, "Le joueur est déjà la cible d'un sort de déconcentration.", BrightRed)
                        Player(n).Char(Player(n).CharNum).def = Player(n).Char(Player(n).CharNum).def - Val(Spell(SpellNum).data3)
                        Player(n).Char(Player(n).CharNum).magi = Player(n).Char(Player(n).CharNum).magi - Val(Spell(SpellNum).data3)
                        Player(n).Char(Player(n).CharNum).STR = Player(n).Char(Player(n).CharNum).STR - Val(Spell(SpellNum).data3)
                        Player(n).Char(Player(n).CharNum).Speed = Player(n).Char(Player(n).CharNum).Speed - Val(Spell(SpellNum).data3)
                        Call SendStats(n)
                        Point(n) = SpellNum
                        PointT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                    
                End Select
            
                ' Take away the mana points
                Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
                Call SendMP(Index)
                Casted = True
            Else
                If GetPlayerMap(Index) = GetPlayerMap(n) And Spell(SpellNum).type >= SPELL_TYPE_ADDHP And Spell(SpellNum).type <= SPELL_TYPE_ADDSP Or Spell(SpellNum).type >= SPELL_TYPE_SCRIPT Then
                    Select Case Spell(SpellNum).type
                    
                        Case SPELL_TYPE_ADDHP
                            Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).data1)
                            Call SendDataTo(n, "BLITNPCDMG" & SEP_CHAR & Spell(SpellNum).data1 & SEP_CHAR & 0 & SEP_CHAR & 1 & END_CHAR)
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
                            If Point(n) > 0 Or PointT(n) > 0 Then Call PlayerMsg(Index, "Le joueur est déjà la cible d'un sort d'amélioration.", BrightRed)
                            Player(n).Char(Player(n).CharNum).def = Player(n).Char(Player(n).CharNum).def + Val(Spell(SpellNum).data3)
                            Player(n).Char(Player(n).CharNum).magi = Player(n).Char(Player(n).CharNum).magi + Val(Spell(SpellNum).data3)
                            Player(n).Char(Player(n).CharNum).STR = Player(n).Char(Player(n).CharNum).STR + Val(Spell(SpellNum).data3)
                            Player(n).Char(Player(n).CharNum).Speed = Player(n).Char(Player(n).CharNum).Speed + Val(Spell(SpellNum).data3)
                            Call SendStats(n)
                            Point(n) = SpellNum
                            PointT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                                    
                        Case SPELL_TYPE_DECONC
                            If Point(n) > 0 Or PointT(n) > 0 Then Call PlayerMsg(Index, "Le joueur est déjà la cible d'un sort de déconcentration", BrightRed)
                            Player(n).Char(Player(n).CharNum).def = Player(n).Char(Player(n).CharNum).def - Val(Spell(SpellNum).data3)
                            Player(n).Char(Player(n).CharNum).magi = Player(n).Char(Player(n).CharNum).magi - Val(Spell(SpellNum).data3)
                            Player(n).Char(Player(n).CharNum).STR = Player(n).Char(Player(n).CharNum).STR - Val(Spell(SpellNum).data3)
                            Player(n).Char(Player(n).CharNum).Speed = Player(n).Char(Player(n).CharNum).Speed - Val(Spell(SpellNum).data3)
                            Call SendStats(n)
                            Point(n) = SpellNum
                            PointT(n) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                                    
                    End Select
                    
                    ' Take away the mana points
                    Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
                    Call SendMP(Index)
                    Casted = True
                Else
                    Call BattleMsg(Index, "Vous n'avez pas put envoyer le sort!", BrightRed, 0)
                End If
            End If
        Else
            Call PlayerMsg(Index, "Vous n'avez pas put envoyer le sort!(cible hors ligne)", BrightRed)
        End If
    Else
        If CInt(Sqr((GetPlayerX(Index) - MapNpc(GetPlayerMap(Index), n).X) ^ 2 + ((GetPlayerY(Index) - MapNpc(GetPlayerMap(Index), n).Y) ^ 2))) > Spell(SpellNum).Range Then
            Call BattleMsg(Index, "Vous êtes trop loin pour toucher la cible.", BrightRed, 0)
            Exit Sub
        End If
        
        Player(Index).TargetType = TARGET_TYPE_NPC
        
        If Npc(MapNpc(GetPlayerMap(Index), n).Num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(Index), n).Num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(MapNpc(GetPlayerMap(Index), n).Num).Behavior <> NPC_BEHAVIOR_QUETEUR And Npc(MapNpc(GetPlayerMap(Index), n).Num).Behavior <> NPC_BEHAVIOR_SCRIPT Then
            
            Select Case Spell(SpellNum).type
                Case SPELL_TYPE_ADDHP
                    MapNpc(GetPlayerMap(Index), n).HP = MapNpc(GetPlayerMap(Index), n).HP + Spell(SpellNum).data1
                    Call SendDataTo(n, "BLITPLAYERDMG" & SEP_CHAR & Spell(SpellNum).data1 & SEP_CHAR & 1 & END_CHAR)
                
                Case SPELL_TYPE_SUBHP
                    Damage = ((GetPlayerMAGI(Index) \ 4) + Spell(SpellNum).data1) - (Npc(MapNpc(GetPlayerMap(Index), n).Num).def \ 2)
                    If Damage > 0 Then Call AttackNpc(Index, n, Damage) Else Call BattleMsg(Index, "Votre sort n'est pas assez puissant pour blesser " & Trim$(Npc(MapNpc(GetPlayerMap(Index), n).Num).Name) & "!", BrightRed, 0)
                    
                Case SPELL_TYPE_ADDMP
                    MapNpc(GetPlayerMap(Index), n).MP = MapNpc(GetPlayerMap(Index), n).MP + Spell(SpellNum).data1
                
                Case SPELL_TYPE_SUBMP
                    MapNpc(GetPlayerMap(Index), n).MP = MapNpc(GetPlayerMap(Index), n).MP - Spell(SpellNum).data1
            
                Case SPELL_TYPE_ADDSP
                    MapNpc(GetPlayerMap(Index), n).SP = MapNpc(GetPlayerMap(Index), n).SP + Spell(SpellNum).data1
                
                Case SPELL_TYPE_SUBSP
                    MapNpc(GetPlayerMap(Index), n).SP = MapNpc(GetPlayerMap(Index), n).SP - Spell(SpellNum).data1
                
                Case SPELL_TYPE_PARALY
                    Call PNJOnOff(n, GetPlayerMap(Index))
                    ParaN(n, GetPlayerMap(Index)) = True
                    ParaNT(n, GetPlayerMap(Index)) = GetTickCount + Val(Spell(SpellNum).data1 * 1000)
                    
            End Select
        
            ' Take away the mana points
            Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
            Call SendMP(Index)
            Casted = True
        Else
            Call BattleMsg(Index, "Vous n'avez pas pu envoyer le sort!(cible non ennemi)", BrightRed, 0)
        End If
    End If
End If

    If Casted Then
        Player(Index).AttackTimer = GetTickCount
        Player(Index).CastedSpell = YES
        Call SendDataToMap(GetPlayerMap(Index), "spellanim" & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & Player(Index).Target & SEP_CHAR & Player(Index).CastedSpell & END_CHAR)
        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "magic" & SEP_CHAR & Spell(SpellNum).Sound & END_CHAR)
    End If
Exit Sub
er:
Casted = False
On Error Resume Next
If Index < 0 Or Index > MAX_PLAYERS Then Exit Sub
Call PlayerMsg(Index, "Le sort n'a pas put être lancé.", BrightRed)
Call AddLog("le : " & Date & "     à : " & Time & "...Erreur de lancement d'un sort du joueur : " & GetPlayerName(Index) & ",Compte : " & GetPlayerLogin(Index) & ",Slot : " & SpellSlot & ". Détails : Num :" & Err.Number & " Description : " & Err.description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur de lancement d'un sort du joueur : " & GetPlayerName(Index), BrightRed, True)
End Sub

Function GetSpellReqLevel(ByVal Index As Long, ByVal SpellNum As Long)
    GetSpellReqLevel = Spell(SpellNum).LevelReq
End Function

Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
Dim i As Long, n As Long

    CanPlayerCriticalHit = False
        
    If GetPlayerWeaponSlot(Index) > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            i = (GetPlayerStr(Index) \ 2) + (GetPlayerLevel(Index) \ 2)
    
            n = Int(Rnd * 100) + 1
            If n <= i Then CanPlayerCriticalHit = True
        End If
    End If
End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
Dim i As Long, n As Long, ShieldSlot As Long

    CanPlayerBlockHit = False
    
    ShieldSlot = GetPlayerShieldSlot(Index)
    
    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            i = (GetPlayerDEF(Index) \ 2) + (GetPlayerLevel(Index) \ 2)
        
            n = Int(Rnd * 100) + 1
            If n <= i Then CanPlayerBlockHit = True
        End If
    End If
End Function

Function CanPlayerEsquiveHit(ByVal Index As Long) As Boolean
Dim i As Long, n As Long

    CanPlayerEsquiveHit = False
    
        n = Int(Rnd * 2)
        If n = 1 Then
            i = Int(GetPlayerSPEED(Index) * 0.576)
        
            n = Int(Rnd * 100) + 1
            If n <= i Then CanPlayerEsquiveHit = True
        End If
    
End Function

Sub CheckEquippedItems(ByVal Index As Long)
Dim Slot As Long, ItemNum As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    Slot = GetPlayerWeaponSlot(Index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)
        
        If ItemNum > 0 Then
            If item(ItemNum).type <> ITEM_TYPE_WEAPON Then Call SetPlayerWeaponSlot(Index, 0)
        Else
            Call SetPlayerWeaponSlot(Index, 0)
        End If
    End If

    Slot = GetPlayerArmorSlot(Index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)
        
        If ItemNum > 0 Then
            If item(ItemNum).type <> ITEM_TYPE_ARMOR And item(ItemNum).type <> ITEM_TYPE_MONTURE Then
                Call SetPlayerArmorSlot(Index, 0)
            End If
        Else
            Call SetPlayerArmorSlot(Index, 0)
        End If
    End If

    Slot = GetPlayerHelmetSlot(Index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)
        
        If ItemNum > 0 Then
            If item(ItemNum).type <> ITEM_TYPE_HELMET Then Call SetPlayerHelmetSlot(Index, 0)
        Else
            Call SetPlayerHelmetSlot(Index, 0)
        End If
    End If

    Slot = GetPlayerShieldSlot(Index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)
        
        If ItemNum > 0 Then
            If item(ItemNum).type <> ITEM_TYPE_SHIELD Then Call SetPlayerShieldSlot(Index, 0)
        Else
            Call SetPlayerShieldSlot(Index, 0)
        End If
    End If
    
    Slot = GetPlayerPetSlot(Index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)
        
        If ItemNum > 0 Then
            If item(ItemNum).type <> ITEM_TYPE_PET Then Call SetPlayerPetSlot(Index, 0)
        Else
            Call SetPlayerPetSlot(Index, 0)
        End If
    End If
End Sub

Public Sub ShowPLR(ByVal Index As Long)
Dim ls As ListItem
On Error Resume Next

    If frmServer.lvUsers.ListItems.Count > 0 And IsPlaying(Index) = True Then
        frmServer.lvUsers.ListItems.Remove Index
    End If
    Set ls = frmServer.lvUsers.ListItems.add(Index, , Index)
    
    If IsPlaying(Index) = False Then
        ls.SubItems(1) = vbNullString
        ls.SubItems(2) = vbNullString
        ls.SubItems(3) = vbNullString
        ls.SubItems(4) = vbNullString
        ls.SubItems(5) = vbNullString
    Else
        ls.SubItems(1) = GetPlayerLogin(Index)
        ls.SubItems(2) = GetPlayerName(Index)
        ls.SubItems(3) = GetPlayerLevel(Index)
        ls.SubItems(4) = GetPlayerSprite(Index)
        ls.SubItems(5) = GetPlayerAccess(Index)
    End If
End Sub

Public Sub RemovePLR()
    frmServer.lvUsers.ListItems.Clear
End Sub

Function CanAttackNpcWithArrow(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
Dim MapNum As Long, npcnum As Long
Dim AttackSpeed As Long

On Error GoTo er:
If CLng(Npc(MapNpc(GetPlayerMap(Attacker), MapNpcNum).Num).Vol) <> 0 Then CanAttackNpcWithArrow = False: Exit Function
   If GetPlayerWeaponSlot(Attacker) > 0 Then
       AttackSpeed = item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
   Else
       AttackSpeed = 1000
   End If

CanAttackNpcWithArrow = False

' Check For subscript out of range
If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then Exit Function

' Check For subscript out of range
If MapNpc(GetPlayerMap(Attacker), MapNpcNum).Num <= 0 Then Exit Function

MapNum = GetPlayerMap(Attacker)
npcnum = MapNpc(MapNum, MapNpcNum).Num

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

Function AvMonture(ByVal Index As Long) As Boolean
    If Not IsPlaying(Index) Then Exit Function
    
    If GetPlayerArmorSlot(Index) > 0 Then If item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).type = ITEM_TYPE_MONTURE Then AvMonture = True Else AvMonture = False
End Function

Sub Debloque(ByVal Index As Long)
Dim Packet As String

If Not IsPlaying(Index) Then Exit Sub

On Error Resume Next

If GetPlayerX(Index) = MAX_MAPX / 2 And GetPlayerY(Index) = MAX_MAPY / 2 Then
    If GetPlayerX(Index) + 1 < MAX_MAPX Then
        Call SetPlayerX(Index, GetPlayerX(Index) + 1)
    ElseIf GetPlayerX(Index) - 1 > 0 Then
        Call SetPlayerX(Index, GetPlayerX(Index) - 1)
    ElseIf GetPlayerY(Index) + 1 < MAX_MAPY Then
        Call SetPlayerY(Index, GetPlayerY(Index) + 1)
    ElseIf GetPlayerY(Index) - 1 > 0 Then
        Call SetPlayerY(Index, GetPlayerY(Index) - 1)
    End If
Else
    Call SetPlayerX(Index, MAX_MAPX / 2)
    Call SetPlayerY(Index, MAX_MAPY / 2)
End If

Packet = "PLAYERXY" & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & END_CHAR
Call SendDataToMap(GetPlayerMap(Index), Packet)

End Sub

Function ACoter(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
On Error Resume Next

ACoter = False
If Index < 1 Or Index > MAX_PLAYERS Or MapNpcNum < 1 Or MapNpcNum > 15 Then Exit Function

If GetPlayerX(Index) - 1 = MapNpc(GetPlayerMap(Index), MapNpcNum).X And GetPlayerY(Index) = MapNpc(GetPlayerMap(Index), MapNpcNum).Y Then ACoter = True: Exit Function
If GetPlayerX(Index) = MapNpc(GetPlayerMap(Index), MapNpcNum).X And GetPlayerY(Index) - 1 = MapNpc(GetPlayerMap(Index), MapNpcNum).Y Then ACoter = True: Exit Function
If GetPlayerX(Index) = MapNpc(GetPlayerMap(Index), MapNpcNum).X And GetPlayerY(Index) + 1 = MapNpc(GetPlayerMap(Index), MapNpcNum).Y Then ACoter = True: Exit Function
If GetPlayerX(Index) + 1 = MapNpc(GetPlayerMap(Index), MapNpcNum).X And GetPlayerY(Index) = MapNpc(GetPlayerMap(Index), MapNpcNum).Y Then ACoter = True: Exit Function
End Function

Sub EnMonture(ByVal Index As Long)
Dim s As Long
s = Val(GetVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "monture"))
Call SetPlayerSprite(Index, s)
Call SendPlayerData(Index)
End Sub

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

Function NbObjet(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim i As Long

    NbObjet = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function
        
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If GetPlayerInvItemValue(Index, i) <= 0 Then
                NbObjet = NbObjet + 1
            Else
                NbObjet = NbObjet + GetPlayerInvItemValue(Index, i)
            End If
        End If
    Next i
End Function

