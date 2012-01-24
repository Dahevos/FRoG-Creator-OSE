Attribute VB_Name = "ModMetier"
Public Const MAX_DATA_METIER = 100

Sub craft(ByVal index As Long, ByVal rec As Long)
Dim n As Byte, i As Byte, v As Byte, w As Byte, r As Byte, rb As Boolean
Dim rin As Boolean
Dim pin(0 To MAX_DATA_METIER) As Boolean, piv(0 To MAX_DATA_METIER) As Boolean
Dim pis(0 To MAX_DATA_METIER) As Byte, vr As Boolean
Dim nb_item As Byte
n = Player(index).Char(Player(index).CharNum).metier
If n > 0 Then
    If metier(n).type = METIER_CRAFT Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type <> TILE_TYPE_METIER Then
            Call BattleMsg(index, "(Metier) Il vous faut une table de Craft !", Red, 0)
                Exit Sub
        End If
    
        rb = False
        For i = 0 To MAX_DATA_METIER
            If metier(n).data(i, 0) = rec Then
                r = i
                rb = True
                Exit For
            End If
        Next i
        If Not rb Then
            Call BattleMsg(index, "(Metier) Vous ne pouvez pas faire cette recettes", Red, 0)
            Exit Sub
        End If
        With recette(metier(n).data(r, 0))
            rin = False
            For i = 0 To MAX_DATA_METIER
                If .InCraft(i, 0) > 0 Then
                    rin = True
                    Exit For
                End If
            Next i
            If rin = False Then
                Call BattleMsg(index, "(Metier) Cette recette n'est pas bonnes", Red, 0)
                Exit Sub
            End If
            If .craft(0) <= 0 Then
                Call BattleMsg(index, "(Metier) Il n'y a rien a créer avec cette recette.", Red, 0)
                Exit Sub
            End If
            For i = 1 To MAX_INV
                For v = 0 To MAX_DATA_METIER
                    If .InCraft(v, 0) > 0 Then
                        If pin(v) = False Then
                             If .InCraft(v, 0) = GetPlayerInvItemNum(index, i) Then
                                vr = False
                                For w = 0 To MAX_DATA_METIER
                                    If pis(w) = i Then
                                        vr = True
                                    End If
                                Next w
                                If vr = False Then
                                    pin(v) = True
                                    pis(v) = i
                                End If
                             End If
                        End If
                        If pin(v) = True And piv(v) = False Then
                            If .InCraft(v, 1) <= GetPlayerInvItemValue(index, i) Then
                                piv(v) = True
                             End If
                        End If
                    End If
                Next v
            Next i
            
            For i = 0 To MAX_DATA_METIER
                If (pin(i) = True And piv(i) = False) Or (pin(i) = False And .InCraft(i, 0) > 0) Then
                    Call BattleMsg(index, "(Metier) Vous n'avez pas tout les objets pour créer cette recette.", Red, 0)
                    Exit Sub
                End If
            Next i
            
            v = FindOpenInvSlot(index, .craft(0))
            If v <> 0 Then
                nb_item = 0
                For i = 0 To MAX_DATA_METIER
                    If pin(i) = True And piv(i) = True Then
                        nb_item = nb_item + 1
                        If GetPlayerInvItemValue(index, pis(i)) - .InCraft(i, 1) <= 0 Then
                            Call SetPlayerInvItemNum(index, pis(i), 0)
                            Call SetPlayerInvItemValue(index, pis(i), 0)
                        Else
                            Call SetPlayerInvItemValue(index, pis(i), GetPlayerInvItemValue(index, pis(i)) - .InCraft(i, 1))
                        End If
                    End If
                Next i
                Math.Randomize
                w = Math.Round(Math.Rnd * 100)
                If w > 0 And w <= CraftReussite(index, nb_item) Then
                    Call SetPlayerInvItemNum(index, v, .craft(0))
                    Call SetPlayerInvItemValue(index, v, GetPlayerInvItemValue(index, v) + .craft(1))
                    If (item(.craft(0)).type >= ITEM_TYPE_WEAPON) And (item(.craft(0)).type <= ITEM_TYPE_SHIELD) Then Call SetPlayerInvItemDur(index, v, item(.craft(0)).data1) Else Call SetPlayerInvItemDur(index, v, 0)
                    Call SendInventory(index)
                    Call BattleMsg(index, "(Metier) Vous avez Crafter l'objet: " & item(.craft(0)).Name, BrightBlue, 0)
                    If Player(index).Char(Player(index).CharNum).MetierLvl < 200 Then
                        Player(index).Char(Player(index).CharNum).MetierExp = Player(index).Char(Player(index).CharNum).MetierExp + metier(n).data(r, 1)
                        Call BattleMsg(index, "(Metier) Vous avez gagné " & metier(n).data(r, 1) & " pts d'expérience.", BrightBlue, 0)
                    Else
                        Call BattleMsg(index, "(Metier) Vous ne pouver plus gagnez d'expérience", BrightBlue, 0)
                    End If
                Else
                    Call SendInventory(index)
                    Call BattleMsg(index, "(Metier) Vous avez rater le Craft de l'objet: " & item(.craft(0)).Name, Red, 0)
                    If Player(index).Char(Player(index).CharNum).MetierLvl < 200 Then
                        Player(index).Char(Player(index).CharNum).MetierExp = Player(index).Char(Player(index).CharNum).MetierExp + Math.Round(metier(n).data(r, 1) / 2)
                        Call BattleMsg(index, "(Metier) Vous avez gagné " & Math.Round(metier(n).data(r, 1) / 2) & " pts d'expérience.", BrightBlue, 0)
                    Else
                        Call BattleMsg(index, "(Metier) Vous ne pouver plus gagnez d'expérience", BrightBlue, 0)
                    End If
                End If
                Call checkLvlUpMetier(index)
            Else
                Call BattleMsg(index, "(Metier) Vous n'avez pas de place dans votre inventaire.", Red, 0)
                Exit Sub
            End If
        End With
    End If
End If
End Sub

Public Function InMetier(ByVal metiernum As Long, ByVal npcnum As Long) As Byte
Dim i As Byte
    For i = 0 To MAX_DATA_METIER
        If metier(metiernum).data(i, 0) = npcnum Then
            InMetier = i
            Exit Function
        End If
    Next i
    InMetier = 10
End Function

Public Function CraftReussite(ByVal index As Long, ByVal nb_item As Byte) As Byte
    CraftReussite = 20
    If nb_item > 2 Then
        If CraftReussite <= CraftReussite - ((nb_item - 1) * 5) Then
            CraftReussite = 0
        Else
            CraftReussite = CraftReussite - ((nb_item - 1) * 5)
        End If
    End If
    If CraftReussite + (Player(index).Char(Player(index).CharNum).MetierLvl - 1) > 99 Then
        CraftReussite = 99
    Else
        CraftReussite = CraftReussite + (Player(index).Char(Player(index).CharNum).MetierLvl - 1)
    End If
End Function

Public Function DoubleDrop(ByVal index As Long) As Byte
    DoubleDrop = 0
    If DoubleDrop + Math.Round(Player(index).Char(Player(index).CharNum).MetierLvl / 2) > 99 Then
        DoubleDrop = 99
    Else
        DoubleDrop = DoubleDrop + Math.Round(Player(index).Char(Player(index).CharNum).MetierLvl / 2)
    End If
End Function

Public Sub checkLvlUpMetier(ByVal index As Long)
    If Player(index).Char(Player(index).CharNum).metier > 0 Then
        If Player(index).Char(Player(index).CharNum).MetierLvl < 200 Then
            Do While ((Player(index).Char(Player(index).CharNum).MetierLvl + 1) * 2) <= Player(index).Char(Player(index).CharNum).MetierExp
                Player(index).Char(Player(index).CharNum).MetierExp = Player(index).Char(Player(index).CharNum).MetierExp - ((Player(index).Char(Player(index).CharNum).MetierLvl + 1) * 2)
                Player(index).Char(Player(index).CharNum).MetierLvl = Player(index).Char(Player(index).CharNum).MetierLvl + 1
                Call BattleMsg(index, "(Metier) Vous avez gagné un niveau.", BrightBlue, 0)
            Loop
        End If
    End If
End Sub
