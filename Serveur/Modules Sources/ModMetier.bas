Attribute VB_Name = "ModMetier"
Public Const MAX_DATA_METIER = 100
Public Const MAX_DATA_RECETTE = 9

Sub craftv2(ByVal Index As Long, ByVal rec As Long)
Dim n As Byte, i As Byte, v As Byte, w As Byte, s As Byte, r As Byte
Dim nb_item As Byte, f As Byte

    n = Player(Index).Char(Player(Index).CharNum).metier
    
    If n <= 0 Then
        Call BattleMsg(Index, "(Metier) Vous n'avez pas de métier", Red, 0)
        Exit Sub
    End If
    
    If metier(n).type <> METIER_CRAFT Then
        Call BattleMsg(Index, "(Metier) Vous n'avez pas le bon métier", Red, 0)
        Exit Sub
    End If
    
    For i = 0 To MAX_DATA_RECETTE
        If metier(n).data(i, 0) = rec Then
            r = i
            Exit For
        Else
            If i = MAX_DATA_RECETTE Then
                Call BattleMsg(Index, "(Metier) Cette recette n'est pas dans votre métier.", Red, 0)
                Exit Sub
            End If
        End If
    Next i
    
    With recette(metier(n).data(r, 0))
        If .craft(0) <= 0 Then
            Call BattleMsg(Index, "(Metier) Il n'y a rien a créer avec cette recette.", Red, 0)
            Exit Sub
        End If
        
        For i = 0 To MAX_DATA_RECETTE
            If .InCraft(i, 0) > 0 Then
                Exit For
            Else
                If i = MAX_DATA_RECETTE Then
                    Call BattleMsg(Index, "(Metier) Cette recette n'est pas bonne", Red, 0)
                End If
            End If
        Next i
        
        For i = 0 To MAX_DATA_RECETTE
            If .InCraft(i, 0) > 0 Then
                If AObjet(Index, .InCraft(i, 0)) > 0 Then
                    If NbObjet(Index, .InCraft(i, 0)) < .InCraft(i, 1) Then
                        Call BattleMsg(Index, "(Metier) Vous n'avez pas toute les ressources pour réaliser se craft. (2)", Red, 0)
                        Exit Sub
                    End If
                Else
                    Call BattleMsg(Index, "(Metier) Vous n'avez pas toute les ressources pour réaliser se craft.", Red, 0)
                    Exit Sub
                End If
            End If
        Next i
        
        v = FindOpenInvSlot(Index, .craft(0))
        If v > 0 Then
            nb_item = 0
            For i = 0 To MAX_DATA_RECETTE
                If .InCraft(i, 0) > 0 Then
                    nb_item = nb_item + 1
                    f = AObjet(Index, .InCraft(i, 0))
                    If GetPlayerInvItemValue(Index, f) - .InCraft(i, 1) <= 0 Then
                        Call SetPlayerInvItemNum(Index, f, 0)
                        Call SetPlayerInvItemValue(Index, f, 0)
                    Else
                        Call SetPlayerInvItemValue(Index, f, GetPlayerInvItemValue(Index, f) - .InCraft(i, 1))
                    End If
                End If
            Next i
            Math.Randomize
            w = Math.Round(Math.Rnd * 101)
            If w > 0 And w <= CraftReussiteV2(Index, nb_item) Then
                Call SetPlayerInvItemNum(Index, v, .craft(0))
                Call SetPlayerInvItemValue(Index, v, GetPlayerInvItemValue(Index, v) + .craft(1))
                If (item(.craft(0)).type >= ITEM_TYPE_WEAPON) And (item(.craft(0)).type <= ITEM_TYPE_SHIELD) Then Call SetPlayerInvItemDur(Index, v, item(.craft(0)).data1) Else Call SetPlayerInvItemDur(Index, v, 0)
                Call SendInventory(Index)
                Call BattleMsg(Index, "(Metier) Vous avez Crafter l'objet: " & item(.craft(0)).Name, BrightBlue, 0)
                If Player(Index).Char(Player(Index).CharNum).MetierLvl < 200 Then
                    Player(Index).Char(Player(Index).CharNum).MetierExp = Player(Index).Char(Player(Index).CharNum).MetierExp + metier(n).data(r, 1)
                    Call BattleMsg(Index, "(Metier) Vous avez gagné " & metier(n).data(r, 1) & " pts d'expérience.", BrightBlue, 0)
                Else
                    Call BattleMsg(Index, "(Metier) Vous ne pouver plus gagnez d'expérience", BrightBlue, 0)
                End If
            Else
                Call SendInventory(Index)
                Call BattleMsg(Index, "(Metier) Vous avez rater le Craft de l'objet: " & item(.craft(0)).Name, Red, 0)
                If Player(Index).Char(Player(Index).CharNum).MetierLvl < 200 Then
                    Player(Index).Char(Player(Index).CharNum).MetierExp = Player(Index).Char(Player(Index).CharNum).MetierExp + Math.Round(metier(n).data(r, 1) / 2)
                    Call BattleMsg(Index, "(Metier) Vous avez gagné " & Math.Round(metier(n).data(r, 1) / 2) & " pts d'expérience.", BrightBlue, 0)
                Else
                    Call BattleMsg(Index, "(Metier) Vous ne pouver plus gagnez d'expérience", BrightBlue, 0)
                End If
            End If
            Call checkLvlUpMetier(Index)
        Else
            Call BattleMsg(Index, "(Metier) Vous n'avez pas de place dans votre inventaire.", Red, 0)
            Exit Sub
        End If
    End With
End Sub


Sub craft(ByVal Index As Long, ByVal rec As Long)
Dim n As Byte, i As Byte, v As Byte, w As Byte, r As Byte, rb As Boolean
Dim rin As Boolean
Dim pin(0 To MAX_DATA_METIER) As Boolean, piv(0 To MAX_DATA_METIER) As Boolean
Dim pis(0 To MAX_DATA_METIER) As Byte, vr As Boolean
Dim nb_item As Byte
n = Player(Index).Char(Player(Index).CharNum).metier
If n > 0 Then
    If metier(n).type = METIER_CRAFT Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type <> TILE_TYPE_CRAFT Then
            Call BattleMsg(Index, "(Metier) Il vous faut une table de Craft .", Red, 0)
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
            Call BattleMsg(Index, "(Metier) Vous ne pouvez pas faire cette recettes", Red, 0)
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
                Call BattleMsg(Index, "(Metier) Cette recette n'est pas bonne", Red, 0)
                Exit Sub
            End If
            
            If .craft(0) <= 0 Then
                Call BattleMsg(Index, "(Metier) Il n'y a rien a créer avec cette recette.", Red, 0)
                Exit Sub
            End If

            For i = 1 To MAX_INV
                For v = 0 To MAX_DATA_METIER
                    If .InCraft(v, 0) > 0 Then
                        If pin(v) = False Then
                             If .InCraft(v, 0) = GetPlayerInvItemNum(Index, i) Then
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
                            If .InCraft(v, 1) <= GetPlayerInvItemValue(Index, i) Then
                                piv(v) = True
                             End If
                        End If
                    End If
                Next v
            Next i
            For i = 0 To MAX_DATA_METIER
                If (pin(i) = True And piv(i) = False) Or (pin(i) = False And .InCraft(i, 0) > 0) Then
                    Call BattleMsg(Index, "(Metier) Vous n'avez pas tout les objets pour créer cette recette.", Red, 0)
                    'Exit Sub
                End If
            Next i
            v = FindOpenInvSlot(Index, .craft(0))
            If v > 0 Then
                nb_item = 0
                For i = 0 To MAX_DATA_METIER
                    If pin(i) = True And piv(i) = True Then
                        nb_item = nb_item + 1
                        If GetPlayerInvItemValue(Index, pis(i)) - .InCraft(i, 1) <= 0 Then
                            Call SetPlayerInvItemNum(Index, pis(i), 0)
                            Call SetPlayerInvItemValue(Index, pis(i), 0)
                        Else
                            Call SetPlayerInvItemValue(Index, pis(i), GetPlayerInvItemValue(Index, pis(i)) - .InCraft(i, 1))
                        End If
                    End If
                Next i
                Math.Randomize
                w = Math.Round(Math.Rnd * 100)
                If w > 0 And w <= CraftReussite(Index, nb_item) Then
                    Call SetPlayerInvItemNum(Index, v, .craft(0))
                    Call SetPlayerInvItemValue(Index, v, GetPlayerInvItemValue(Index, v) + .craft(1))
                    If (item(.craft(0)).type >= ITEM_TYPE_WEAPON) And (item(.craft(0)).type <= ITEM_TYPE_SHIELD) Then Call SetPlayerInvItemDur(Index, v, item(.craft(0)).data1) Else Call SetPlayerInvItemDur(Index, v, 0)
                    Call SendInventory(Index)
                    Call BattleMsg(Index, "(Metier) Vous avez Crafter l'objet: " & item(.craft(0)).Name, BrightBlue, 0)
                    If Player(Index).Char(Player(Index).CharNum).MetierLvl < 200 Then
                        Player(Index).Char(Player(Index).CharNum).MetierExp = Player(Index).Char(Player(Index).CharNum).MetierExp + metier(n).data(r, 1)
                        Call BattleMsg(Index, "(Metier) Vous avez gagné " & metier(n).data(r, 1) & " pts d'expérience.", BrightBlue, 0)
                    Else
                        Call BattleMsg(Index, "(Metier) Vous ne pouver plus gagnez d'expérience", BrightBlue, 0)
                    End If
                Else
                    Call SendInventory(Index)
                    Call BattleMsg(Index, "(Metier) Vous avez rater le Craft de l'objet: " & item(.craft(0)).Name, Red, 0)
                    If Player(Index).Char(Player(Index).CharNum).MetierLvl < 200 Then
                        Player(Index).Char(Player(Index).CharNum).MetierExp = Player(Index).Char(Player(Index).CharNum).MetierExp + Math.Round(metier(n).data(r, 1) / 2)
                        Call BattleMsg(Index, "(Metier) Vous avez gagné " & Math.Round(metier(n).data(r, 1) / 2) & " pts d'expérience.", BrightBlue, 0)
                    Else
                        Call BattleMsg(Index, "(Metier) Vous ne pouver plus gagnez d'expérience", BrightBlue, 0)
                    End If
                End If
                Call checkLvlUpMetier(Index)
            Else
                Call BattleMsg(Index, "(Metier) Vous n'avez pas de place dans votre inventaire.", Red, 0)
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

Public Function CraftReussite(ByVal Index As Long, ByVal nb_item As Byte) As Byte
    CraftReussite = 20
    If nb_item > 2 Then
        If CraftReussite <= CraftReussite - ((nb_item - 1) * 5) Then
            CraftReussite = 0
        Else
            CraftReussite = CraftReussite - ((nb_item - 1) * 5)
        End If
    End If
    If CraftReussite + (Player(Index).Char(Player(Index).CharNum).MetierLvl - 1) > 99 Then
        CraftReussite = 99
    Else
        CraftReussite = CraftReussite + (Player(Index).Char(Player(Index).CharNum).MetierLvl - 1)
    End If
End Function

Public Function CraftReussiteV2(ByVal Index As Long, ByVal nb_item As Byte) As Integer
    CraftReussiteV2 = 20
    If nb_item > 2 Then
        If 0 <= CraftReussiteV2 - ((nb_item - 1) * 5) Then
            CraftReussiteV2 = 0
        Else
            CraftReussiteV2 = CraftReussiteV2 - ((nb_item - 1) * 5)
        End If
    End If
    If CraftReussiteV2 + (Player(Index).Char(Player(Index).CharNum).MetierLvl - 1) > 90 Then
        CraftReussiteV2 = 90
    Else
        CraftReussiteV2 = CraftReussiteV2 + (Player(Index).Char(Player(Index).CharNum).MetierLvl - 1)
    End If
End Function

Public Function DoubleDrop(ByVal Index As Long) As Byte
    DoubleDrop = 0
    If DoubleDrop + Math.Round(Player(Index).Char(Player(Index).CharNum).MetierLvl / 2) > 99 Then
        DoubleDrop = 99
    Else
        DoubleDrop = DoubleDrop + Math.Round(Player(Index).Char(Player(Index).CharNum).MetierLvl / 2)
    End If
End Function

Public Sub checkLvlUpMetier(ByVal Index As Long)
    If Player(Index).Char(Player(Index).CharNum).metier > 0 Then
        If Player(Index).Char(Player(Index).CharNum).MetierLvl < 200 Then
            Do While ((Player(Index).Char(Player(Index).CharNum).MetierLvl + 1) * 2) <= Player(Index).Char(Player(Index).CharNum).MetierExp
                Player(Index).Char(Player(Index).CharNum).MetierExp = Player(Index).Char(Player(Index).CharNum).MetierExp - ((Player(Index).Char(Player(Index).CharNum).MetierLvl + 1) * 2)
                Player(Index).Char(Player(Index).CharNum).MetierLvl = Player(Index).Char(Player(Index).CharNum).MetierLvl + 1
                Call BattleMsg(Index, "(Metier) Vous avez gagné un niveau.", BrightBlue, 0)
            Loop
        End If
    End If
End Sub
