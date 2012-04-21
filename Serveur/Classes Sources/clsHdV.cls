VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHdV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ObjChild As clsHdVChild
Private Achats As Collection
Private Ventes As Collection

Private Type ObjChildRec
    Index As Integer
    ItemNum As Integer
    ItemVal As Byte
    ItemDurMin As Integer
    UnitPrice As Long

    CurTime As Date
    Actif As Boolean
End Type

Private ObjChildLS() As ObjChildRec

Private Const TradingItemIndex As Integer = 1    ' Le numéro de l'objet représentant la monnaie de l'HdV
Private Const DeleteOneByOne As Boolean = False  ' False = Ne supprime pas les Achats/Ventes Annulés ou Terminés
Private Const OptimiseOnClose As Boolean = False ' False = Ne supprime pas les Achats/Ventes Annulés ou Terminés à la fermeture
Private Const DateToOptimise As Single = 7       ' 7 = Nombres de jours après lesquels les Achats/Ventes en cours sont supprimés lors d'une optimisation

Private Sub Class_Initialize()
    Set Achats = New Collection
    Set Ventes = New Collection
    
    If LCase$(Dir(App.Path & "\achats", vbDirectory)) <> "achats" Then Call MkDir(App.Path & "\Achats")
    If LCase$(Dir(App.Path & "\ventes", vbDirectory)) <> "ventes" Then Call MkDir(App.Path & "\Ventes")
    
    LoadHdV
End Sub

Private Sub Class_Terminate()
    If OptimiseOnClose Then OptimiseAchatVente True
    SaveHdV
    
    Set Achats = Nothing
    Set Ventes = Nothing
    Set ObjChild = Nothing
End Sub

Public Function AddAchat(ByVal Index As Integer, ByVal UnitPrice As Long, ByVal ItemNum As Integer, ByVal ItemVal As Byte, ByVal ItemDurMin As Integer, Optional ByVal Check As Boolean = True) As Integer
Dim s As Byte, i As Integer
    AddAchat = -1
    If ItemVal = 0 Then Exit Function
    
    s = CByte(AObjet(Index, TradingItemIndex))
    If s = 0 Then Exit Function
    If Player(Index).Char(Player(Index).CharNum).Inv(s).value < (ItemVal * UnitPrice) Then Exit Function
    TakeItem Index, TradingItemIndex, (ItemVal * UnitPrice)
    
    AddAchat = 0
    If Check Then
        i = 1
        Do While i < Ventes.Count
            Set ObjChild = Ventes.item(i)
            If ObjChild.Actif And ObjChild.ItemNum = ItemNum And ObjChild.ItemDurMin >= ItemDurMin Then
                If ObjChild.UnitPrice <= UnitPrice And ItemVal > 0 Then
                    s = CByte(FindOpenInvSlot(Index, ItemNum))
                    If s = 0 Then PlayerMsg Index, "Certaines ventes correspondent, mais vous n'avez plus d'espace libre.", 4: Exit Do
                    If item(ItemNum).type = ITEM_TYPE_CURRENCY Then
                        If ItemVal >= ObjChild.ItemVal Then
                            ItemVal = ItemVal - ObjChild.ItemVal
                            If ObjChild.UnitPrice < UnitPrice Then GiveItem Index, TradingItemIndex, UnitPrice - ObjChild.UnitPrice
                            GiveItem ObjChild.Index, TradingItemIndex, ObjChild.UnitPrice
                            GiveItem Index, ItemNum, ObjChild.ItemVal
                            
                            If DeleteOneByOne Then
                                Ventes.Remove i
                                i = i - 1
                            Else
                                ObjChild.Actif = False
                            End If
                        Else
                            ObjChild.ItemVal = ObjChild.ItemVal - ItemVal
                            If ObjChild.UnitPrice < UnitPrice Then GiveItem Index, TradingItemIndex, UnitPrice - ObjChild.UnitPrice
                            GiveItem ObjChild.Index, TradingItemIndex, ObjChild.UnitPrice
                            GiveItem Index, ItemNum, ItemVal
                            ItemVal = 0
                        End If
                    Else
                        ItemVal = ItemVal - 1
                        If ObjChild.UnitPrice < UnitPrice Then GiveItem Index, TradingItemIndex, UnitPrice - ObjChild.UnitPrice
                        GiveItem ObjChild.Index, TradingItemIndex, ObjChild.UnitPrice
                        GiveItem Index, ItemNum, 1
                        
                        If DeleteOneByOne Then
                            Ventes.Remove i
                            i = i - 1
                        Else
                            ObjChild.Actif = False
                        End If
                    End If
                End If
            End If
            If ItemVal = 0 Then Exit Function
            i = i + 1
        Loop
    End If
    
    Set ObjChild = New clsHdVChild
    ObjChild.CurTime = Date + Time
    
    ObjChild.Index = Index
    ObjChild.UnitPrice = UnitPrice
    ObjChild.ItemNum = ItemNum
    ObjChild.ItemVal = ItemVal
    ObjChild.ItemDurMin = ItemDurMin
    ObjChild.Actif = True
    
    Achats.add ObjChild
    
    Set ObjChild = Nothing
    AddAchat = Achats.Count
End Function

Public Function AddVente(ByVal Index As Integer, ByVal UnitPrice As Long, ByVal ItemNum As Integer, ByVal ItemVal As Integer, ByVal ItemDurMin As Integer, Optional ByVal Check As Boolean = True) As Integer
Dim s As Byte, i As Integer
    AddVente = -1
    
    s = CByte(AObjet(Index, ItemNum))
    If s = 0 Then Exit Function
    If Player(Index).Char(Player(Index).CharNum).Inv(s).value < ItemVal Then Exit Function
    TakeItem Index, ItemNum, ItemVal
    
    AddVente = 0
    
    If Check Then
        i = 1
        Do While i < Achats.Count
            Set ObjChild = Achats.item(i)
            If ObjChild.Actif And ObjChild.ItemNum = ItemNum And ObjChild.ItemDurMin <= ItemDurMin Then
                If ObjChild.UnitPrice >= UnitPrice Then
                    s = CByte(FindOpenInvSlot(Index, TradingItemIndex))
                    If s = 0 Then PlayerMsg Index, "Certains achats correspondent, mais vous n'avez plus d'espace libre.", 4: Exit Do
                    If item(ItemNum).type = ITEM_TYPE_CURRENCY Then
                        If ItemVal >= ObjChild.ItemVal Then
                            ItemVal = ItemVal - ObjChild.ItemVal
                            If ObjChild.UnitPrice > UnitPrice Then GiveItem ObjChild.Index, TradingItemIndex, ObjChild.UnitPrice - UnitPrice
                            GiveItem Index, TradingItemIndex, UnitPrice
                            GiveItem ObjChild.Index, ItemNum, ObjChild.ItemVal
                            
                            If DeleteOneByOne Then
                                Achats.Remove i
                                i = i - 1
                            Else
                                ObjChild.Actif = False
                            End If
                        Else
                            ObjChild.ItemVal = ObjChild.ItemVal - ItemVal
                            If ObjChild.UnitPrice > UnitPrice Then GiveItem ObjChild.Index, TradingItemIndex, ObjChild.UnitPrice - UnitPrice
                            GiveItem Index, TradingItemIndex, UnitPrice
                            GiveItem ObjChild.Index, ItemNum, ItemVal
                            ItemVal = 0
                        End If
                        If ItemVal = 0 Then Exit Function
                    Else
                        ObjChild.ItemVal = ObjChild.ItemVal - 1
                        If ObjChild.UnitPrice > UnitPrice Then GiveItem ObjChild.Index, TradingItemIndex, ObjChild.UnitPrice - UnitPrice
                        GiveItem Index, TradingItemIndex, UnitPrice
                        GiveItem ObjChild.Index, ItemNum, 1
                        
                        If ObjChild.ItemVal = 0 Then
                            If DeleteOneByOne Then
                                Achats.Remove i
                                i = i - 1
                            Else
                                ObjChild.Actif = False
                            End If
                        End If
                        Exit Function
                    End If
                End If
            End If
            i = i + 1
        Loop

    End If
    
    Set ObjChild = New clsHdVChild
    ObjChild.CurTime = Date + Time
    
    ObjChild.Index = Index
    ObjChild.UnitPrice = UnitPrice
    ObjChild.ItemNum = ItemNum
    ObjChild.ItemVal = ItemVal
    ObjChild.ItemDurMin = ItemDurMin
    ObjChild.Actif = True
    
    Ventes.add ObjChild
    
    Set ObjChild = Nothing
    AddVente = Ventes.Count
End Function

Public Sub CancelAchat(ByVal Index As Integer, ByVal AchatIndex As Integer)
Dim i As Integer
Dim Count As Byte

    If Not DeleteOneByOne Then
        If AchatIndex < 1 Or AchatIndex > Achats.Count Then Exit Sub
        Set ObjChild = Achats.item(AchatIndex)
        If ObjChild.Index <> Index Then Exit Sub
        ObjChild.Actif = False
        
        Count = CByte(FindOpenInvSlot(Index, TradingItemIndex))
        If Count = 0 Then Exit Sub
        Player(Index).Char(Player(Index).CharNum).Inv(Count).num = TradingItemIndex
        Player(Index).Char(Player(Index).CharNum).Inv(Count).value = Player(Index).Char(Player(Index).CharNum).Inv(Count).value + (ObjChild.ItemVal * ObjChild.UnitPrice)
        Player(Index).Char(Player(Index).CharNum).Inv(Count).Dur = ObjChild.ItemDurMin
        SendInventoryUpdate Index, Count
        ObjChild.ItemDurMin = 0
        ObjChild.ItemNum = 0
        ObjChild.ItemVal = 0
        ObjChild.UnitPrice = 0
        
    ElseIf AchatIndex > 0 Then
        For i = 1 To Achats.Count
            Set ObjChild = Achats.item(i)
            If ObjChild.Index = Index Then
                Count = Count + 1
                If Count = AchatIndex Then
                    Count = CByte(FindOpenInvSlot(Index, TradingItemIndex))
                    If Count = 0 Then Exit Sub
                    Player(Index).Char(Player(Index).CharNum).Inv(Count).num = TradingItemIndex
                    Player(Index).Char(Player(Index).CharNum).Inv(Count).value = Player(Index).Char(Player(Index).CharNum).Inv(Count).value + (ObjChild.ItemVal * ObjChild.UnitPrice)
                    Player(Index).Char(Player(Index).CharNum).Inv(Count).Dur = ObjChild.ItemDurMin
                    SendInventoryUpdate Index, Count
        
                    Achats.Remove i
                    Exit Sub
                End If
            End If
        Next
    End If
End Sub

Public Sub CancelVente(ByVal Index As Integer, ByVal VenteIndex As Integer)
Dim i As Integer
Dim Count As Byte

    If Not DeleteOneByOne Then
        If VenteIndex < 1 Or VenteIndex > Ventes.Count Then Exit Sub
        Set ObjChild = Ventes.item(VenteIndex)
        If ObjChild.Index <> Index Then Exit Sub
        ObjChild.Actif = False
        
        Count = CByte(FindOpenInvSlot(Index, ObjChild.ItemNum))
        If Count = 0 Then Exit Sub
        Player(Index).Char(Player(Index).CharNum).Inv(Count).num = ObjChild.ItemNum
        Player(Index).Char(Player(Index).CharNum).Inv(Count).value = Player(Index).Char(Player(Index).CharNum).Inv(Count).value + ObjChild.ItemVal
        Player(Index).Char(Player(Index).CharNum).Inv(Count).Dur = ObjChild.ItemDurMin
        SendInventoryUpdate Index, Count
        ObjChild.ItemDurMin = 0
        ObjChild.ItemNum = 0
        ObjChild.ItemVal = 0
        ObjChild.UnitPrice = 0
        
    ElseIf VenteIndex > 0 Then
        For i = 1 To Ventes.Count
            Set ObjChild = Ventes.item(i)
            If ObjChild.Index = Index Then
                Count = Count + 1
                If Count = VenteIndex Then
                    Count = CByte(FindOpenInvSlot(Index, ObjChild.ItemNum))
                    If Count = 0 Then Exit Sub
                    Player(Index).Char(Player(Index).CharNum).Inv(Count).num = ObjChild.ItemNum
                    Player(Index).Char(Player(Index).CharNum).Inv(Count).value = Player(Index).Char(Player(Index).CharNum).Inv(Count).value + ObjChild.ItemVal
                    Player(Index).Char(Player(Index).CharNum).Inv(Count).Dur = ObjChild.ItemDurMin
                    SendInventoryUpdate Index, Count
                    
                    Ventes.Remove i
                    Exit Sub
                End If
            End If
        Next
    End If
End Sub

Public Sub OptimiseAchatVente(ByVal Del As Boolean, Optional ByVal DateDel As Double)
Dim i As Integer, Max As Integer, CurTime As Date
    If Not Del Then Exit Sub
    If Not DeleteOneByOne And Del Then
        If Achats.Count > Ventes.Count Then Max = Achats.Count Else Max = Ventes.Count
        For i = Max To 1 Step -1
            If i <= Achats.Count Then
                Set ObjChild = Achats.item(i)
                If Not ObjChild.Actif Then Achats.Remove i
            End If
            If i <= Ventes.Count Then
                Set ObjChild = Ventes.item(i)
                If Not ObjChild.Actif Then Ventes.Remove i
            End If
        Next
    End If
    If Del And DateDel > 0 Then
        If Achats.Count > Ventes.Count Then Max = Achats.Count Else Max = Ventes.Count
        CurTime = Date + Time
        For i = Max To 1 Step -1
            If i <= Achats.Count Then
                Set ObjChild = Achats.item(i)
                If ObjChild.CurTime + DateDel < CurTime Then Achats.Remove i
            End If
            If i <= Ventes.Count Then
                Set ObjChild = Ventes.item(i)
                If ObjChild.CurTime + DateDel < CurTime Then Ventes.Remove i
            End If
        Next
    End If
End Sub

Sub CheckAll()
Dim I2 As Integer, i As Integer, s As Byte, S1 As Byte, S2 As Byte, DiffMoney As Long
Dim ObjV As clsHdVChild, ObjA As clsHdVChild
    
    For i = Achats.Count To 1 Step -1
        Set ObjA = Achats.item(i)
        If ObjA.Actif Then
            For I2 = Ventes.Count To 1 Step -1
                Set ObjV = Ventes.item(I2)
                If ObjV.Actif And ObjV.ItemNum = ObjA.ItemNum And ObjV.UnitPrice <= ObjA.UnitPrice And ObjV.ItemDurMin >= ObjA.ItemDurMin Then
                    DiffMoney = ObjA.UnitPrice - ObjV.UnitPrice
                    
                    s = CByte(FindOpenInvSlot(ObjA.Index, ObjA.ItemNum))
                    S1 = CByte(FindOpenInvSlot(ObjV.Index, TradingItemIndex))
                    If s = 0 Or S1 = 0 Then GoTo EndFor
                    
                    If ObjA.ItemVal >= ObjV.ItemVal Then
                        
                        Player(ObjV.Index).Char(Player(ObjV.Index).CharNum).Inv(S1).num = TradingItemIndex
                        Player(ObjV.Index).Char(Player(ObjV.Index).CharNum).Inv(S1).value = Player(ObjV.Index).Char(Player(ObjV.Index).CharNum).Inv(S1).value + IIf(ObjV.ItemVal = 0, 1, ObjV.ItemVal) * ObjV.UnitPrice
                        SendInventoryUpdate ObjV.Index, S1
                        
                        Player(ObjA.Index).Char(Player(ObjA.Index).CharNum).Inv(s).num = ObjA.ItemNum
                        Player(ObjA.Index).Char(Player(ObjA.Index).CharNum).Inv(s).Dur = ObjV.ItemDurMin
                        Player(ObjA.Index).Char(Player(ObjA.Index).CharNum).Inv(s).value = Player(ObjA.Index).Char(Player(ObjA.Index).CharNum).Inv(s).value + ObjV.ItemVal
                        SendInventoryUpdate ObjA.Index, s
                        
                        If DiffMoney > 0 Then
                            S2 = CByte(FindOpenInvSlot(ObjA.Index, TradingItemIndex))
                            If S2 = 0 Then GoTo SkipIf
                            Player(ObjA.Index).Char(Player(ObjA.Index).CharNum).Inv(S2).num = TradingItemIndex
                            Player(ObjA.Index).Char(Player(ObjA.Index).CharNum).Inv(S2).value = Player(ObjA.Index).Char(Player(ObjA.Index).CharNum).Inv(S2).value + ObjV.ItemVal * DiffMoney
                            SendInventoryUpdate ObjA.Index, S2
                        End If
SkipIf:
                        S1 = 0
                    End If
                    
                    If ObjV.ItemVal >= ObjA.ItemVal Then
                        If S1 > 0 Then
                            
                        End If
                        s = 0
                    End If
                    
                    If Not DeleteOneByOne Then
                        If s = 0 Then ObjA.Actif = False
                        If S1 = 0 Then ObjV.Actif = False
                    Else
                        If s = 0 Then Achats.Remove i
                        If S1 = 0 Then Ventes.Remove I2
                    End If
                    
                    Set ObjV = Nothing
                    Exit For
                End If
EndFor:
                Set ObjV = Nothing
            Next
        End If
        Set ObjA = Nothing
    Next
End Sub

Public Property Get ListeAchats() As Collection
    Set ListeAchats = Achats
End Property

Public Property Get ListeVentes() As Collection
    Set ListeVentes = Ventes
End Property

Public Sub SaveHdV()
Dim FileName As String
Dim i As Integer
Dim f As Long

    f = FreeFile
    If Achats.Count > 0 Then
        ReDim ObjChildLS(1 To Achats.Count) As ObjChildRec
        For i = 1 To Achats.Count
            FileName = "Achats\" & i & ".hdv"
            
            Set ObjChild = Achats.item(i)
            ObjChildLS(i).Actif = ObjChild.Actif
            ObjChildLS(i).CurTime = ObjChild.CurTime
            ObjChildLS(i).Index = ObjChild.Index
            ObjChildLS(i).ItemDurMin = ObjChild.ItemDurMin
            ObjChildLS(i).ItemNum = ObjChild.ItemNum
            ObjChildLS(i).ItemVal = ObjChild.ItemVal
            ObjChildLS(i).UnitPrice = ObjChild.UnitPrice
            
            Open App.Path & "\" & FileName For Binary As #f
                Put #f, , ObjChildLS(i)
            Close #f
        Next i
    End If
    f = FreeFile
    If Ventes.Count > 0 Then
        ReDim ObjChildLS(1 To Ventes.Count) As ObjChildRec
        For i = 1 To Ventes.Count
            FileName = "Ventes\" & i & ".hdv"
            
            Set ObjChild = Ventes.item(i)
            ObjChildLS(i).Actif = ObjChild.Actif
            ObjChildLS(i).CurTime = ObjChild.CurTime
            ObjChildLS(i).Index = ObjChild.Index
            ObjChildLS(i).ItemDurMin = ObjChild.ItemDurMin
            ObjChildLS(i).ItemNum = ObjChild.ItemNum
            ObjChildLS(i).ItemVal = ObjChild.ItemVal
            ObjChildLS(i).UnitPrice = ObjChild.UnitPrice
            
            Open App.Path & "\" & FileName For Binary As #f
                Put #f, , ObjChildLS
            Close #f
        Next i
    End If
    
    Set ObjChild = Nothing
End Sub

Public Sub LoadHdV()
Dim FileName As String
Dim i As Integer
Dim f As Long
    
    FileName = "Achats\1.hdv"
    i = 1
    Do While FileExist(FileName)
        f = FreeFile
        ReDim Preserve ObjChildLS(1 To i) As ObjChildRec
        Open FileName For Binary Access Read As #f
            Get #f, , ObjChildLS(i)
        Close #f
        Set ObjChild = New clsHdVChild
        ObjChild.Actif = ObjChildLS(i).Actif
        ObjChild.CurTime = ObjChildLS(i).CurTime
        ObjChild.Index = ObjChildLS(i).Index
        ObjChild.ItemDurMin = ObjChildLS(i).ItemDurMin
        ObjChild.ItemNum = ObjChildLS(i).ItemNum
        ObjChild.ItemVal = ObjChildLS(i).ItemVal
        ObjChild.UnitPrice = ObjChildLS(i).UnitPrice
        Achats.add ObjChild
        
        i = i + 1
        FileName = "Achats\" & i & ".hdv"
    Loop
    FileName = "Ventes\1.hdv"
    i = 1
    Do While FileExist(FileName)
        f = FreeFile
        ReDim Preserve ObjChildLS(1 To i) As ObjChildRec
        Open App.Path & "\" & FileName For Binary As #f
            Get #f, , ObjChildLS(i)
        Close #f
        Set ObjChild = New clsHdVChild
        ObjChild.Actif = ObjChildLS(i).Actif
        ObjChild.CurTime = ObjChildLS(i).CurTime
        ObjChild.Index = ObjChildLS(i).Index
        ObjChild.ItemDurMin = ObjChildLS(i).ItemDurMin
        ObjChild.ItemNum = ObjChildLS(i).ItemNum
        ObjChild.ItemVal = ObjChildLS(i).ItemVal
        ObjChild.UnitPrice = ObjChildLS(i).UnitPrice
        Ventes.add ObjChild
        
        i = i + 1
        FileName = "Ventes\" & i & ".hdv"
    Loop
    
    Set ObjChild = Nothing
End Sub
