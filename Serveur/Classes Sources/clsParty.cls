VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'::::::::::::::::::::::::::::
'::::::: Informations :::::::
'::::::::::::::::::::::::::::
'
'Tout ce qui est As Byte est relatif au party
'(Index à l'intérieur de l'équipe)
'
'Tout ce qui est As Integer est relatif au joueur
'(Index normal du joueur)

Private Partys() As clsPartys

Public Function CreateParty(ByVal Starter As Integer, ByVal Member As Integer) As Byte
Dim PartyIndex As Byte
    For PartyIndex = LBound(Partys) To UBound(Partys)
        If Partys(PartyIndex).Count = 0 Then Exit For
    Next
    If PartyIndex = UBound(Partys) + 1 Then
        ReDim Preserve Partys(LBound(Partys) To PartyIndex) As clsPartys
    End If
    
    Set Partys(PartyIndex) = New clsPartys
    Partys(PartyIndex).InitParty PartyIndex, Starter, Member
End Function

Public Sub DisbandParty(ByVal PartyIndex As Byte)
Dim I As Integer
    If PartyIndex = LBound(Partys) And PartyIndex = UBound(Partys) Then
        For I = Partys(PartyIndex).Count To 1 Step -1
            Partys(PartyIndex).RemoveMember I
        Next I
        ReDim Partys(PartyIndex) As clsPartys
    Else
        If PartyIndex < UBound(Partys) Then
            For I = Partys(PartyIndex).Count To 1 Step -1
                Partys(PartyIndex).RemoveMember I
            Next I
            For I = PartyIndex + 1 To UBound(Partys)
                Partys(I - 1) = Partys(I)
            Next I
            ReDim Preserve Partys(LBound(Partys) To I - 2) As clsPartys
        Else
            For I = Partys(PartyIndex).Count To 1 Step -1
                Partys(PartyIndex).RemoveMember I
            Next I
            ReDim Preserve Partys(LBound(Partys) To PartyIndex - 1) As clsPartys
        End If
    End If
End Sub

Public Function AddMember(ByVal PartyIndex As Byte, ByVal PlayerIndex As Integer) As Byte
    If PartyIndex < LBound(Partys) Or PartyIndex > UBound(Partys) Or PlayerIndex < 0 Or PlayerIndex > MAX_PLAYERS Then Exit Function
    AddMember = Partys(PartyIndex).AddMember(PlayerIndex)
End Function

Public Function MemberIndex(ByVal PartyIndex As Byte, ByVal PlayerIndex As Integer) As Byte
    If PartyIndex < LBound(Partys) Or PartyIndex > UBound(Partys) Or PlayerIndex < 0 Or PlayerIndex > MAX_PLAYERS Then Exit Function
    MemberIndex = Partys(PartyIndex).MemberIndex(PlayerIndex)
End Function

Public Function PlayerIndex(ByVal PartyIndex As Byte, ByVal MemberIndex As Integer) As Byte
    If PartyIndex < LBound(Partys) Or PartyIndex > UBound(Partys) Then Exit Function
    PlayerIndex = Partys(PartyIndex).PlayerIndex(MemberIndex)
End Function

Public Sub RemoveMember(ByVal PartyIndex As Byte, ByVal MemberIndex As Byte, Optional ByVal PlayerIndex As Integer = 0)
    If PartyIndex < 1 Or PartyIndex > UBound(Partys) Then Exit Sub
    Partys(PartyIndex).RemoveMember MemberIndex, PlayerIndex
    
    If Partys(PartyIndex).Count = 0 Then
        DisbandParty (PartyIndex)
    End If
End Sub

Public Property Get ShareExp(ByVal PartyIndex As Byte) As Byte
    If PartyIndex >= LBound(Partys) And PartyIndex <= UBound(Partys) Then ShareExp = Partys(PartyIndex).ShareExp
End Property

Public Property Get MemberCount(ByVal PartyIndex As Byte) As Byte
    If PartyIndex >= LBound(Partys) And PartyIndex <= UBound(Partys) Then MemberCount = Partys(PartyIndex).Count
End Property

Public Property Get Leader(ByVal PartyIndex As Byte) As Byte
    If PartyIndex >= LBound(Partys) And PartyIndex <= UBound(Partys) Then Leader = Partys(PartyIndex).Leader
End Property

'Public Property Let Leader(ByVal PartyIndex As Byte)
'    If PartyIndex >= LBound(Partys) And PartyIndex <= UBound(Partys) Then
'        If Leader > 0 And Leader <= Partys(PartyIndex).Count Then Partys(PartyIndex).Leader = Leader
'    End If
'End Property

Public Sub SetNewLeader(ByVal PartyIndex As Byte, ByVal Leader As Byte, Optional ByVal LeaderIndex As Integer = 0)
    If PartyIndex >= LBound(Partys) And PartyIndex <= UBound(Partys) Then
        If Leader > 0 And Leader <= Partys(PartyIndex).Count Then
            Partys(PartyIndex).Leader = Leader
        ElseIf LeaderIndex > 0 And LeaderIndex <= MAX_PLAYERS Then
            If Player(LeaderIndex).InParty = PartyIndex Then
                If Partys(PartyIndex).MemberIndex(LeaderIndex) > 0 Then Partys(PartyIndex).Leader = Partys(PartyIndex).MemberIndex(LeaderIndex)
            End If
        End If
    End If
End Sub

Private Sub Class_Initialize()
    ReDim Partys(1 To 1) As clsPartys
    Set Partys(1) = New clsPartys
End Sub
