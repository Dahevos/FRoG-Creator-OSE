VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPartys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Leader As Byte
Private Member(1 To MAX_PARTY_MEMBERS) As Integer
Public ShareExp As Byte

Private CountMember As Byte
Private PartyIndex As Byte

Private Sub Class_Initialize()
    CountMember = 0
End Sub

Public Sub InitParty(ByVal Party As Byte, ByVal Starter As Integer, ByVal Second As Integer)
    ShareExp = 2
    Leader = 1
    
    PartyIndex = Party
    
    AddMember Starter
    AddMember Second
End Sub

Public Function AddMember(ByVal Index As Integer) As Byte
    If Index < 0 Or Index > MAX_PLAYERS Then Exit Function
    
    CountMember = CountMember + 1
    Member(CountMember) = Index
    Player(Index).PartyPlayer = CountMember
    Player(Index).InParty = PartyIndex
    
    ' Partage des barres de vie
    Dim X As Byte, i As Byte
    For i = 1 To CountMember
        For X = 1 To CountMember
            Call SendDataTo(Member(i), "partyhp" & SEP_CHAR & Member(X) & SEP_CHAR & Player(Member(X)).InParty & SEP_CHAR & GetPlayerMaxHP(Member(X)) & SEP_CHAR & Player(Member(X)).Char(Player(Member(X)).CharNum).HP & SEP_CHAR & GetPlayerMaxMP(Member(X)) & SEP_CHAR & Player(Member(X)).Char(Player(Member(X)).CharNum).MP & END_CHAR)
        Next X
    Next i
    
    AddMember = CountMember
End Function

Public Sub RemoveMember(ByVal MIndex As Byte, Optional ByVal PIndex As Integer)
Dim i As Byte, X As Byte

    If CountMember > 0 Then
        If MIndex = 0 And PIndex > 0 And PIndex < MAX_PLAYERS Then
            For i = 1 To CountMember
                If Member(i) = PIndex Then
                    'If i <> Leader Then
                        For X = 1 To CountMember
                            'If X <> i Then Call SendDataTo(Member(i), "partyindex" & SEP_CHAR & Member(X) & SEP_CHAR & Player(Member(X)).InParty & END_CHAR)
                            Call SendDataTo(Member(X), "partyindex" & SEP_CHAR & Member(i) & SEP_CHAR & 0 & END_CHAR)
                        Next X
                        Player(Member(i)).InParty = 0
                        Player(Member(i)).PartyPlayer = 0
                        Player(Member(i)).InvitedBy = 0
                        
                        If i < CountMember Then
                            For X = i + 1 To CountMember
                                Member(X - 1) = Member(X)
                                
                                Player(Member(X)).PartyPlayer = X - 1
                            Next X
                            Member(X - 1) = 0
                        Else
                            Member(i) = 0
                        End If
                        
                        CountMember = CountMember - 1
                    'End If
                End If
            Next i
        'ElseIf MIndex > 0 And MIndex <= CountMember And MIndex <> Leader Then
        ElseIf MIndex > 0 And MIndex <= CountMember Then
            For X = 1 To CountMember
                Call SendDataTo(Member(X), "partyindex" & SEP_CHAR & Member(MIndex) & SEP_CHAR & 0 & END_CHAR)
            Next X
            
            Player(Member(MIndex)).InParty = 0
            Player(Member(MIndex)).PartyPlayer = 0
            Player(Member(MIndex)).InvitedBy = 0
                        
            Member(MIndex) = 0
            If MIndex < CountMember Then
                For X = MIndex + 1 To CountMember
                    Member(X - 1) = Member(X)
                    
                    Player(Member(X)).PartyPlayer = X - 1
                Next X
                Member(X - 1) = 0
            Else
                Member(MIndex) = 0
            End If
            CountMember = CountMember - 1
        End If
    End If
End Sub

Public Function MemberIndex(ByVal PIndex As Integer) As Byte
Dim i As Byte
    
    For i = 1 To CountMember
        If Member(i) = PIndex Then MemberIndex = i: Exit Function
    Next i
End Function

Public Property Get PlayerIndex(ByVal MIndex As Byte) As Integer
    If MIndex > 0 And MIndex <= CountMember Then PlayerIndex = Member(MIndex)
End Property

Public Property Get Count() As Byte
    Count = CountMember
End Property
