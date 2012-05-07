Attribute VB_Name = "modInfoBulle"
Option Explicit
Public IBVisible As Boolean
Public IBCharge As Boolean
Public IBJoueur As Boolean
Public IBAdmin As Boolean
Public IBErr As Boolean
Public IBCJoueur As Long
Public IBCAdmin As Long

Sub ChargIBOpt()
If Not FileExist("\Options.ini") Then Exit Sub
IBJoueur = CBool(GetVar(App.Path & "\Options.ini", "INFOBULLE", "MJoueur"))
IBAdmin = CBool(GetVar(App.Path & "\Options.ini", "INFOBULLE", "MAdmin"))
IBErr = CBool(GetVar(App.Path & "\Options.ini", "INFOBULLE", "MErr"))
End Sub
Sub SauvIBOpt()
Call PutVar(App.Path & "\Options.ini", "INFOBULLE", "MJoueur", STR$(IBJoueur))
Call PutVar(App.Path & "\Options.ini", "INFOBULLE", "MAdmin", STR$(IBAdmin))
Call PutVar(App.Path & "\Options.ini", "INFOBULLE", "MErr", STR$(IBErr))
End Sub
Sub IBMsg(ByVal Msg As String, Optional ByVal Coul As Long)
On Error Resume Next

If Coul = BrightRed Then
frmServer.ctlSysTrayBalloon.BalloonTipShow _
        "FRoG Server", _
        Msg, _
        NIIF_ERROR, _
        8000
Else
frmServer.ctlSysTrayBalloon.BalloonTipShow _
        "FRoG Server", _
        Msg, _
        NIIF_INFO, _
        8000
End If
End Sub



