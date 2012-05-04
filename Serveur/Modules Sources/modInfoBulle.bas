Attribute VB_Name = "modInfoBulle"
Option Explicit
Public IBVisible As Boolean
Public IBCharge As Boolean
Public IBTout As Boolean
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
IBTout = CBool(GetVar(App.Path & "\Options.ini", "INFOBULLE", "MTout"))
IBCJoueur = 1
IBCAdmin = 1
If IBJoueur Then frmOptInfoBulle.mj.value = Checked Else frmOptInfoBulle.mj.value = Unchecked
If IBErr Then frmOptInfoBulle.mer.value = Checked Else frmOptInfoBulle.mer.value = Unchecked
If IBAdmin Then frmOptInfoBulle.ma.value = Checked Else frmOptInfoBulle.ma.value = Unchecked
If IBTout Then frmOptInfoBulle.mt.value = Checked Else frmOptInfoBulle.mt.value = Unchecked
End Sub
Sub SauvIBOpt()
Call PutVar(App.Path & "\Options.ini", "INFOBULLE", "MTout", STR$(IBTout))
Call PutVar(App.Path & "\Options.ini", "INFOBULLE", "MJoueur", STR$(IBJoueur))
Call PutVar(App.Path & "\Options.ini", "INFOBULLE", "MAdmin", STR$(IBAdmin))
Call PutVar(App.Path & "\Options.ini", "INFOBULLE", "MErr", STR$(IBErr))
End Sub
Sub IBMsg(ByVal Msg As String, Optional ByVal Coul As Long)
On Error Resume Next

If Coul = BrightRed And IBErr Then
frmServer.ctlSysTrayBalloon.BalloonTipShow _
        "FRoG Server", _
        Msg, _
        NIIF_ERROR, _
        8000
ElseIf Coul = 1 And (IBAdmin Or IBJoueur Or IBTout) Then
frmServer.ctlSysTrayBalloon.BalloonTipShow _
        "FRoG Server", _
        Msg, _
        NIIF_INFO, _
        8000

Else
frmServer.ctlSysTrayBalloon.BalloonTipShow _
        "FRoG Server", _
        Msg, _
        NIIF_INFO, _
        8000
End If
End Sub



