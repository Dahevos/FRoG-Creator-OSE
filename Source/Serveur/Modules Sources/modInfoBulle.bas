Attribute VB_Name = "modInfoBulle"
Option Explicit

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOPMOST = -1
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Tra As Byte
Public IBVisible As Boolean
Public IBCharge As Boolean
Public IBTout As Boolean
Public IBJoueur As Boolean
Public IBAdmin As Boolean
Public IBErr As Boolean
Public IBCJoueur As Long
Public IBCAdmin As Long

Public Function TransRegion(frm As Form, TranslucenceLevel As Byte, Crk As Long) As Boolean

SetWindowLong frm.hWnd, GWL_EXSTYLE, WS_EX_LAYERED
SetLayeredWindowAttributes frm.hWnd, Crk, TranslucenceLevel, &H3
TransRegion = Err.LastDllError = 0

End Function

Public Sub IBMsg(ByVal Msg As String, Optional ByVal Coul As Long, Optional ByVal Gras As Boolean)
Dim i As Long

On Error Resume Next
If Not IBJoueur And Coul = IBCJoueur And Not IBTout Then Exit Sub
If Not IBAdmin And Coul = IBCAdmin And Not IBTout Then Exit Sub

Call AddLog(Msg, "Logs\InfoBulle.txt")

For i = 1 To 10
    If IBMsgs(i).Texte = vbNullString Then
        If Coul <= 0 Then Coul = Black
        If Gras <> False And Gras <> True Then Gras = False
        IBMsgs(i).Texte = Msg
        IBMsgs(i).Couleur = Coul
        IBMsgs(i).Gra = Gras
        Exit For
    End If
Next i

End Sub

Sub AffIBMsg(ByVal Index As Long, ByVal Msg As String, Optional ByVal Coul As Long, Optional ByVal Gras As Boolean)
On Error Resume Next
Dim Te As String
Dim i As Long

Call AddLog(Msg, "Logs\InfoBulle.txt")

If Index > 0 And Index < 10 Then
    IBMsgs(Index).Texte = vbNullString
    IBMsgs(Index).Couleur = 0
    IBMsgs(Index).Gra = False
End If

Te = Msg
If Len(Te) > 37 Then
    For i = 0 To ((Len(Te) \ 37))
        If i > 0 Then Msg = Msg & vbCrLf & Mid$(Te, (37 * i) + 1, 37) Else Msg = Mid$(Te, 1, 37)
    Next i
End If

frmInfoBulle.Timer1.Enabled = False
frmInfoBulle.Timer2.Enabled = False
frmInfoBulle.Timer3.Enabled = False

frmInfoBulle.msgs.Caption = Msg
If Coul > 15 Then frmInfoBulle.msgs.ForeColor = Coul Else frmInfoBulle.msgs.ForeColor = QBColor(Coul)
frmInfoBulle.msgs.FontBold = Gras
Tra = 0
frmInfoBulle.Timer1.Enabled = True
frmInfoBulle.Timer2.Enabled = True
frmInfoBulle.Visible = True
IBCharge = True

End Sub

Sub CheckIBMsg()
Dim i As Long

If IBVisible Or IBCharge Then Exit Sub

For i = 1 To 10
    If IBMsgs(i).Texte <> vbNullString Then
        If IBCharge Or IBVisible Then Exit Sub
        IBCharge = True
        Call AffIBMsg(i, IBMsgs(i).Texte, IBMsgs(i).Couleur, IBMsgs(i).Gra)
    End If
Next i

End Sub

Sub VideIBMsg()
Dim i As Long

For i = 1 To 10
    IBMsgs(i).Texte = vbNullString
    IBMsgs(i).Couleur = Black
    IBMsgs(i).Gra = False
Next i

End Sub

Sub SauvIBOpt()
Call PutVar(App.Path & "\Options.ini", "INFOBULLE", "MTout", STR$(IBTout))
Call PutVar(App.Path & "\Options.ini", "INFOBULLE", "MJoueur", STR$(IBJoueur))
Call PutVar(App.Path & "\Options.ini", "INFOBULLE", "MAdmin", STR$(IBAdmin))
Call PutVar(App.Path & "\Options.ini", "INFOBULLE", "MErr", STR$(IBErr))
Call PutVar(App.Path & "\Options.ini", "INFOBULLE", "CJoueur", STR$(IBCJoueur))
Call PutVar(App.Path & "\Options.ini", "INFOBULLE", "CAdmin", STR$(IBCAdmin))
End Sub

Sub ChargIBOpt()
If Not FileExist("\Options.ini") Then Exit Sub
IBJoueur = CBool(GetVar(App.Path & "\Options.ini", "INFOBULLE", "MJoueur"))
IBAdmin = CBool(GetVar(App.Path & "\Options.ini", "INFOBULLE", "MAdmin"))
IBErr = CBool(GetVar(App.Path & "\Options.ini", "INFOBULLE", "MErr"))
IBTout = CBool(GetVar(App.Path & "\Options.ini", "INFOBULLE", "MTout"))
IBCJoueur = CLng(GetVar(App.Path & "\Options.ini", "INFOBULLE", "CJoueur"))
IBCAdmin = CLng(GetVar(App.Path & "\Options.ini", "INFOBULLE", "CAdmin"))
If IBJoueur Then frmOptInfoBulle.mj.value = Checked Else frmOptInfoBulle.mj.value = Unchecked
If IBErr Then frmOptInfoBulle.mer.value = Checked Else frmOptInfoBulle.mer.value = Unchecked
If IBAdmin Then frmOptInfoBulle.ma.value = Checked Else frmOptInfoBulle.ma.value = Unchecked
If IBTout Then frmOptInfoBulle.mt.value = Checked Else frmOptInfoBulle.mt.value = Unchecked
frmOptInfoBulle.jcoul.BackColor = IBCJoueur
frmOptInfoBulle.acoul.BackColor = IBCAdmin
End Sub

