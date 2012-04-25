Attribute VB_Name = "ModINI"
Option Explicit
Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal FileName$)

Public Sub WriteINI(INISection As String, INIKey As String, INIValue As String, INIFile As String)
    Call WritePrivateProfileString(INISection, INIKey, INIValue, INIFile)
End Sub

Public Function ReadINI(INISection As String, INIKey As String, INIFile As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = vbNullString
  
    sSpaces = Space$(5000)
  
    Call GetPrivateProfileString$(INISection, INIKey, szReturn, sSpaces, Len(sSpaces), INIFile)
  
    ReadINI = RTrim$(sSpaces)
    ReadINI = Left$(ReadINI, Len(ReadINI) - 1)
End Function

Public Sub InitAccountOpt()
On Error Resume Next
    AccOpt.InfName = ReadINI("INFO", "Account", App.Path & "\Config\Account.ini")
    AccOpt.InfPass = ReadINI("INFO", "Password", App.Path & "\Config\Account.ini")
    AccOpt.SpeechBubbles = CBool(ReadINI("CONFIG", "SpeechBubbles", App.Path & "\Config\Account.ini"))
    AccOpt.NpcBar = CBool(ReadINI("CONFIG", "NpcBar", App.Path & "\Config\Account.ini"))
    AccOpt.NpcName = CBool(ReadINI("CONFIG", "NPCName", App.Path & "\Config\Account.ini"))
    AccOpt.NpcDamage = CBool(ReadINI("CONFIG", "NPCDamage", App.Path & "\Config\Account.ini"))
    AccOpt.PlayBar = CBool(ReadINI("CONFIG", "PlayerBar", App.Path & "\Config\Account.ini"))
    AccOpt.PlayName = CBool(ReadINI("CONFIG", "PlayerName", App.Path & "\Config\Account.ini"))
    AccOpt.PlayDamage = CBool(ReadINI("CONFIG", "PlayerDamage", App.Path & "\Config\Account.ini"))
    AccOpt.MapGrid = CBool(ReadINI("CONFIG", "MapGrid", App.Path & "\Config\Account.ini"))
    AccOpt.Music = CBool(ReadINI("CONFIG", "Music", App.Path & "\Config\Account.ini"))
    AccOpt.Sound = CBool(ReadINI("CONFIG", "Sound", App.Path & "\Config\Account.ini"))
    AccOpt.Autoscroll = CBool(ReadINI("CONFIG", "AutoScroll", App.Path & "\Config\Account.ini"))
    AccOpt.NomObjet = CBool(ReadINI("CONFIG", "NomObjet", App.Path & "\Config\Account.ini"))
    AccOpt.LowEffect = CBool(ReadINI("CONFIG", "LowEffect", App.Path & "\Config\Account.ini"))
    DISPLAY_BUBBLE_TIME = ReadINI("CONFIG", "bubbletime", App.Path & "\Config\Client.ini")
    If (Not IsNumeric(DISPLAY_BUBBLE_TIME)) And (Not DISPLAY_BUBBLE_TIME > 0) Then
    DISPLAY_BUBBLE_TIME = 4000
    End If
    frmMirage.txtTempsBulles.Text = DISPLAY_BUBBLE_TIME
End Sub

