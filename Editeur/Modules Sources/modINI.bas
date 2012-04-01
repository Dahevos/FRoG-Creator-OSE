Attribute VB_Name = "ModINI"
Option Explicit
Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal filename$)
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal filename$)

Public Sub WriteINI(INISection As String, INIKey As String, INIValue As String, INIFile As String)
    Call WritePrivateProfileString(INISection, INIKey, INIValue, INIFile)
End Sub

Public Function ReadINI(INISection As String, INIKey As String, INIFile As String) As String
    Dim StringBuffer As String
    Dim StringBufferSize As Long
    
    StringBuffer = Space$(255)
    StringBufferSize = Len(StringBuffer)
    
    StringBufferSize = GetPrivateProfileString(INISection, INIKey, vbNullString, StringBuffer, StringBufferSize, INIFile)
    
    If StringBufferSize > 0 Then ReadINI = Left$(StringBuffer, StringBufferSize) Else ReadINI = vbNullString
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
    AccOpt.CPreVisu = CBool(ReadINI("CONFIG", "PreVisu", App.Path & "\Config\Account.ini"))
    AccOpt.LowEffect = CBool(ReadINI("CONFIG", "LowEffect", App.Path & "\Config\Account.ini"))
End Sub
