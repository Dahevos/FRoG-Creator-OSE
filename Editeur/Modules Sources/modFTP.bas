Attribute VB_Name = "modFTP"
Const FTP_TRANSFER_TYPE_UNKNOWN As Long = &H0
Const FTP_TRANSFER_TYPE_ASCII As Long = &H1
Const FTP_TRANSFER_TYPE_BINARY As Long = &H2
Const INTERNET_DEFAULT_FTP_PORT As Byte = 21
Const INTERNET_SERVICE_FTP As Byte = 1
Const INTERNET_FLAG_PASSIVE As Long = &H8000000
Const INTERNET_OPEN_TYPE_PRECONFIG As Byte = 0
Const INTERNET_OPEN_TYPE_DIRECT As Byte = 1
Const INTERNET_OPEN_TYPE_PROXY As Byte = 3
Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY As Byte = 4
Const MAX_PATH As Integer = 260

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Long
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hConnect As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Long, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Boolean
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hConnect As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
Const PassiveConnection As Boolean = True
Private hOpen As Long

Public Sub Envoi(FTP As String, USER As String, PASS As String, Fichier As String, FichierFTP As String, PATHS As String)
    Dim hConnection As Long, hOpen As Long, sOrgpaths  As String

    hOpen = InternetOpen("FrogCreator", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

    hConnection = InternetConnect(hOpen, FTP, INTERNET_DEFAULT_FTP_PORT, USER, PASS, INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
    DoEvents

    sOrgpaths = String(MAX_PATH, 0)

    FtpGetCurrentDirectory hConnection, sOrgpaths, Len(sOrgpaths)
    DoEvents
 
    FtpSetCurrentDirectory hConnection, PATHS

    FtpPutFile hConnection, App.Path & "\" & Fichier, FichierFTP, FTP_TRANSFER_TYPE_UNKNOWN, 0
    DoEvents

    FtpGetFile hConnection, App.Path & "\" & Fichier, FichierFTP, False, 0, FTP_TRANSFER_TYPE_UNKNOWN, 0

    FtpSetCurrentDirectory hConnection, sOrgpaths
    DoEvents

    InternetCloseHandle hConnection

    InternetCloseHandle hOpen
    DoEvents
End Sub

Public Sub Supprimer(FTP As String, USER As String, PASS As String, FichierFTP As String, PATHS As String)
    Dim hConnection As Long, hOpen As Long, sOrgpaths  As String

    hOpen = InternetOpen("FrogCreator", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    DoEvents

    hConnection = InternetConnect(hOpen, FTP, INTERNET_DEFAULT_FTP_PORT, USER, PASS, INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
    
    sOrgpaths = String(MAX_PATH, 0)
    DoEvents

    FtpGetCurrentDirectory hConnection, sOrgpaths, Len(sOrgpaths)
     
    FtpSetCurrentDirectory hConnection, PATHS
    DoEvents

    FtpDeleteFile hConnection, FichierFTP
        
    FtpSetCurrentDirectory hConnection, sOrgpaths
    DoEvents

    InternetCloseHandle hConnection
  
    InternetCloseHandle hOpen
    DoEvents

End Sub

Public Sub Telecharger(FTP As String, USER As String, PASS As String, Fichier As String, FichierFTP As String, PATHS As String)
    Dim hConnection As Long, hOpen As Long, sOrgpaths  As String

    hOpen = InternetOpen("FrogCreator", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    DoEvents

    hConnection = InternetConnect(hOpen, FTP, INTERNET_DEFAULT_FTP_PORT, USER, PASS, INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
  
    sOrgpaths = String(MAX_PATH, 0)
    DoEvents

    FtpGetCurrentDirectory hConnection, sOrgpaths, Len(sOrgpaths)
 
    FtpSetCurrentDirectory hConnection, PATHS
    DoEvents

    FtpGetFile hConnection, FichierFTP, Fichier, True, 0, FTP_TRANSFER_TYPE_UNKNOWN, 0

    FtpSetCurrentDirectory hConnection, sOrgpaths
    DoEvents

    InternetCloseHandle hConnection
    
    InternetCloseHandle hOpen
    DoEvents

End Sub

Public Sub TestConection(FTP As String, USER As String, PASS As String, PATHS As String)
Dim hConnection As Long, hOpen As Long, sOrgpaths  As String
Dim Connecter As Boolean

    hOpen = InternetOpen("FrogCreator", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

    FrmOptFTP.bar.value = 10
    
    hConnection = InternetConnect(hOpen, FTP, INTERNET_DEFAULT_FTP_PORT, USER, PASS, INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
    DoEvents

    FrmOptFTP.bar.value = 30
    DoEvents

    sOrgpaths = String(MAX_PATH, 0)

    FrmOptFTP.bar.value = 50
    DoEvents
    
    Connecter = FtpGetCurrentDirectory(hConnection, sOrgpaths, Len(sOrgpaths))
    DoEvents
    
    FrmOptFTP.bar.value = 70
    DoEvents
    
    InternetCloseHandle hConnection
    
    FrmOptFTP.bar.value = 100
    InternetCloseHandle hOpen
    DoEvents
    If Connecter = True Then
        MsgBox "Le logiciel c'est connecter avec succès.", vbInformation
    Else
        MsgBox "Le logiciel n'arrive pas à ce connecter au serveur. Vérifiez vos informations, votre connections et si vous n'êtes pas déjà connecter au ftp avec un autre logiciel.", vbCritical
    End If
    FrmOptFTP.bar.Visible = False
End Sub

Public Function ConnexionFTP(FTP As String, USER As String, PASS As String) As Long
Dim Connecter As Boolean
Dim sOrgpaths  As String
    On Error Resume Next
    sOrgpaths = String(MAX_PATH, 0)
    
    hOpen = InternetOpen("FrogCreator", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

    ConnexionFTP = InternetConnect(hOpen, FTP, INTERNET_DEFAULT_FTP_PORT, USER, PASS, INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
    
    Connecter = FtpGetCurrentDirectory(ConnexionFTP, sOrgpaths, Len(sOrgpaths))
    DoEvents
    
    If Not Connecter Then
        MsgBox "Le logiciel n'arrive pas à ce connecter au serveur. Vérifiez vos informations, votre connections et si vous n'êtes pas déjà connecter au ftp avec un autre logiciel.", vbCritical
        Call FermerFTP(ConnexionFTP)
        ConnexionFTP = 0
    End If
End Function
    
Public Sub EnvoiFTP(hConnection As Long, FTP As String, USER As String, PASS As String, Fichier As String, FichierFTP As String, PATHS As String)
    On Error Resume Next
    Dim sOrgpaths  As String
    
    sOrgpaths = String(MAX_PATH, 0)

    FtpGetCurrentDirectory hConnection, sOrgpaths, Len(sOrgpaths)
    DoEvents
 
    FtpSetCurrentDirectory hConnection, PATHS
    DoEvents

    FtpPutFile hConnection, App.Path & "\" & Fichier, FichierFTP, FTP_TRANSFER_TYPE_UNKNOWN, 0
    DoEvents

    FtpGetFile hConnection, App.Path & "\" & Fichier, FichierFTP, False, 0, FTP_TRANSFER_TYPE_UNKNOWN, 0
    DoEvents

    FtpSetCurrentDirectory hConnection, sOrgpaths
    DoEvents
End Sub

Public Sub FermerFTP(hConnection As Long)
    On Error Resume Next
    InternetCloseHandle hConnection

    InternetCloseHandle hOpen
    DoEvents
End Sub
